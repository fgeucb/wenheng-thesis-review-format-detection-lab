from __future__ import annotations

import os
from pathlib import Path


class WordAutomationError(RuntimeError):
    pass


class WordSession:
    """拥有独立 Word 进程；只关闭本类创建的文档和实例。"""

    def __init__(self) -> None:
        self.app = None
        self._documents: list[object] = []

    def __enter__(self) -> "WordSession":
        try:
            import pythoncom
            from win32com.client import DispatchEx

            pythoncom.CoInitialize()
            self._pythoncom = pythoncom
            self.app = DispatchEx("Word.Application")
            self.app.Visible = False
            self.app.DisplayAlerts = 0
            try:
                self.app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:
                pass
            self.tune_for_detection()
            return self
        except Exception as exc:
            self._uninitialize()
            raise WordAutomationError(
                "无法启动 Microsoft Word。请确认 Windows 已安装桌面版 Word，且 pywin32 安装正常。"
            ) from exc

    # -- 性能开关（只影响本次自动化实例，退出即释放） --------------------

    def tune_for_detection(self) -> None:
        """关闭与检测无关的 Word 后台任务。

        屏幕刷新、拼写/语法检查、自动断字等都会在每次 COM 访问时被触发，是
        "读一个属性也要几十毫秒"的主要来源之一。这里逐项关闭，并把后台重排
        （``Options.Pagination``）关掉——本阶段只读格式，不需要页码；
        读页码前请调用 :meth:`resume_pagination` 恢复。
        """

        self._set_option("ScreenUpdating", False)
        self._set_option("CheckSpellingAsYouType", False)
        self._set_option("CheckGrammarAsYouType", False)
        self._set_option("CheckGrammarWithSpelling", False)
        self.pause_pagination()

    def _set_option(self, name: str, value) -> None:
        """设置 ``Application.Options.<name>``；不存在或不可写时忽略。"""

        if self.app is None:
            return
        try:
            setattr(self.app.Options, name, value)
        except Exception:
            pass

    def pause_pagination(self) -> None:
        """关闭后台重排：读格式阶段显著更快，但此时 ``Information(3)`` 不可用。"""

        self._set_option("Pagination", False)

    def resume_pagination(self) -> None:
        """恢复后台重排并重新分页：之后才能读页码/节号。"""

        self._set_option("Pagination", True)
        for document in self._documents[:]:
            try:
                document.Repaginate()
            except Exception:
                pass

    def open_readonly(self, path: Path):
        self._validate_input(path)
        try:
            document = self.app.Documents.Open(
                FileName=str(path),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Revert=False,
                NoEncodingDialog=True,
            )
            self._documents.append(document)
            return document
        except Exception as exc:
            raise WordAutomationError(f"无法读取 Word 文档：{path}。文件可能损坏、被保护或格式不受支持。") from exc

    def create_output_copy(self, source_path: Path, output_path: Path, overwrite: bool = False):
        self._validate_input(source_path)
        if source_path.resolve() == output_path.resolve():
            raise WordAutomationError("输出路径不能与待检测文档相同，程序不会覆盖输入原件。")
        if output_path.suffix.casefold() != ".docx":
            raise WordAutomationError("输出文件扩展名必须是 .docx。")
        if output_path.exists() and not overwrite:
            raise WordAutomationError(f"输出文件已存在：{output_path}。如需覆盖，请显式使用 --overwrite。")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        if output_path.exists():
            output_path.unlink()
        source = None
        try:
            source = self.app.Documents.Open(
                FileName=str(source_path),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Revert=False,
                NoEncodingDialog=True,
            )
            self._documents.append(source)
            source.SaveAs2(FileName=str(output_path), FileFormat=16, AddToRecentFiles=False)
            source.Close(SaveChanges=False)
            self._documents.remove(source)
            source = None
            output = self.app.Documents.Open(
                FileName=str(output_path),
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
                Revert=False,
                NoEncodingDialog=True,
            )
            self._documents.append(output)
            if bool(output.ReadOnly):
                raise WordAutomationError("输出副本以只读方式打开，无法写入批注。")
            return output
        except WordAutomationError:
            raise
        except Exception as exc:
            raise WordAutomationError(f"无法创建输出副本：{output_path}") from exc

    def close_document(self, document, save: bool = False) -> None:
        try:
            document.Close(SaveChanges=bool(save))
        finally:
            if document in self._documents:
                self._documents.remove(document)

    def __exit__(self, exc_type, exc, traceback) -> None:
        for document in reversed(self._documents[:]):
            try:
                document.Close(SaveChanges=False)
            except Exception:
                pass
        self._documents.clear()
        if self.app is not None:
            try:
                self.app.Quit(SaveChanges=False)
            except Exception:
                pass
            self.app = None
        self._uninitialize()

    @staticmethod
    def _validate_input(path: Path) -> None:
        if path.suffix.casefold() not in {".doc", ".docx"}:
            raise WordAutomationError(f"仅支持 .doc 或 .docx 输入：{path}")
        if not path.is_file():
            raise WordAutomationError(f"输入文件不存在：{path}")

    def _uninitialize(self) -> None:
        pythoncom = getattr(self, "_pythoncom", None)
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._pythoncom = None

