import argparse
import json
import sys
import traceback
from pathlib import Path
from zipfile import BadZipFile, ZipFile
from xml.etree import ElementTree


def emit(kind, message):
    print(json.dumps({"type": kind, "message": message}, ensure_ascii=False), flush=True)


def validate_docx(filename):
    try:
        with ZipFile(filename) as archive:
            entries = archive.infolist()
            if len(entries) > 10000 or sum(entry.file_size for entry in entries) > 512 * 1024 * 1024:
                raise ValueError("文档解压后过大，请精简图片和附件后重试。")
            names = set(archive.namelist())
            required = {"[Content_Types].xml", "_rels/.rels", "word/document.xml"}
            if not required.issubset(names) or any(entry.flag_bits & 1 for entry in entries):
                raise ValueError("请选择未加密的有效 DOCX 文档。")
            content_types = ElementTree.fromstring(archive.read("[Content_Types].xml"))
            expected = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
            if not any(item.get("PartName") == "/word/document.xml" and item.get("ContentType") == expected for item in content_types):
                raise ValueError("文件内容不是受支持的 DOCX 文档。")
    except (BadZipFile, ElementTree.ParseError, KeyError) as error:
        raise ValueError("文档损坏或不是有效的 DOCX 文件，请使用 Word 重新保存后重试。") from error


def main():
    parser = argparse.ArgumentParser()
    for name in ("checker-root", "template", "document", "output", "report-json", "cache-dir"):
        parser.add_argument("--" + name, type=Path, required=True)
    args = parser.parse_args()
    job = None
    try:
        emit("stage", "正在验证两份 DOCX 文档")
        validate_docx(args.template)
        validate_docx(args.document)
        sys.path.insert(0, str(args.checker_root / "src"))
        import win32api
        import win32con
        import win32job
        import win32process
        from thesis_format_checker.core.word_session import WordSession
        from thesis_format_checker.engine.pipeline import compare_documents
        from thesis_format_checker.engine.reporting import write_json_report

        job = win32job.CreateJobObject(None, "")
        limits = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
        limits["BasicLimitInformation"]["LimitFlags"] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, limits)
        original_enter = WordSession.__enter__
        original_exit = WordSession.__exit__
        word_errors = []

        def enter_word(session):
            try:
                result = original_enter(session)
                probe = session.app.Documents.Add(Visible=False)
                try:
                    thread_id, process_id = win32process.GetWindowThreadProcessId(probe.Windows.Item(1).Hwnd)
                    process = win32api.OpenProcess(win32con.PROCESS_SET_QUOTA | win32con.PROCESS_TERMINATE, False, process_id)
                    try:
                        win32job.AssignProcessToJobObject(job, process)
                    finally:
                        win32api.CloseHandle(process)
                finally:
                    probe.Close(SaveChanges=False)
            except Exception as error:
                word_errors.append(error)
                traceback.print_exc(file=sys.stderr)
                original_exit(session, None, None, None)
                raise RuntimeError("无法启动并隔离 Word 检测进程") from error
            emit("stage", "正在读取 Word 格式与分页信息")
            return result

        def exit_word(session, *exception):
            try:
                return original_exit(session, *exception)
            finally:
                emit("stage", "正在整理格式差异并写入批注")

        WordSession.__enter__ = enter_word
        WordSession.__exit__ = exit_word
        emit("stage", "正在对照模板检查论文格式")
        summary = compare_documents(args.template, args.document, args.output, cache_dir=args.cache_dir, show_comments=True)
        if word_errors:
            raise RuntimeError("Word 检测进程未能正常运行") from word_errors[0]
        emit("stage", "正在保存批注文档与检测报告")
        validate_docx(args.output)
        write_json_report(summary, args.report_json)
        return 0
    except Exception as error:
        traceback.print_exc(file=sys.stderr)
        message = str(error) if isinstance(error, ValueError) else "格式检测未能完成，请确认 Word 可正常使用，且文档未损坏或设置密码。"
        emit("error", message)
        return 3
    finally:
        if job is not None:
            job.Close()


if __name__ == "__main__":
    raise SystemExit(main())
