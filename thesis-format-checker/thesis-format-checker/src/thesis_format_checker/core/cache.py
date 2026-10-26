"""检测结果缓存：同一份模板/论文重复检测时复用已解析的基准与段落属性。

为什么需要缓存
==============

一次完整检测最贵的部分是 Microsoft Word（COM）：模板有效格式解析、逐段段落属性
读取与分页统计合计要花几分钟；而同一篇论文往往要被反复检测（改一处再跑一次）。
本模块把**完全由输入决定**的中间结果缓存到磁盘：

* ``word-flow``：旧版 Word 流程的结论（模板基准、段落属性比对结果、分页事实）。
  命中时整趟 Word 流程都被跳过，二次检测只剩 docx 侧的确定性与批注写入；
* ``layout``：排版几何快照（题注/图片/表格坐标与按页锚点）。命中时不再打开 Word
  读坐标，页眉页脚批注照样能落到对应页的版心顶部/底部。

缓存键的组成
============

``缓存结构版本 + 源码指纹 + 输入文件内容指纹 + 影响结果的可选参数``

* 输入文件改一个字节 → 内容指纹变化 → 该文件相关的缓存自动失效；
* 任何源码文件变化 → 源码指纹变化 → 全部缓存自动失效（开发期间不会读到旧结论）；
* 因此不需要使用者手工清理缓存，也不存在“改了模板/代码还命中旧结果”的风险。

安全与边界
==========

* 缓存只保存中间结果，**不保存也不修改输入文档**，输出文档仍由原文件 + 批注生成；
* 写盘采用“临时文件 + 原子替换”，读坏的缓存文件会被当作未命中并删除；
* 缓存失败（磁盘满、无权限）只是少一次加速，绝不中断检测；
* ``THESIS_CHECKER_NO_CACHE=1`` 或命令行 ``--no-cache`` 可完全关闭缓存。
"""

from __future__ import annotations

import hashlib
import json
import os
from dataclasses import fields, is_dataclass
from enum import Enum
from pathlib import Path
from typing import Any

from thesis_format_checker.core.models import Finding, RunSummary, Severity, Status

#: 缓存结构版本。缓存内字段有增减时递增，旧缓存自动失效。
CACHE_SCHEMA_VERSION = "1"

#: 源码根目录（``.../thesis_format_checker``），用于计算源码指纹。
PACKAGE_ROOT = Path(__file__).resolve().parents[1]

#: 环境变量：指定缓存目录；不设置时使用系统本地应用数据目录。
ENV_CACHE_DIR = "THESIS_CHECKER_CACHE_DIR"

#: 环境变量：设为 1 时完全关闭缓存（等价于命令行 ``--no-cache``）。
ENV_NO_CACHE = "THESIS_CHECKER_NO_CACHE"

#: 源码指纹在本进程内只算一次（一次检测里会被多次用来拼缓存键）。
_CODE_FINGERPRINT: str | None = None


def default_cache_dir() -> Path:
    """默认缓存目录：``%LOCALAPPDATA%\\thesis-checker\\cache``（可用环境变量覆盖）。"""

    override = os.environ.get(ENV_CACHE_DIR)
    if override:
        return Path(override)
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("XDG_CACHE_HOME")
    if base:
        return Path(base) / "thesis-checker" / "cache"
    return Path.home() / ".cache" / "thesis-checker"


def code_fingerprint() -> str:
    """源码指纹：所有 ``*.py`` 的“相对路径 + 大小 + 修改时间”的哈希。

    只做文件状态读取（不读文件内容），几十个文件耗时毫秒级；任何源码改动都会
    改变修改时间，从而让缓存自动失效。
    """

    global _CODE_FINGERPRINT
    if _CODE_FINGERPRINT is not None:
        return _CODE_FINGERPRINT
    digest = hashlib.sha256()
    for path in sorted(PACKAGE_ROOT.rglob("*.py")):
        try:
            stat = path.stat()
        except OSError:
            continue
        digest.update(str(path.relative_to(PACKAGE_ROOT)).encode("utf-8", "surrogatepass"))
        digest.update(f"|{stat.st_size}|{stat.st_mtime_ns}\n".encode("ascii"))
    _CODE_FINGERPRINT = digest.hexdigest()[:16]
    return _CODE_FINGERPRINT


def file_fingerprint(path) -> str:
    """文件内容指纹（SHA-256 前 32 位十六进制）。文件不存在时返回空串。"""

    digest = hashlib.sha256()
    try:
        with open(path, "rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError:
        return ""
    return digest.hexdigest()[:32]


def clean_params(value):
    """把影响结论的参数整理成可 JSON 化、键顺序稳定的形式。

    映射配置里常出现整数键（节号），直接 ``json.dumps(sort_keys=True)`` 会因为
    键类型不一致而报错、或让同样的配置算出不同的键。这里统一转成字符串键。
    """

    if isinstance(value, dict):
        return {str(key): clean_params(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [clean_params(item) for item in value]
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


# ---------------------------------------------------------------------------
# 序列化：dataclass / 枚举 / 任意键字典 → JSON，再还原
# ---------------------------------------------------------------------------

_DATACLASS_REGISTRY: dict[str, type] | None = None


def _registry() -> dict[str, type]:
    """可还原的 dataclass 类型表（延迟构建，避免模块导入环）。"""

    global _DATACLASS_REGISTRY
    if _DATACLASS_REGISTRY is None:
        from thesis_format_checker.core import models, word_layout

        registry: dict[str, type] = {}
        for module in (models, word_layout):
            for name in dir(module):
                value = getattr(module, name)
                if isinstance(value, type) and is_dataclass(value):
                    registry[name] = value
        _DATACLASS_REGISTRY = registry
    return _DATACLASS_REGISTRY


def encode(value: Any) -> Any:
    """把任意检测中间结果编码成 JSON 友好结构（保留类型与整数键）。"""

    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, Enum):
        return value.value
    if isinstance(value, Path):
        return str(value)
    if is_dataclass(value) and not isinstance(value, type):
        return {"__kind__": type(value).__name__,
                "__fields__": {item.name: encode(getattr(value, item.name))
                               for item in fields(value)}}
    if isinstance(value, dict):
        # 整数键（Word 段号 → 页码）必须原样还原，因此不直接用 JSON 对象的键。
        return {"__kind__": "dict",
                "__items__": [[encode(key), encode(item)] for key, item in value.items()]}
    if isinstance(value, (list, tuple)):
        return {"__kind__": "tuple" if isinstance(value, tuple) else "list",
                "__items__": [encode(item) for item in value]}
    if isinstance(value, (set, frozenset)):
        return {"__kind__": "set", "__items__": [encode(item) for item in sorted(value, key=str)]}
    return str(value)      # 兜底：不可识别的对象按文本保存，至少不丢内容


def decode(value: Any) -> Any:
    """:func:`encode` 的逆操作。"""

    if isinstance(value, list):
        # encode 之外的历史数据（纯 JSON 列表）也照常还原。
        return [decode(item) for item in value]
    if not isinstance(value, dict):
        return value
    kind = value.get("__kind__")
    if kind is None:
        return {key: decode(item) for key, item in value.items()}
    if kind == "dict":
        return {decode(key): decode(item) for key, item in value.get("__items__", [])}
    if kind in {"list", "tuple", "set"}:
        items = [decode(item) for item in value.get("__items__", [])]
        if kind == "tuple":
            return tuple(items)
        if kind == "set":
            return set(items)
        return items
    dataclass_type = _registry().get(kind)
    if dataclass_type is None:
        return {key: decode(item) for key, item in value.get("__fields__", {}).items()}
    return dataclass_type(**{key: decode(item)
                             for key, item in value.get("__fields__", {}).items()})


# ---------------------------------------------------------------------------
# 缓存条目：Word 流程结论与排版快照
# ---------------------------------------------------------------------------

def summary_payload(summary: RunSummary) -> dict:
    """Word 流程结论 → 可缓存的字典（只保存与输入相关、可复现的部分）。"""

    return {
        "sections_checked": summary.sections_checked,
        "body_paragraphs_checked": summary.body_paragraphs_checked,
        "findings": encode(list(summary.findings)),
        "unknowns": list(summary.unknowns),
        "render_facts": encode(dict(summary.render_facts or {})),
    }


def summary_from_payload(payload: dict) -> RunSummary:
    """缓存字典 → :class:`RunSummary`（``output_path`` 由调用方重新填写）。"""

    findings = [item for item in decode(payload.get("findings") or [])
                if isinstance(item, Finding)]
    for finding in findings:
        # 枚举经 JSON 往返后是字符串，这里统一还原成枚举，保证与现场检测一致。
        if not isinstance(finding.status, Status):
            finding.status = Status(finding.status)
        if finding.severity is not None and not isinstance(finding.severity, Severity):
            finding.severity = Severity(finding.severity)
        finding.merged_locations = list(finding.merged_locations or [])
    return RunSummary(
        sections_checked=int(payload.get("sections_checked") or 0),
        body_paragraphs_checked=int(payload.get("body_paragraphs_checked") or 0),
        findings=findings,
        unknowns=[str(item) for item in (payload.get("unknowns") or [])],
        render_facts=decode(payload.get("render_facts") or {}),
    )


class ResultCache:
    """磁盘缓存：按“类型 + 缓存键”保存/读取 JSON。"""

    def __init__(self, directory=None, enabled: bool = True) -> None:
        self.directory = Path(directory) if directory else default_cache_dir()
        if os.environ.get(ENV_NO_CACHE):
            enabled = False
        self.enabled = bool(enabled)
        #: 命中/未命中的条目名，用于在摘要里告诉使用者“这次为什么快/慢”。
        self.hits: list[str] = []
        self.misses: list[str] = []
        #: 缓存不可用时的说明（只提示，不影响检测）。
        self.notes: list[str] = []

    # -- 缓存键 ----------------------------------------------------------

    def key(self, kind: str, *, template=None, document=None, params=None) -> str:
        """计算缓存键：结构版本 + 源码指纹 + 输入内容指纹 + 参数。"""

        parts = [CACHE_SCHEMA_VERSION, code_fingerprint(), kind]
        if template is not None:
            parts.append("T:" + file_fingerprint(template))
        if document is not None:
            parts.append("D:" + file_fingerprint(document))
        if params is not None:
            parts.append("P:" + json.dumps(clean_params(params), sort_keys=True,
                                           ensure_ascii=False, default=str))
        return hashlib.sha256("|".join(parts).encode("utf-8")).hexdigest()[:32]

    # -- 读写 ------------------------------------------------------------

    def _path(self, kind: str, key: str) -> Path:
        return self.directory / kind / f"{key}.json"

    def load(self, kind: str, key: str, label: str = "") -> Any | None:
        """读缓存并还原成 Python 对象；未命中、损坏或缓存关闭时返回 ``None``。"""

        if not self.enabled:
            return None
        path = self._path(kind, key)
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except FileNotFoundError:
            self.misses.append(label or kind)
            return None
        except (OSError, ValueError):
            # 损坏的缓存按未命中处理并删除，绝不让它影响检测结果。
            try:
                path.unlink()
            except OSError:
                pass
            self.notes.append(f"缓存文件损坏，已忽略：{path.name}")
            self.misses.append(label or kind)
            return None
        if not isinstance(payload, dict) or payload.get("schema") != CACHE_SCHEMA_VERSION:
            self.misses.append(label or kind)
            return None
        self.hits.append(label or kind)
        return decode(payload.get("data"))

    def store(self, kind: str, key: str, data: Any, label: str = "") -> bool:
        """写缓存（原子替换，自动序列化 dataclass/整数键字典）；失败只记说明。"""

        if not self.enabled:
            return False
        path = self._path(kind, key)
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            temporary = path.with_suffix(".json.tmp")
            temporary.write_text(
                json.dumps({"schema": CACHE_SCHEMA_VERSION, "data": encode(data)},
                           ensure_ascii=False),
                encoding="utf-8")
            os.replace(temporary, path)
        except (OSError, TypeError, ValueError) as exc:
            self.notes.append(f"缓存写入失败（{label or kind}）：{type(exc).__name__}")
            return False
        return True

    # -- 摘要 ------------------------------------------------------------

    def stats(self) -> dict:
        """给 :class:`RunSummary` 用的缓存统计。"""

        return {"enabled": self.enabled, "directory": str(self.directory),
                "hits": list(self.hits), "misses": list(self.misses),
                "notes": list(self.notes)}


__all__ = [
    "CACHE_SCHEMA_VERSION",
    "ENV_CACHE_DIR",
    "ENV_NO_CACHE",
    "ResultCache",
    "code_fingerprint",
    "decode",
    "default_cache_dir",
    "encode",
    "file_fingerprint",
    "summary_from_payload",
    "summary_payload",
]
