"""检测规则目录（Rule Catalog）：大类、规则族、细粒度规则与等级的唯一权威清单。

三层结构
========

* :class:`CategorySpec` —— **一个大类 = 一个类**，对应 ``checks/cXX_*.py``；
* :class:`RuleSpec`     —— **规则族**，沿用 001–165 编号，是“归类与追溯”单位，
  也是今天检测代码的输出单位（方法名 ``check_<族编号>``）；
* :class:`RuleItem`     —— **细粒度规则（一条只讲一件事）**，编号为
  ``<族编号>.<序号>``（被拆分的族）或直接等于族编号（本身只讲一件事）。

每条细粒度规则都带：标题、等级（错误/警告/提醒）、状态（已实现/聚合输出/待实现/
待确认）、定位方式。清单数据在 :mod:`thesis_format_checker.checks.rule_specs`，
本模块负责把它和取证实现（:class:`ProviderSpec`）组合成可查询的目录。

依赖大模型语义判断的原始规则已整体删除，删除清单与原因见
:data:`REMOVED_FAMILIES`（``rule_specs``）与 ``docs/规则清单.md``。

维护者只要打开目录就能知道：某条细粒度规则属于哪个大类、等级多高、是否已实现、
由哪段取证产出、结论定位到哪里；``thesis-checker rules`` 直接打印这份目录。
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

from thesis_format_checker.checks.rule_specs import (
    FINE_RULE_ROWS,
    REMOVED_FAMILIES,
    family_of,
)
from thesis_format_checker.core.models import SEVERITY_ORDER, Severity


class RuleStatus(str, Enum):
    """实现状态。

    细粒度规则（:class:`RuleItem`）使用 ``IMPLEMENTED``/``AGGREGATED``/``PENDING``/
    ``DEFERRED``；规则族（:class:`RuleSpec`）在此基础上使用 ``LEGACY`` 与 ``REMOVED``。
    """

    IMPLEMENTED = "implemented"  # 已实现，且该条有独立结论
    AGGREGATED = "aggregated"    # 已实现，但与同族其他条目合并为一条批注（待拆分）
    LEGACY = "legacy"            # 已实现（旧版 Word 有效格式流程）
    PENDING = "pending"          # 待实现：可以确定性检测，尚未开发
    DEFERRED = "deferred"        # 待确认：业务口径先确认
    REMOVED = "removed"          # 已删除：必须依赖大模型语义


STATUS_LABELS: dict[RuleStatus, str] = {
    RuleStatus.IMPLEMENTED: "已实现",
    RuleStatus.AGGREGATED: "已实现（聚合输出，待拆分）",
    RuleStatus.LEGACY: "已实现（旧版 Word 有效格式检测）",
    RuleStatus.PENDING: "待实现（可确定性检测）",
    RuleStatus.DEFERRED: "待确认（业务口径）",
    RuleStatus.REMOVED: "已删除（依赖大模型语义）",
}


@dataclass(frozen=True)
class CategorySpec:
    """一个大类的元数据。

    Attributes:
        key: 稳定标识，例如 ``"c07"``，同时是 ``CheckResult.for_category`` 的入参。
        index: 1 开始的中文序号，与《格式检测错误总结》一致。
        name: 中文大类名称。
        module: 实现该大类的模块名（``checks`` 包内），便于文档与自检定位。
        checker: 该大类检测器类名。
        summary: 一句话说明这一类在检查什么。
    """

    key: str
    index: int
    name: str
    module: str
    checker: str
    summary: str

    @property
    def first_rule(self) -> str:
        """本大类第一条规则的编号（规则编号连续，每类 11 条）。"""

        return f"{11 * (self.index - 1) + 1:03d}"

    @property
    def last_rule(self) -> str:
        """本大类最后一条规则的编号。"""

        return f"{11 * self.index:03d}"

    @property
    def rule_ids(self) -> tuple[str, ...]:
        """本大类包含的全部规则编号。"""

        start = 11 * (self.index - 1) + 1
        return tuple(f"{number:03d}" for number in range(start, start + 11))


@dataclass(frozen=True)
class RuleSpec:
    """一条规则的元数据，对应大类检测器里的一个方法。"""

    rule_id: str
    title: str
    category: str
    status: RuleStatus
    providers: tuple[str, ...] = ()
    note: str = ""
    #: 规则族等级：取族内最严重的一条细粒度规则。
    severity: Severity = Severity.ERROR
    #: 该族的结论当前是否为聚合输出（同族多条细规则合并成一条批注）。
    aggregated: bool = False
    #: 该族包含的细粒度规则。
    items: tuple["RuleItem", ...] = ()

    @property
    def pending_items(self) -> tuple["RuleItem", ...]:
        """族内尚未实现的细粒度规则。"""

        return tuple(item for item in self.items
                     if item.status in (RuleStatus.PENDING, RuleStatus.DEFERRED))

    @property
    def method(self) -> str:
        """规则对应的方法名，永远为 ``check_<三位编号>``。"""

        return f"check_{self.rule_id}"

    @property
    def number(self) -> int:
        """规则编号的整数形式，便于排序和范围判断。"""

        return int(self.rule_id)


@dataclass(frozen=True)
class RuleItem:
    """细粒度规则：**一条只讲一件事情**。"""

    code: str        #: 编号，如 ``005.2``；未拆分的族直接使用族编号
    title: str       #: 这一件事是什么
    family: str      #: 所属规则族（``001``–``165``）
    category: str    #: 大类 key，如 ``c07``
    severity: Severity   #: 等级
    status: RuleStatus   #: 实现状态
    location: str = ""   #: 定位方式说明（用于“尽可能精准定位”）

    @property
    def method(self) -> str:
        """对应的检测方法名（在所属大类的检测器类上）。"""

        return f"check_{self.family}"


@dataclass(frozen=True)
class ProviderSpec:
    """一段“取证实现”的元数据（``detectors/`` 模块或旧版 Word 流程）。

    ``ProviderSpec`` 是重构前后行为一致的保证：引擎按 :data:`PROVIDERS`
    的顺序调用它们，得到的 :class:`~thesis_format_checker.core.models.Finding` 顺序与
    重构前完全相同，因此原生批注的顺序、文档内容都不受影响。
    """

    key: str
    label: str
    module: str
    namespace: str | None
    extract: str | None
    check: str
    rules: tuple[str, ...]
    extract_takes_mapping: bool = False
    check_takes_mapping: bool = False


# ---------------------------------------------------------------------------
# 15 个大类
# ---------------------------------------------------------------------------

CATEGORIES: tuple[CategorySpec, ...] = (
    CategorySpec("c01", 1, "结构完整性检测", "c01_structure", "StructureChecker",
                 "论文是否包含学校要求的组成部分，以及它们的先后顺序。"),
    CategorySpec("c02", 2, "封面与前置部分检测", "c02_front", "FrontMatterChecker",
                 "封面、声明、摘要、关键词等前置区域的字段与版式。"),
    CategorySpec("c03", 3, "页面设置检测", "c03_pages", "PageSetupChecker",
                 "纸张、方向、页边距、装订线、分页与裁切风险。"),
    CategorySpec("c04", 4, "页眉页脚与页码检测", "c04_headers", "HeaderFooterChecker",
                 "页眉页脚内容、格式、链接关系与页码编号。"),
    CategorySpec("c05", 5, "标题与编号检测", "c05_headings", "HeadingChecker",
                 "各级标题的格式、编号连续性与编号形式。"),
    CategorySpec("c06", 6, "目录检测", "c06_toc", "TocChecker",
                 "目录是否自动生成、条目与正文是否同步、目录格式。"),
    CategorySpec("c07", 7, "字体与段落检测", "c07_paragraphs", "ParagraphFormatChecker",
                 "字体、字号、对齐、缩进、行距与段间距。"),
    CategorySpec("c08", 8, "图表检测", "c08_figures", "FigureTableChecker",
                 "图题表题位置与格式、编号序列、三线表与表格格式。"),
    CategorySpec("c09", 9, "公式检测", "c09_equations", "EquationChecker",
                 "公式对齐、编号格式与序列、公式字体与间距。"),
    CategorySpec("c10", 10, "引用、注释与交叉引用检测", "c10_citations", "CitationNoteChecker",
                 "脚注尾注设置、交叉引用机制与域错误信息。"),
    CategorySpec("c11", 11, "参考文献检测", "c11_bibliography", "BibliographyChecker",
                 "参考文献著录结构、必填字段、标点与排列顺序。"),
    CategorySpec("c12", 12, "字符、符号、标点、数字与空格检测", "c12_characters", "CharacterChecker",
                 "全半角、中英文边界空格、空行与日期编号形式。"),
    CategorySpec("c13", 13, "一致性与关联性检测", "c13_consistency", "ConsistencyChecker",
                 "题目、目录、图表引用、同类对象格式的内部一致性。"),
    CategorySpec("c14", 14, "文字与语言规范检测", "c14_language", "LanguageChecker",
                 "错别字、语法、术语、缩略语与学术语言风格。"),
    CategorySpec("c15", 15, "附录及特殊内容检测", "c15_appendix", "AppendixChecker",
                 "附录编号与格式、特殊版式、成果条目与文档末尾残留。"),
)

CATEGORY_BY_KEY: dict[str, CategorySpec] = {item.key: item for item in CATEGORIES}


# ---------------------------------------------------------------------------
# 165 条规则的标题（与《格式检测错误总结》逐条对应）
# ---------------------------------------------------------------------------

RULE_TITLES: dict[str, str] = {
    # 一、结构完整性检测
    "001": "缺少学校要求的封面或题名页",
    "002": "缺少原创性声明、独创性声明或授权使用声明",
    "003": "缺少中文摘要或中文摘要内容不完整",
    "004": "缺少英文摘要或英文摘要与中文摘要不对应",
    "005": "缺少中文关键词、英文关键词或关键词数量不符合要求",
    "006": "缺少目录或目录不是由标题样式自动生成",
    "007": "正文缺少学校规定的必要章节",
    "008": "缺少参考文献部分或正文引用文献但文末无参考文献表",
    "009": "致谢、附录、攻读学位期间成果等必需部分缺失",
    "010": "各组成部分的排列顺序不符合学校模板要求",
    "011": "存在无内容的章节、重复章节或明显遗漏的章节",
    # 二、封面与前置部分检测
    "012": "封面上的论文题目缺失或与正文、题名页中的题目不一致",
    "013": "封面题目字体、字号、加粗方式或行距不符合模板要求",
    "014": "作者姓名、学号、学院、专业、班级等基本信息缺失或填写错误",
    "015": "指导教师姓名、职称或单位信息缺失或格式不统一",
    "016": "提交日期、答辩日期或学位授予日期缺失或日期格式错误",
    "017": "学校名称、校徽的位置、尺寸、颜色或清晰度不符合模板要求",
    "018": "封面元素的对齐方式、上下间距或左右位置与模板不一致",
    "019": "密级、分类号、UDC等要求填写的字段缺失或填写位置错误",
    "020": "原创性声明中的签名、日期或必填信息缺失",
    "021": "中文摘要与英文摘要的标题、正文或关键词排版格式不一致",
    "022": "前置部分错误显示页眉、页脚或正文使用的阿拉伯数字页码",
    # 三、页面设置检测
    "023": "纸张大小不是学校要求的A4或页面宽高设置错误",
    "024": "页面方向错误，将应为纵向的页面设置为横向或反之",
    "025": "上、下、左、右页边距不符合学校模板规定",
    "026": "装订线位置或装订线宽度设置错误或缺失",
    "027": "页眉距页面顶端、页脚距页面底端的距离不符合要求",
    "028": "同一章节或同一文档中存在不必要的不同页面尺寸或页边距",
    "029": "分节符使用错误导致页面设置、页码或页眉页脚异常变化",
    "030": "章节之间未按要求另起一页或误用大量空行代替分页符",
    "031": "奇偶页、首页不同等版式选项设置与模板要求不一致",
    "032": "文档中存在非必要空白页或应保留的空白页被删除",
    "033": "页面内容超出可打印区域，存在被裁切的风险",
    # 四、页眉页脚与页码检测
    "034": "应设置页眉的页面缺少页眉或不应有页眉的页面出现页眉",
    "035": "页眉文字内容与学校名称、论文题目或当前章节名称不一致",
    "036": "页眉字体、字号、对齐方式或下边框样式不符合要求",
    "037": "奇数页与偶数页的页眉内容或位置未按要求分别设置",
    "038": "同一节中的页眉未与前一节断开或错误链接到前一节",
    "039": "页脚包含多余文字、横线或与模板不一致的内容",
    "040": "页码缺失、重复、跳号或出现两个页码",
    "041": "前置部分未使用罗马数字或正文未使用阿拉伯数字编号",
    "042": "页码起始值错误，正文首页未按规定从1开始编号",
    "043": "页码位置、字体、字号或对齐方式不符合模板要求",
    "044": "封面、声明页等不应显示页码的页面显示了页码",
    # 五、标题与编号检测
    "045": "一级标题的字体、字号、加粗方式或对齐方式不符合要求",
    "046": "二级标题的字体、字号、缩进或段前段后间距不符合要求",
    "047": "三级及以下标题的格式与对应层级的模板样式不一致",
    "048": "标题层级跳级，例如一级标题后直接出现三级标题",
    "049": "同一层级标题使用了不同的字体、字号、缩进或编号形式",
    "050": "章节编号不连续，存在跳号、重号或错误排序",
    "051": "标题编号格式错误",
    "052": "标题编号与标题文字之间的空格、制表符或标点使用错误",
    "053": "无编号标题被错误编号或应编号标题缺少编号",
    "054": "标题位于页面末行，与其后的正文、图表或公式分离",
    "055": "使用手工输入编号而非多级列表导致编号无法自动更新",
    # 六、目录检测
    "056": "目录未自动生成，而是手工输入或手工修改",
    "057": "正文已有标题未出现在目录中或目录包含正文不存在的标题",
    "058": "目录中的标题文字、编号或层级与正文不一致",
    "059": "目录页码与正文实际页码不对应",
    "060": "目录层级显示范围不符合要求，层级过多或过少",
    "061": "目录各级条目的字体、字号、加粗或缩进格式错误",
    "062": "目录条目与页码之间缺少点引导线或引导线样式不一致",
    "063": "目录页码未右对齐或不同条目的页码位置不齐",
    "064": "目录标题的字体、字号、位置或段落间距不符合模板要求",
    "065": "目录中出现孤立条目、空白条目、重复条目或错误换行",
    "066": "修改正文后未更新目录导致标题或页码仍为旧内容",
    # 七、字体与段落检测
    "067": "正文中文字体不是学校规定的指定字体",
    "068": "正文英文及数字字体不是规定的指定字体",
    "069": "正文、标题、图表或注释的字号不符合对应样式要求",
    "070": "同一用途的文字出现字体、字号、颜色或加粗格式不一致",
    "071": "正文未按要求两端对齐或误设为居中、左对齐、右对齐",
    "072": "正文首行未缩进2个字符或使用空格、制表符模拟首行缩进",
    "073": "不应缩进的标题、图题、表题或参考文献被设置了首行缩进",
    "074": "行距不是学校规定的数值",
    "075": "段前、段后间距不符合要求或用空行代替段落间距",
    "076": "段落中出现异常悬挂缩进、左缩进或右缩进",
    "077": "孤行控制、段中不分页或与下段同页等分页设置使用不当",
    # 八、图表检测
    "078": "图题未置于图片下方或表题未置于表格上方",
    "079": "图题、表题的字体、字号、对齐方式或段落间距不符合要求",
    "080": "图表编号不连续，存在跳号、重号或章节号错误",
    "081": "图表编号格式错误，如多种写法无规则混用",
    "082": "正文提及的图表编号不存在或图表未在正文中被引用",
    "083": "表格未按要求采用三线表，存在多余竖线或横线",
    "084": "表头、表身、合计行的字体、对齐、行高或底纹格式不规范",
    "085": "图表超出页边距、尺寸过小、比例失真或跨页位置不合理",
    "086": "图片分辨率过低、文字模糊、背景不洁或颜色不适合打印",
    "087": "跨页表格未重复标题行或表格被不合理拆分",
    "088": "图注、表注、数据来源或单位标注缺失或其位置与格式错误",
    # 九、公式检测
    "089": "公式未使用公式编辑器排版，而是以普通文本或低清图片呈现",
    "090": "独立公式未居中或公式编号未在同一行右对齐",
    "091": "公式编号格式不符合要求",
    "092": "公式编号不连续，存在跳号、重号或章节号错误",
    "093": "正文引用的公式编号与实际公式编号不一致或不存在",
    "094": "变量未使用斜体，向量、矩阵、常量或函数的正斜体使用错误",
    "095": "上标、下标、分式、根号、积分号或求和号的层级和位置错误",
    "096": "公式中中文、英文、数字和符号的字体或字号不统一",
    "097": "公式前后段落间距过大、过小或与正文行距明显不协调",
    "098": "长公式未合理换行，超出页边距或被页面裁切",
    "099": "公式后标点缺失或误用，导致公式与上下文的语法关系错误",
    # 十、引用、注释与交叉引用检测
    "100": "直接引用他人文字却未标注来源、页码或引号",
    "101": "正文引用标注的位置错误，未置于相应内容或标点的规定位置",
    "102": "同一文档中顺序编码制与著者—出版年制混用",
    "103": "引用序号的括号、上标、字体或字号格式不统一",
    "104": "引用序号不连续、重复或未按首次出现顺序排列",
    "105": "脚注与尾注混用或注释类型不符合学校要求",
    "106": "脚注编号未按页或全文连续编号，出现跳号、重号或重置错误",
    "107": "脚注文字的字体、字号、行距、缩进或分隔线格式错误",
    "108": "交叉引用由手工输入，更新后引用编号未同步变化",
    "109": "交叉引用指向错误的章节、图、表、公式、注释或参考文献",
    "110": "链接或交叉引用显示“错误！未找到引用源”等域错误信息",
    # 十一、参考文献检测
    "111": "参考文献著录格式不符合GB/T 7714—2015或学校指定标准",
    "112": "文献类型标识错误或缺失，如[J]、[M]、[D]、[C]、[EB/OL]使用不当",
    "113": "作者姓名著录格式错误，作者人数、省略规则或“等”使用不当",
    "114": "题名、期刊名、出版社、出版地、年份、卷期或页码信息缺失",
    "115": "期刊的年、卷、期、起止页码排列顺序或标点格式错误",
    "116": "专著、学位论文、会议论文、标准、专利等不同类型文献套用了错误模板",
    "117": "DOI、URL或引用日期缺失、失效或书写格式不规范",
    "118": "中文文献使用英文标点或英文文献误用中文标点",
    "119": "英文文献的作者大写、题名大小写或期刊名称缩写不统一",
    "120": "正文引用与参考文献表无法一一对应，存在多列、漏列或编号不一致",
    "121": "参考文献排序方式错误，未按引用顺序或作者字母顺序排列",
    # 十二、字符、符号、标点、数字与空格检测
    "122": "中文句子中误用英文标点或英文句子中误用中文标点",
    "123": "逗号、句号、引号或括号缺失、多余或成对不完整",
    "124": "全角与半角字符混用，导致标点、字母或数字宽度不一致",
    "125": "中文与英文、数字之间的空格使用不符合统一规则",
    "126": "行首、行尾、标点前后或词语之间存在多余空格",
    "127": "英文单词之间、数字与单位之间存在缺失空格或错误空格",
    "128": "连续出现多个空行或用空行控制分页和段落间距",
    "129": "日期、时间、年份或编号格式不统一",
    "130": "计量单位大小写、正斜体、前缀或单位符号书写错误",
    "131": "百分号、摄氏度、范围号、连接号、破折号或省略号使用不规范",
    "132": "文档中包含乱码、不可见控制字符、异常换行符或无意义特殊符号",
    # 十三、一致性与关联性检测
    "133": "同一术语、概念、人名、地名或机构名称在文中写法不一致",
    "134": "论文题目在封面、声明、摘要、页眉和正文中的表述不一致",
    "135": "中文摘要与英文摘要中的题目、关键词、数据或结论不对应",
    "136": "目录中的章节标题、编号或页码与正文不一致",
    "137": "正文引用的图、表、公式、章节编号与实际对象不一致",
    "138": "图表中的数据、单位或名称与正文描述不一致",
    "139": "同级标题、同类图表、同类公式或同类注释的格式不一致",
    "140": "同一变量、缩写或符号在不同章节中的含义或格式不一致",
    "141": "正文引用与参考文献表之间存在对应关系错误",
    "142": "章节、小节或段落之间出现重复内容、顺序颠倒或逻辑断裂",
    "143": "文中统计总数、百分比、样本量或合计值前后不一致",
    # 十四、文字与语言规范检测
    "144": "存在错别字、形近字误用、漏字或重复字",
    "145": "存在用词不当、词语搭配错误或口语化表达",
    "146": "存在成分残缺、成分赘余、语序不当或句式杂糅等语法错误",
    "147": "句子过长、指代不清或歧义，影响学术内容理解",
    "148": "专有名词、学科术语或专业缩写的写法不规范或不统一",
    "149": "缩略语首次出现时未给出中文全称、英文全称或必要说明",
    "150": "语言风格不符合学术写作要求，使用夸张、主观或随意表达",
    "151": "数字、量词、年代、世纪或序数词的文字表达不规范",
    "152": "涉及国家、机构、民族、地区等名称时使用了不规范表述",
    "153": "政策、法律法规、党政用语或规范性文件名称引用不准确",
    "154": "文中含有批注残留、修订痕迹、模板提示语或待办占位文字",
    # 十五、附录及特殊内容检测
    "155": "正文已引用附录但文末缺少相应附录或附录未被正文引用",
    "156": "附录编号不连续或“附录A、附录B”等编号格式错误",
    "157": "附录标题的字体、字号、对齐方式或编号格式不符合要求",
    "158": "附录中的图、表、公式未采用独立编号，如“图A.1”“表B.1”",
    "159": "附录中的图表、公式、参考资料或数据来源格式与正文规则冲突",
    "160": "致谢标题、正文的字体、行距、措辞或署名格式不符合要求",
    "161": "攻读学位期间发表成果的条目格式、排序或作者标注不规范",
    "162": "问卷、访谈提纲、源代码、原始数据等特殊材料排版混乱或内容不完整",
    "163": "保密内容、个人隐私或敏感数据未按要求脱敏、隐藏或标注",
    "164": "超大表格、横向页面、折页或彩色页面未按特殊版式要求设置",
    "165": "文档末尾存在多余空白页、无关内容、隐藏对象或未删除的模板示例",
}


# ---------------------------------------------------------------------------
# 取证实现（Provider）：检测顺序即批注生成顺序，重构前后保持一致
# ---------------------------------------------------------------------------

PROVIDERS: tuple[ProviderSpec, ...] = (
    ProviderSpec(
        key="outline",
        label="模板章节结构提取与缺失检测（以模板自身的目录/标题为基准）",
        module="thesis_format_checker.detectors.outline",
        namespace="outline",
        extract="extract_outline_baseline",
        check="check_outline_errors",
        rules=("001", "002", "003", "004", "006", "007", "008", "009"),
    ),
    ProviderSpec(
        key="structure",
        label="结构完整性基线（封面/声明/关键词/目录/顺序）",
        module="thesis_format_checker.detectors.structure",
        namespace="structure",
        extract="extract_structure_baseline",
        check="check_structure_errors",
        rules=("001", "002", "005", "006", "009", "010", "011"),
    ),
    ProviderSpec(
        key="front",
        label="封面与前置部分基线（题目/指导教师/日期/摘要）",
        module="thesis_format_checker.detectors.front",
        namespace="front",
        extract="extract_front_baseline",
        check="check_front_errors",
        rules=("012", "013", "014", "015", "016", "018", "019", "020", "021"),
    ),
    ProviderSpec(
        key="cover_media",
        label="封面图片与艺术字（校徽、学校名称图片）",
        module="thesis_format_checker.detectors.cover_media",
        namespace="cover_media",
        extract="extract_cover_media_baseline",
        check="check_cover_media_errors",
        rules=("017",),
    ),
    ProviderSpec(
        key="pages",
        label="页面设置基线（方向/装订线/页眉距离/分页/版式/裁切）",
        module="thesis_format_checker.detectors.pages",
        namespace="pages",
        extract="extract_page_baseline",
        check="check_page_errors",
        rules=("022", "024", "026", "027", "029", "030", "031", "033"),
        check_takes_mapping=True,
    ),
    ProviderSpec(
        key="headers",
        label="页眉页脚与页码基线",
        module="thesis_format_checker.detectors.headers",
        namespace="headers",
        extract="extract_headers_baseline",
        check="check_header_footer_errors",
        rules=("034", "036", "037", "038", "039", "040", "041", "042", "043", "044"),
        check_takes_mapping=True,
    ),
    ProviderSpec(
        key="headings",
        label="标题层级与编号基线",
        module="thesis_format_checker.detectors.headings",
        namespace="headings",
        extract="extract_heading_numbering_baseline",
        check="check_heading_numbering_errors",
        rules=("048", "049", "050", "051", "052", "053", "054", "055"),
        extract_takes_mapping=True,
        check_takes_mapping=True,
    ),
    ProviderSpec(
        key="toc",
        label="目录基线（自动生成结构、条目、页码、格式）",
        module="thesis_format_checker.detectors.toc",
        namespace="toc",
        extract="extract_toc_baseline",
        check="check_toc_errors",
        rules=("056", "057", "058", "059", "060", "061", "062", "063", "064", "065", "066"),
        extract_takes_mapping=True,
        check_takes_mapping=True,
    ),
    ProviderSpec(
        key="paragraphs",
        label="段落格式基线（注释字号、同类一致性、对齐、缩进、分页）",
        module="thesis_format_checker.detectors.paragraphs",
        namespace="paragraphs",
        extract="extract_paragraph_baseline",
        check="check_paragraph_errors",
        rules=("069", "070", "071", "073", "076", "077"),
        extract_takes_mapping=True,
        check_takes_mapping=True,
    ),
    ProviderSpec(
        key="figures",
        label="图表基线（题注位置与格式、编号序列、三线表、表格格式）",
        module="thesis_format_checker.detectors.figures",
        namespace="figures",
        extract="extract_figure_table_baseline",
        check="check_figure_table_errors",
        rules=("078", "079", "080", "081", "083", "084"),
        extract_takes_mapping=True,
        check_takes_mapping=True,
    ),
    ProviderSpec(
        key="equations",
        label="公式基线（对齐、编号格式与序列、字体、间距、引用）",
        module="thesis_format_checker.detectors.equations",
        namespace="equations",
        extract="extract_equation_baseline",
        check="check_equation_errors",
        rules=("090", "091", "092", "093", "096", "097"),
        extract_takes_mapping=True,
        check_takes_mapping=True,
    ),
    ProviderSpec(
        key="notes",
        label="注释与交叉引用基线",
        module="thesis_format_checker.detectors.notes",
        namespace="notes",
        extract="extract_notes_baseline",
        check="check_note_reference_errors",
        rules=("103", "104", "105", "106", "107", "108", "110"),
        extract_takes_mapping=True,
        check_takes_mapping=True,
    ),
    ProviderSpec(
        key="bibliography",
        label="参考文献著录基线",
        module="thesis_format_checker.detectors.bibliography",
        namespace="bibliography",
        extract="extract_bibliography_baseline",
        check="check_bibliography_errors",
        rules=("111", "112", "113", "114", "115", "117", "118", "121"),
        extract_takes_mapping=True,
        check_takes_mapping=True,
    ),
    ProviderSpec(
        key="characters",
        label="字符与空格基线（全半角、边界空格、空行、日期形式）",
        module="thesis_format_checker.detectors.characters",
        namespace="characters",
        extract="extract_character_baseline",
        check="check_character_errors",
        rules=("124", "125", "126", "127", "128", "129", "130", "131", "132"),
        extract_takes_mapping=True,
        check_takes_mapping=True,
    ),
    ProviderSpec(
        key="consistency",
        label="内部一致性基线（题目、目录、图表引用）",
        module="thesis_format_checker.detectors.consistency",
        namespace="consistency",
        extract="extract_consistency_baseline",
        check="check_consistency_errors",
        rules=("134", "136", "137"),
        extract_takes_mapping=True,
        check_takes_mapping=True,
    ),
    ProviderSpec(
        key="formats",
        label="同类对象格式一致性基线（含特殊版式）",
        module="thesis_format_checker.detectors.formats",
        namespace="formats",
        extract="extract_format_consistency_baseline",
        check="check_format_consistency_errors",
        rules=("139", "164"),
        extract_takes_mapping=True,
        check_takes_mapping=True,
    ),
    ProviderSpec(
        key="appendices",
        label="附录基线（编号、标题格式、独立编号、与正文规则冲突）",
        module="thesis_format_checker.detectors.appendices",
        namespace="appendices",
        extract="extract_appendix_baseline",
        check="check_appendix_errors",
        rules=("155", "156", "157", "158", "159"),
        extract_takes_mapping=True,
        check_takes_mapping=True,
    ),
    ProviderSpec(
        key="hygiene",
        label="全文与文末卫生（残留批注/修订/占位、致谢格式、文末隐藏对象）",
        module="thesis_format_checker.detectors.hygiene",
        namespace="hygiene",
        extract="extract_hygiene_baseline",
        check="check_hygiene_errors",
        rules=("154", "160", "165"),
    ),
    ProviderSpec(
        key="language_lite",
        label="文字规范（不依赖大模型）：错别字词表、重复字、缩略语全称、数字写法、英文文献大小写",
        module="thesis_format_checker.detectors.language_lite",
        namespace="language_lite",
        extract="extract_language_baseline",
        check="check_language_errors",
        rules=("119", "144", "149", "151"),
    ),
    ProviderSpec(
        key="legacy",
        label="旧版 Word 有效格式流程（页边距/字体/字号/缩进/行距/标点）",
        module="thesis_format_checker.detectors.word_effective.runner",
        namespace=None,
        extract=None,
        check="compare_documents",
        rules=("023", "025", "035", "045", "046", "047", "067", "068", "069", "072",
               "074", "075", "122", "123"),
    ),
)

# ``legacy`` 提供者不是 docx 结构检测，单独标识，供状态判定使用。
LEGACY_PROVIDER_KEY = "legacy"
LEGACY_RULE_IDS: tuple[str, ...] = next(
    item.rules for item in PROVIDERS if item.key == LEGACY_PROVIDER_KEY
)


# ---------------------------------------------------------------------------
# 细粒度规则：状态标记、已删除清单与边界说明
# ---------------------------------------------------------------------------

#: 细粒度规则的状态标记 → 枚举
#: ``i`` 独立输出 / ``g`` 聚合输出（待拆分）/ ``p`` 待实现 / ``d`` 待确认
_ITEM_STATUS: dict[str, RuleStatus] = {
    "i": RuleStatus.IMPLEMENTED,
    "g": RuleStatus.AGGREGATED,
    "p": RuleStatus.PENDING,
    "d": RuleStatus.DEFERRED,
}

#: 已删除规则族 → 删除原因（必须依赖大模型语义理解）
REMOVED_REASONS: dict[str, str] = dict(REMOVED_FAMILIES)

#: 由“排版几何/分页事实”阶段（``detectors/rendered.py``）实现的规则族。
#: 调度层在 Word 流程拿到分页/坐标事实后直接调用它，因此这些族没有 check_sections 提供者，
#: 但状态必须显示为“已实现”，否则规则清单会与实际行为不一致。
RENDERED_FAMILIES: tuple[str, ...] = ("032", "087", "098")

# 已实现规则中“只实现了一部分”的边界说明。
RULE_NOTES: dict[str, str] = {
    "073": "只检测标题与图表题注的缩进；参考文献缩进见 073.3。",
    "121": "只检测文末可读取的排序形式，不判断“按正文首次引用顺序”的排序。",
    "159": "只比较可读取的格式规则；附录资料来源的真实性、完整性不检测。",
    # 以下规则族本阶段明确不实现，或已由其他规则覆盖（原因即备注）。
    "028": "页面尺寸/页边距差异是否“必要”依赖大模型语义判断，本阶段不实现。",
    "032": "空白页按 Word 实际分页判定（032.1 已实现）；“是否必要”的语义判断不检测。",
    "040": "页码跳号需 Word 实际分页；只读 docx 只能核对节起始值（040.1/040.2 已实现）。",
    "054": "标题与内容是否被分页分离按 Word 实际分页判定；纯 docx 单跑时只能给“无法判定”。",
    "056": "目录域是否被手工修改需要 Word 域结果比对（056.1 已实现）。",
    "057": "目录多余条目已由 136 目录一致性检测输出（057.1 已实现）。",
    "058": "目录文字/编号/层级不一致已由 136 目录一致性检测输出。",
    "059": "目录缓存页码与正文标题的打印页码比对，需要 Word 实际分页；跑完整流程时给出确定结论。",
    "060": "目录显示层级按 TOC 域 \\o 开关比对模板（060.1/060.2 已实现）。",
    "063": "目录页码位置是否对齐需实际渲染（063.1 已实现）。",
    "065": "目录孤立条目需 Word 分页判断是否跨页分离。",
    "069": "标题字号已由 045.3/046.3/047.3 按标题级别输出，不再重复。",
    "080": "图表编号重号待与 080.1 编号连续性检测一并细化。",
    "082": "正文与图表编号的对应关系已由 137.1 输出。",
    "085": "图表超出页边距已由 033.1 内容超界检测输出；尺寸是否过小属主观要求。",
    "087": "表格跨页与标题行复用按 Word 分页 + 每行“重复标题行”设置判定（087.1 已实现）。",
    "088": "题注缺失/数据来源/单位标注需结合图表语义判断，口径待确认。",
    "089": "判断“是否应使用公式编辑器”需语义识别疑似公式。",
    "093": "公式引用与对象一致性已由 137.3 输出。",
    "094": "数学变量/常量的正斜体规范依赖公式语义判断。",
    "095": "上下标层级是否正确依赖公式语义判断。",
    "096": "公式内部字体一致性待公式字符级解析（096.1/096.2 已实现）。",
    "098": "长公式是否越界按 Word 实际坐标测量（098.1 已实现）。",
    "106": "脚注显示编号由 Word 计算，id 与显示编号不一定一致，需 Word 判定。",
    "107": "脚注字号待与 107.1 字体检测一并细化。",
    "109": "交叉引用目标缺失已由 137.2 输出。",
    "127": "英文单词之间缺空格需词典/语义判断，本阶段不实现。",
    "161": "攻读学位期间成果条目的格式与排序口径待确认。",
    "165": "文末多余空白页按 Word 分页判定（165.1）；隐藏对象与残留示例见 165.2/165.3。",
}

#: 未实现 / 已删除的汇总说明（CLI 摘要使用）。
SKIPPED_SUMMARY_LINES: tuple[str, ...] = (
    f"【已删除】共 {len(REMOVED_FAMILIES)} 个规则族必须依赖大模型语义，已整体删除（见 docs/规则清单.md）",
    "【大类十四】文字与语言规范：可确定性部分已实现（119.1 英文作者大小写、144 错别字与重复字词、"
    "149 缩略语首次全称、151 数字写法），其余需大模型语义的族已删除",
    "【聚合输出】标记为“聚合输出”的细粒度规则已检测，但当前与同族其他条目合并成一条批注",
)


# ---------------------------------------------------------------------------
# 组装规则目录
# ---------------------------------------------------------------------------

def category_of(rule_id: str) -> CategorySpec:
    """返回规则族所属的大类。族编号连续，每类 11 条。"""

    number = int(rule_id)
    if not 1 <= number <= len(CATEGORIES) * 11:
        raise KeyError(f"规则编号超出范围：{rule_id}")
    return CATEGORIES[(number - 1) // 11]


def _build_fine_rules() -> dict[str, RuleItem]:
    """把 :data:`FINE_RULE_ROWS` 组装成细粒度规则索引。"""

    items: dict[str, RuleItem] = {}
    for code, title, severity, status, location in FINE_RULE_ROWS:
        family = family_of(code)
        items[code] = RuleItem(
            code=code,
            title=title,
            family=family,
            category=category_of(family).key,
            severity=severity,
            status=_ITEM_STATUS[status],
            location=location,
        )
    return items


#: 全部细粒度规则，键为 ``<族编号>[.<序号>]``。
FINE_RULES: dict[str, RuleItem] = _build_fine_rules()

_FINE_BY_FAMILY: dict[str, tuple[RuleItem, ...]] = {}
for _item in FINE_RULES.values():
    _FINE_BY_FAMILY[_item.family] = _FINE_BY_FAMILY.get(_item.family, ()) + (_item,)

#: 已删除的规则族编号。
REMOVED_RULE_IDS: tuple[str, ...] = tuple(rule_id for rule_id, _ in REMOVED_FAMILIES)

#: 业务口径待确认的规则族编号（族内存在待确认条目）。
def _provider_index() -> dict[str, tuple[str, ...]]:
    """反向索引：规则编号 → 产出它的取证实现（可能有多个，例如 069）。"""

    index: dict[str, list[str]] = {}
    for provider in PROVIDERS:
        for rule_id in provider.rules:
            index.setdefault(rule_id, []).append(provider.key)
    return {rule_id: tuple(keys) for rule_id, keys in index.items()}


def _status_of(rule_id: str, providers: tuple[str, ...]) -> RuleStatus:
    """判定规则族的实现状态。

    * 已删除的族 → ``REMOVED``；
    * 有提供者 → 已实现；其中只有旧版 Word 流程参与时标记为 ``LEGACY``；
    * 否则按族内细粒度规则的状态归为 ``DEFERRED``（待确认）或 ``PENDING``（待实现）。
    """

    if rule_id in REMOVED_REASONS:
        return RuleStatus.REMOVED
    # 由“排版几何”阶段实现的族（detectors/rendered.py）：调度层在 Word 流程之后
    # 直接调用，不经过 check_sections 的提供者机制，因此单独登记。
    if rule_id in RENDERED_FAMILIES:
        return RuleStatus.IMPLEMENTED
    if providers:
        if set(providers) == {LEGACY_PROVIDER_KEY}:
            return RuleStatus.LEGACY
        return RuleStatus.IMPLEMENTED
    # 族内细粒度规则全部为“待确认”、且没有待实现条目时，整族标记为待确认。
    items = _FINE_BY_FAMILY.get(rule_id, ())
    if items and any(item.status is RuleStatus.DEFERRED for item in items) \
            and not any(item.status is RuleStatus.PENDING for item in items):
        return RuleStatus.DEFERRED
    return RuleStatus.PENDING


def _family_severity(items: tuple[RuleItem, ...]) -> Severity:
    """规则族等级取族内最严重的一条（没有细粒度条目时按“错误”处理）。"""

    if not items:
        return Severity.ERROR
    return max((item.severity for item in items), key=lambda value: SEVERITY_ORDER[value])


def _build_rule(rule_id: str, providers: tuple[str, ...]) -> RuleSpec:
    """组装单个规则族的元数据。"""

    try:
        title = RULE_TITLES[rule_id]
    except KeyError as exc:  # 目录必须完整，缺失即视为代码错误
        raise KeyError(f"规则 {rule_id} 缺少标题，请补全 RULE_TITLES") from exc
    items = _FINE_BY_FAMILY.get(rule_id, ())
    return RuleSpec(
        rule_id=rule_id,
        title=title,
        category=category_of(rule_id).key,
        status=_status_of(rule_id, providers),
        providers=providers,
        note=RULE_NOTES.get(rule_id, ""),
        severity=_family_severity(items),
        aggregated=any(item.status is RuleStatus.AGGREGATED for item in items),
        items=items,
    )


def _build_rules() -> tuple[dict[str, RuleSpec], dict[str, RuleSpec]]:
    """把标题表、细粒度清单、提供者索引组装成规则族目录。"""

    providers = _provider_index()
    kept: dict[str, RuleSpec] = {}
    removed: dict[str, RuleSpec] = {}
    for rule_id in ALL_FAMILY_IDS:
        spec = _build_rule(rule_id, providers.get(rule_id, ()))
        if spec.status is RuleStatus.REMOVED:
            removed[rule_id] = spec
        else:
            kept[rule_id] = spec
    return kept, removed


#: 全部规则族编号（001–165，含已删除的族，便于追溯）。
ALL_FAMILY_IDS: tuple[str, ...] = tuple(f"{number:03d}" for number in range(1, 166))

#: 保留的规则族（已实现 / 待实现 / 待确认），键为族编号。
RULES, REMOVED_RULES = _build_rules()

#: 兼容旧名称：全部规则族编号。
ALL_RULE_IDS: tuple[str, ...] = ALL_FAMILY_IDS

#: 整族尚未实现的规则族（可确定性检测）：这些族一条结论都不会产出。
PENDING_RULE_IDS: tuple[str, ...] = tuple(sorted(
    family for family, spec in RULES.items() if spec.status is RuleStatus.PENDING
))

#: 整族“待确认/本阶段不实现”的规则族（依赖大模型语义、需 Word 渲染或口径待定）。
DEFERRED_RULE_IDS: tuple[str, ...] = tuple(sorted(
    family for family, spec in RULES.items() if spec.status is RuleStatus.DEFERRED
))

#: 已实现、但仍有个别细粒度条目待实现的规则族（例如 022 的“前置页码形式”）。
PARTIALLY_PENDING_RULE_IDS: tuple[str, ...] = tuple(sorted(
    family for family, spec in RULES.items()
    if spec.status in (RuleStatus.IMPLEMENTED, RuleStatus.LEGACY) and spec.pending_items
))


def rules_of_category(key: str, include_removed: bool = True) -> tuple[RuleSpec, ...]:
    """返回某个大类下的 11 个规则族（默认包含已删除的族，便于追溯）。"""

    category = CATEGORY_BY_KEY[key]
    result: list[RuleSpec] = []
    for rule_id in category.rule_ids:
        spec = RULES.get(rule_id) or REMOVED_RULES.get(rule_id)
        if spec is None:
            continue
        if spec.status is RuleStatus.REMOVED and not include_removed:
            continue
        result.append(spec)
    return tuple(result)


def fine_rules_of_category(key: str) -> tuple[RuleItem, ...]:
    """返回某个大类下的全部细粒度规则（保持清单顺序）。"""

    category = CATEGORY_BY_KEY[key]
    return tuple(item for item in FINE_RULES.values() if item.category == category.key)


def fine_rules_of_family(family: str) -> tuple[RuleItem, ...]:
    """返回某个规则族下的细粒度规则。"""

    return _FINE_BY_FAMILY.get(family, ())


def severity_of(rule_id: str) -> Severity:
    """取规则族（或细粒度规则编号）的等级，未知时按“错误”。"""

    if rule_id in FINE_RULES:
        return FINE_RULES[rule_id].severity
    spec = RULES.get(rule_id) or REMOVED_RULES.get(rule_id)
    return spec.severity if spec else Severity.ERROR


def provider_by_key(key: str) -> ProviderSpec:
    """按 key 查找取证实现。"""

    for provider in PROVIDERS:
        if provider.key == key:
            return provider
    raise KeyError(key)


def skipped_reasons() -> list[str]:
    """生成 CLI 摘要里“未实现 / 已删除”的说明列表（按细粒度规则逐条列出）。"""

    lines: list[str] = []
    for item in FINE_RULES.values():
        if item.status is RuleStatus.PENDING:
            lines.append(f"【{item.code}】{item.title}：待实现（可确定性检测）")
        elif item.status is RuleStatus.DEFERRED:
            lines.append(f"【{item.code}】{item.title}：待确认（业务口径）")
    lines.extend(SKIPPED_SUMMARY_LINES)
    return lines


def implemented_rule_ids() -> tuple[str, ...]:
    """返回当前真正会产出结论的规则族编号（含旧版流程）。"""

    return tuple(rule_id for rule_id in ALL_FAMILY_IDS
                 if rule_id in RULES
                 and RULES[rule_id].status in (RuleStatus.IMPLEMENTED, RuleStatus.LEGACY))


__all__ = [
    "ALL_FAMILY_IDS",
    "ALL_RULE_IDS",
    "CATEGORIES",
    "CATEGORY_BY_KEY",
    "CategorySpec",
    "DEFERRED_RULE_IDS",
    "FINE_RULES",
    "LEGACY_PROVIDER_KEY",
    "LEGACY_RULE_IDS",
    "PROVIDERS",
    "PARTIALLY_PENDING_RULE_IDS",
    "PENDING_RULE_IDS",
    "ProviderSpec",
    "REMOVED_REASONS",
    "REMOVED_RULE_IDS",
    "REMOVED_RULES",
    "RULE_NOTES",
    "RULE_TITLES",
    "RULES",
    "RuleItem",
    "RuleSpec",
    "RuleStatus",
    "SKIPPED_SUMMARY_LINES",
    "STATUS_LABELS",
    "category_of",
    "fine_rules_of_category",
    "fine_rules_of_family",
    "implemented_rule_ids",
    "provider_by_key",
    "rules_of_category",
    "severity_of",
    "skipped_reasons",
]
