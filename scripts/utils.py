#!/usr/bin/env python3
"""
docx-formatter 公共工具模块
提供：XML命名空间、EMU换算、字体设置、段落属性提取、奇偶页操作、文件检查等
"""
import sys
import os
from docx.oxml.ns import qn

# ==================== XML 命名空间 ====================
NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'

# python-docx 内部使用的命名空间字典（用于 find/findall）
NSMAP = {'w': NS_W, 'w14': NS_W14}

# 常用 XML tag 前缀（用于 OxmlElement）
W_P = 'w:p'
W_R = 'w:r'
W_T = 'w:t'
W_PPR = 'w:pPr'
W_RPR = 'w:rPr'
W_RFONTS = 'w:rFonts'
W_TBL = 'w:tbl'
W_TBLPR = 'w:tblPr'
W_TBLBORDERS = 'w:tblBorders'
W_JC = 'w:jc'
W_TBW = 'w:tblW'
W_TCW = 'w:tcW'
W_TC = 'w:tc'
W_TCPR = 'w:tcPr'
W_SECTPR = 'w:sectPr'
W_TYPE = 'w:type'
W_OUTLINELVL = 'w:outlineLvl'
W_EVENODD = 'w:evenAndOddHeaders'
W_FLDCHAR = 'w:fldChar'
W_FLDSIMPLE = 'w:fldSimple'
W_INSTRTEXT = 'w:instrText'
W_VAL = f'{{{NS_W}}}val'
W_SZ = f'{{{NS_W}}}sz'
W_SPACE = f'{{{NS_W}}}space'
W_COLOR = f'{{{NS_W}}}color'


def ns_tag(tag_name):
    """将短标签名转换为带命名空间的完整标签名，如 'w:p' -> '{http://...}p'"""
    # 支持带前缀的短形式 'w:tag' 或直接使用 'tag'
    if ':' in tag_name:
        prefix, local = tag_name.split(':', 1)
        if prefix == 'w':
            return f'{{{NS_W}}}{local}'
        elif prefix == 'w14':
            return f'{{{NS_W14}}}{local}'
    return f'{{{NS_W}}}{tag_name}'


def ns_attr(attr_name):
    """将短属性名转换为带命名空间的完整属性名，如 'val' -> '{http://...}val'"""
    return f'{{{NS_W}}}{attr_name}'


# ==================== EMU 换算 ====================
def emu_to_pt(emu):
    """EMU 转磅值（pt），1 pt = 12700 EMU"""
    return emu / 12700 if emu else None


def emu_to_inch(emu):
    """EMU 转英寸，1 inch = 914400 EMU"""
    return emu / 914400 if emu else None


def pt_to_emu(pt):
    """磅值（pt）转 EMU"""
    return int(pt * 12700)


def inch_to_emu(inch):
    """英寸转 EMU"""
    return int(inch * 914400)


# ==================== 字体设置 ====================
def set_run_font(run, name='宋体', size=152400, bold=False, italic=False, underline=False, color=None):
    """设置 run 字体，包含东亚字体属性 w:eastAsia 和完整格式"""
    run.font.name = name
    run.font.size = __import__('docx.shared', fromlist=['Emu']).Emu(size)
    if bold is None:
        run.font.bold = None
    else:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if underline is not None:
        run.font.underline = underline
    if color is not None:
        run.font.color.rgb = color
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), name)


def format_para_runs(p, name='宋体', size=152400, bold=False):
    """格式化段落中所有 run"""
    for run in p.runs:
        set_run_font(run, name, size, bold)


# ==================== 段落属性提取 ====================
def get_para_props(p, max_text_len=60, with_outline=False):
    """提取段落的显式格式属性，返回字典"""
    text = p.text.strip()
    if not text:
        return None
    runs = p.runs
    first_run = runs[0] if runs else None

    style_name = p.style.name if p.style else None

    result = {
        'text': text[:max_text_len],
        'style': style_name,
        'alignment': p.alignment,
        'line_spacing': p.paragraph_format.line_spacing,
        'line_spacing_rule': str(p.paragraph_format.line_spacing_rule),
        'space_before': p.paragraph_format.space_before,
        'space_after': p.paragraph_format.space_after,
        'first_line_indent': p.paragraph_format.first_line_indent,
        'left_indent': p.paragraph_format.left_indent,
        'right_indent': p.paragraph_format.right_indent,
        'font_name': first_run.font.name if first_run else None,
        'font_size': first_run.font.size if first_run else None,
        'bold': first_run.font.bold if first_run else None,
        'italic': first_run.font.italic if first_run else None,
        'underline': first_run.font.underline if first_run else None,
    }

    if with_outline:
        outline_lvl = None
        try:
            pPr = p._element.find(ns_tag(W_PPR))
            if pPr is not None:
                outline = pPr.find(ns_tag(W_OUTLINELVL))
                if outline is not None:
                    outline_lvl = outline.get(ns_attr('val'))
        except:
            pass
        result['outline_lvl'] = outline_lvl

    return result


# ==================== 奇偶页操作 ====================
def get_odd_even_headers(section):
    """检查 section 是否启用奇偶页不同页眉页脚"""
    sectPr = section._sectPr
    return sectPr.find(ns_tag(W_EVENODD)) is not None


def set_odd_even_headers(section, enabled):
    """设置 section 的奇偶页不同页眉页脚"""
    ox = __import__('docx.oxml', fromlist=['OxmlElement']).OxmlElement
    sectPr = section._sectPr
    existing = sectPr.find(ns_tag(W_EVENODD))
    if enabled:
        if existing is None:
            sectPr.append(ox(W_EVENODD))
    else:
        if existing is not None:
            sectPr.remove(existing)


# ==================== 文件/参数辅助 ====================
def check_file_exists(filepath, label='文件'):
    """检查文件是否存在，不存在则打印错误并退出"""
    if not os.path.isfile(filepath):
        print(f"错误: {label} 不存在: {filepath}", file=sys.stderr)
        sys.exit(1)


def parse_args(defaults, min_args=1):
    """
    解析命令行参数，支持按位置传入。
    defaults: 参数默认值列表，如 ['template.docx', 'target.docx']
    min_args: 最少需要的参数数量
    返回: (arg1, arg2, ...) 元组
    """
    args = []
    for i, default in enumerate(defaults):
        arg = sys.argv[i + 1] if len(sys.argv) > i + 1 else default
        args.append(arg)
    for i in range(min_args):
        if args[i] == defaults[i] and not os.path.isfile(args[i]):
            # 使用默认值但文件不存在
            pass  # 由调用方处理
    return tuple(args)


# ==================== 表格边框 XML 操作 ====================
def get_table_borders_xml(table):
    """提取表格的边框 XML 元素（深拷贝）"""
    tbl = table._tbl
    tblPr = tbl.find(ns_tag(W_TBLPR))
    if tblPr is None:
        return None
    borders = tblPr.find(ns_tag(W_TBLBORDERS))
    if borders is None:
        return None
    return __import__('copy', fromlist=['deepcopy']).deepcopy(borders)


def apply_table_borders(table, borders_element):
    """将边框 XML 应用到表格"""
    ox = __import__('docx.oxml', fromlist=['OxmlElement']).OxmlElement
    tbl = table._tbl
    tblPr = tbl.find(ns_tag(W_TBLPR))
    if tblPr is None:
        tblPr = ox(W_TBLPR)
        tbl.insert(0, tblPr)

    old_borders = tblPr.find(ns_tag(W_TBLBORDERS))
    if old_borders is not None:
        tblPr.remove(old_borders)
    tblPr.append(__import__('copy', fromlist=['deepcopy']).deepcopy(borders_element))


# ==================== 表格默认边框 ====================
def apply_default_table_borders(table):
    """给表格应用默认边框（单线，宽度 4 = 1/2 pt）"""
    ox = __import__('docx.oxml', fromlist=['OxmlElement']).OxmlElement
    tbl = table._tbl
    tblPr = tbl.find(ns_tag(W_TBLPR))
    if tblPr is None:
        tblPr = ox(W_TBLPR)
        tbl.insert(0, tblPr)

    borders = tblPr.find(ns_tag(W_TBLBORDERS))
    if borders is None:
        borders = ox(W_TBLBORDERS)
        tblPr.append(borders)

    # 清除旧边框设置
    for child in list(borders):
        borders.remove(child)

    # 添加新边框
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = ox(f'w:{border_name}')
        border.set(ns_attr('val'), 'single')
        border.set(ns_attr('sz'), '4')
        border.set(ns_attr('space'), '0')
        border.set(ns_attr('color'), 'auto')
        borders.append(border)


# ==================== TOC 域检测 ====================
def has_toc_field(doc):
    """检测文档是否已包含 TOC 域"""
    for p in doc.paragraphs:
        # 检查 fldSimple
        fld_simple = p._element.find(f'.//{W_FLDSIMPLE}', NSMAP)
        if fld_simple is not None:
            instr = fld_simple.get(ns_attr('instr'))
            if instr and 'TOC' in instr:
                return True
        # 检查 fldChar
        fld_chars = p._element.findall(f'.//{W_FLDCHAR}', NSMAP)
        for fld in fld_chars:
            fld_type = fld.get(ns_attr('fldCharType'))
            if fld_type == 'begin':
                instr_texts = p._element.findall(f'.//{W_INSTRTEXT}', NSMAP)
                for it in instr_texts:
                    if it.text and 'TOC' in it.text:
                        return True
    return False


# ==================== 彩色输出 ====================
class Colors:
    """终端彩色输出码（如果终端不支持则自动降级）"""
    OK = '\033[92m'      # 绿色
    WARN = '\033[93m'    # 黄色
    ERR = '\033[91m'     # 红色
    INFO = '\033[94m'    # 蓝色
    BOLD = '\033[1m'
    END = '\033[0m'

    @classmethod
    def disable(cls):
        """禁用所有颜色（用于不支持颜色的终端）"""
        cls.OK = cls.WARN = cls.ERR = cls.INFO = cls.BOLD = cls.END = ''


def _supports_color():
    """检测终端是否支持彩色输出"""
    return hasattr(sys.stdout, 'isatty') and sys.stdout.isatty()

if not _supports_color():
    Colors.disable()


def log_ok(msg):
    print(f"{Colors.OK}✓{Colors.END} {msg}")


def log_warn(msg):
    print(f"{Colors.WARN}⚠{Colors.END} {msg}")


def log_err(msg):
    print(f"{Colors.ERR}✗{Colors.END} {msg}", file=sys.stderr)


def log_info(msg):
    print(f"{Colors.INFO}ℹ{Colors.END} {msg}")


# ==================== 统一错误处理 ====================
def fix_quotes(doc):
    """半角引号转全角引号（智能开闭匹配，包含单双引号）

    Args:
        doc: python-docx Document 对象
    """
    def smart_quote_replace(text):
        """智能判断开闭引号"""
        result = []
        for i, ch in enumerate(text):
            if ch == '"':
                # 段落开头、空格或中文标点前为开引号
                if i == 0 or text[i-1] in ' \t\n\r（【《「『"\u201c\u2018：。，？！；、':
                    result.append('\u201c')  # "
                else:
                    result.append('\u201d')  # "
            elif ch == "'":
                if i == 0 or text[i-1] in ' \t\n\r（【《「『\u201c\u2018：。，？！；、':
                    result.append('\u2018')  # '
                else:
                    result.append('\u2019')  # '
            else:
                result.append(ch)
        return ''.join(result)

    for p in doc.paragraphs:
        for run in p.runs:
            if '"' in run.text or "'" in run.text:
                run.text = smart_quote_replace(run.text)


def run_with_errors(func):
    """装饰器：包装 main 函数，统一捕获异常并输出友好错误信息"""
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except FileNotFoundError as e:
            log_err(f"文件不存在: {e}")
            sys.exit(2)
        except PermissionError as e:
            log_err(f"权限不足: {e}")
            sys.exit(3)
        except Exception as e:
            log_err(f"运行时错误: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)
    return wrapper


def check_write_permission(filepath, label='文件'):
    """检查文件是否可写（如文件不存在则检查父目录）"""
    if os.path.exists(filepath):
        if not os.access(filepath, os.W_OK):
            print(f"错误: {label} 没有写入权限: {filepath}", file=sys.stderr)
            sys.exit(3)
    else:
        parent = os.path.dirname(os.path.abspath(filepath)) or '.'
        if not os.access(parent, os.W_OK):
            print(f"错误: 目录没有写入权限: {parent}", file=sys.stderr)
            sys.exit(3)
