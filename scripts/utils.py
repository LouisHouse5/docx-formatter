#!/usr/bin/env python3
"""
docx-formatter 公共工具模块
提供：XML命名空间、EMU换算、字体设置、段落属性提取、奇偶页操作、文件检查等
"""
import sys
import os
from copy import deepcopy
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu

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
    run.font.size = Emu(size)
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
    return deepcopy(borders)


def apply_table_borders(table, borders_element):
    """将边框 XML 应用到表格"""
    tbl = table._tbl
    tblPr = tbl.find(ns_tag(W_TBLPR))
    if tblPr is None:
        tblPr = OxmlElement(W_TBLPR)
        tbl.insert(0, tblPr)

    old_borders = tblPr.find(ns_tag(W_TBLBORDERS))
    if old_borders is not None:
        tblPr.remove(old_borders)
    tblPr.append(deepcopy(borders_element))


# ==================== 表格默认边框 ====================
def apply_default_table_borders(table):
    """给表格应用默认边框（单线，宽度 4 = 1/2 pt）"""
    tbl = table._tbl
    tblPr = tbl.find(ns_tag(W_TBLPR))
    if tblPr is None:
        tblPr = OxmlElement(W_TBLPR)
        tbl.insert(0, tblPr)

    borders = tblPr.find(ns_tag(W_TBLBORDERS))
    if borders is None:
        borders = OxmlElement(W_TBLBORDERS)
        tblPr.append(borders)

    # 清除旧边框设置
    for child in list(borders):
        borders.remove(child)

    # 添加新边框
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
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


def compare_para_props(target_info, tmpl_info):
    """对比目标段落与模板段落的属性差异，返回差异描述列表"""
    issues = []
    if target_info['font_name'] != tmpl_info['font_name'] and tmpl_info['font_name']:
        issues.append(f"font={target_info['font_name']}(应{tmpl_info['font_name']})")
    if target_info['font_size'] != tmpl_info['font_size'] and tmpl_info['font_size']:
        issues.append(f"size={target_info['font_size']}(应{tmpl_info['font_size']})")
    if target_info['alignment'] != tmpl_info['alignment']:
        issues.append(f"align={target_info['alignment']}(应{tmpl_info['alignment']})")
    if target_info['bold'] != tmpl_info['bold'] and tmpl_info['bold'] is not None:
        issues.append(f"bold={target_info['bold']}(应{tmpl_info['bold']})")
    if target_info['line_spacing'] != tmpl_info['line_spacing'] and tmpl_info['line_spacing'] is not None:
        issues.append(f"ls={target_info['line_spacing']}(应{tmpl_info['line_spacing']})")
    if target_info['line_spacing_rule'] != tmpl_info['line_spacing_rule']:
        issues.append(f"ls_rule={target_info['line_spacing_rule']}(应{tmpl_info['line_spacing_rule']})")
    if target_info['style'] != tmpl_info['style'] and tmpl_info['style']:
        issues.append(f"style={target_info['style']}(应{tmpl_info['style']})")
    if target_info['first_line_indent'] != tmpl_info['first_line_indent'] and tmpl_info['first_line_indent'] is not None:
        issues.append(f"fi={target_info['first_line_indent']}(应{tmpl_info['first_line_indent']})")
    if target_info['space_before'] != tmpl_info['space_before'] and tmpl_info['space_before'] is not None:
        issues.append(f"sb={target_info['space_before']}(应{tmpl_info['space_before']})")
    if target_info['space_after'] != tmpl_info['space_after'] and tmpl_info['space_after'] is not None:
        issues.append(f"sa={target_info['space_after']}(应{tmpl_info['space_after']})")
    return issues


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
