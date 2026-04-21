#!/usr/bin/env python3
"""
将模板文件中的页眉页脚复制到目标文件
⚠️ 警告：这会覆盖目标文件的所有页眉页脚内容！
"""
import docx
from copy import deepcopy
import sys

TMPL = sys.argv[1] if len(sys.argv) > 1 else 'template.docx'
TARGET = sys.argv[2] if len(sys.argv) > 2 else 'target.docx'

def _get_odd_even_headers(section):
    """检查 section 是否启用奇偶页不同页眉页脚"""
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    sectPr = section._sectPr
    return sectPr.find(f'{{{ns}}}evenAndOddHeaders') is not None


def _set_odd_even_headers(section, enabled):
    """设置 section 的奇偶页不同页眉页脚"""
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    sectPr = section._sectPr
    existing = sectPr.find(f'{{{ns}}}evenAndOddHeaders')
    if enabled:
        if existing is None:
            sectPr.append(docx.oxml.OxmlElement('w:evenAndOddHeaders'))
    else:
        if existing is not None:
            sectPr.remove(existing)


def copy_headers_footers(tmpl_doc, target_doc):
    """复制页眉页脚"""

    min_sections = min(len(tmpl_doc.sections), len(target_doc.sections))

    print("页眉页脚复制报告:")
    print("-" * 60)

    for i in range(min_sections):
        tmpl_s = tmpl_doc.sections[i]
        target_s = target_doc.sections[i]

        # 复制页眉
        tmpl_header = tmpl_s.header
        target_header = target_s.header

        if tmpl_header.paragraphs or tmpl_header.tables:
            # 清空目标页眉
            for p in list(target_header.paragraphs):
                p._element.getparent().remove(p._element)
            for t in list(target_header.tables):
                t._element.getparent().remove(t._element)

            # 复制模板页眉内容
            for element in tmpl_header._element:
                if element.tag.endswith(('p', 'tbl')):
                    target_header._element.append(deepcopy(element))
            print(f"  [复制] Section {i+1} 页眉")

        # 复制页脚
        tmpl_footer = tmpl_s.footer
        target_footer = target_s.footer

        if tmpl_footer.paragraphs or tmpl_footer.tables:
            # 清空目标页脚
            for p in list(target_footer.paragraphs):
                p._element.getparent().remove(p._element)
            for t in list(target_footer.tables):
                t._element.getparent().remove(t._element)

            # 复制模板页脚内容
            for element in tmpl_footer._element:
                if element.tag.endswith(('p', 'tbl')):
                    target_footer._element.append(deepcopy(element))
            print(f"  [复制] Section {i+1} 页脚")

        # 复制页面设置属性
        target_s.different_first_page_header_footer = tmpl_s.different_first_page_header_footer
        _set_odd_even_headers(target_s, _get_odd_even_headers(tmpl_s))

    print("-" * 60)
    print("页眉页脚复制完成！")

def main():
    tmpl = docx.Document(TMPL)
    doc = docx.Document(TARGET)

    copy_headers_footers(tmpl, doc)
    doc.save(TARGET)
    print(f"已保存: {TARGET}")

if __name__ == '__main__':
    main()
