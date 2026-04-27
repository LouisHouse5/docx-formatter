#!/usr/bin/env python3
"""
将模板文件中的样式定义复制到目标文件
用于解决目标文件缺少模板样式（如 toc 1, toc 2, Heading 1 等）的问题
"""
import docx
from docx.oxml.ns import qn
from copy import deepcopy
import sys
import argparse

from utils import check_file_exists
from utils import run_with_errors, check_write_permission, log_ok, log_warn, log_err

parser = argparse.ArgumentParser(description='将模板样式复制到目标文件')
parser.add_argument('template', nargs='?', default='template.docx', help='模板 docx 文件路径')
parser.add_argument('target', nargs='?', default='target.docx', help='目标 docx 文件路径')
args = parser.parse_args()
TMPL = args.template
TARGET = args.target

def copy_styles(tmpl_doc, target_doc):
    """将模板中的关键样式复制到目标文件"""

    # 要复制的关键样式
    key_styles = ['Normal', 'Heading 1', 'Heading 2', 'Heading 3',
                  'toc 1', 'toc 2', 'toc 3', 'Header', 'Footer',
                  'Title', 'Subtitle']

    tmpl_styles_map = {}
    for style in tmpl_doc.styles:
        if style.name in key_styles:
            tmpl_styles_map[style.name] = style

    target_styles_map = {}
    for style in target_doc.styles:
        target_styles_map[style.name] = style

    print("样式复制报告:")
    print("-" * 60)

    for name in key_styles:
        if name not in tmpl_styles_map:
            print(f"  [跳过] 模板中不存在样式: {name}")
            continue

        tmpl_style = tmpl_styles_map[name]
        tmpl_element = tmpl_style._element

        if name in target_styles_map:
            # 替换现有样式
            target_style = target_styles_map[name]
            target_element = target_style._element
            parent = target_element.getparent()
            idx = parent.index(target_element)
            parent.remove(target_element)
            parent.insert(idx, deepcopy(tmpl_element))
            print(f"  [替换] 样式: {name}")
        else:
            # 添加新样式
            styles_element = target_doc.styles._element
            styles_element.append(deepcopy(tmpl_element))
            print(f"  [新增] 样式: {name}")

    print("-" * 60)
    print("样式复制完成！")

@run_with_errors
def main():
    check_file_exists(TMPL, '模板文件')
    check_file_exists(TARGET, '目标文件')
    check_write_permission(TARGET, '目标文件')

    tmpl = docx.Document(TMPL)
    doc = docx.Document(TARGET)

    copy_styles(tmpl, doc)
    doc.save(TARGET)
    log_ok(f"已保存: {TARGET}")

if __name__ == '__main__':
    main()
