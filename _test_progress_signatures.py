"""静态检查：所有 progress= 调用点的方法签名是否都接受 progress 参数。

背景：第 10 条为长耗时方法增加了 progress 进度回调，逐个手工添加时
容易漏掉某个方法的形参，运行到那一步才抛
"got an unexpected keyword argument 'progress'"。本脚本在不开 Word 的
情况下把这类漏改一次查出来。

运行：python _test_progress_signatures.py
"""
import ast
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
TARGETS = ("main_form.py", "event_handlers.py", "set_document_format.py")

# 需要 progress 回调的 SetDocumentFormat 方法（自 set_document_format.py 解析）
# 以及从其它模块调用的方法名 -> 是否必须接受 progress
SETTER_NAMES = {
    "set_page_margins", "set_standard_line_spacing", "set_content_format",
    "set_content_style", "snapshot_protected_paragraph_format",
    "restore_protected_paragraph_format", "set_main_title_format",
    "set_title_styles", "set_toc_line_spacing", "set_images_and_tables",
    "set_image_paragraph_no_indent", "set_table_paragraph_no_indent",
    "set_table_paragraph_spacing", "set_table_surrounding_spacing",
    "center_all_images", "set_tables_auto_adjust_and_align",
    "set_all_tables_max_width", "add_image_border", "delete_all_pictures",
    "cancel_wrap_as_inline", "get_paragraph_outline_level",
    "_auto_adjust_tables", "_apply_table_alignment",
}


def _args_of(func):
    """返回函数签名里接受的关键字/位置参数名（不含 *args/**kwargs 展开体）。"""
    a = func.args
    names = [x.arg for x in a.posonlyargs + a.args + a.kwonlyargs]
    if a.vararg:
        names.append("*" + a.vararg.arg)
    if a.kwarg:
        names.append("**" + a.kwarg.arg)
    return names


def collect_definitions(path):
    """收集 {方法名: 参数名列表}，只针对 SetDocumentFormat 类内的方法。"""
    with open(path, encoding="utf-8") as fh:
        tree = ast.parse(fh.read(), filename=path)
    defs = {}
    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef) and node.name == "SetDocumentFormat":
            for item in node.body:
                if isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    defs[item.name] = _args_of(item)
    return defs


def collect_progress_calls(path):
    """收集本文件里所有 `xxx(...progress=...)` 调用点：(方法名, 行号)。"""
    with open(path, encoding="utf-8") as fh:
        tree = ast.parse(fh.read(), filename=path)
    calls = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        func = node.func
        name = None
        if isinstance(func, ast.Attribute):
            name = func.attr
        elif isinstance(func, ast.Name):
            name = func.id
        if name is None:
            continue
        if any(kw.arg == "progress" for kw in node.keywords):
            calls.append((name, node.lineno))
    return calls


def main():
    defs = collect_definitions(os.path.join(ROOT, "set_document_format.py"))
    assert defs, "未能从 set_document_format.py 解析出 SetDocumentFormat 方法"

    problems = []
    checked = 0
    for fname in TARGETS:
        for name, lineno in collect_progress_calls(os.path.join(ROOT, fname)):
            if name not in defs:
                # 非 SetDocumentFormat 的方法（如 main_form 自己的
                # _make_status_reporter 等）：跳过
                if name in SETTER_NAMES:
                    problems.append(
                        f"{fname}:{lineno} 调用了未定义的 {name}(progress=…)")
                continue
            checked += 1
            args = defs[name]
            if "progress" not in args and "**" not in "".join(
                    a for a in args if a.startswith("**")):
                has_kwargs = any(a.startswith("**") for a in args)
                if not has_kwargs:
                    problems.append(
                        f"{fname}:{lineno} 调用 {name}(progress=…) "
                        f"但 set_document_format.py 的 {name}"
                        f"({', '.join(args[1:])}) 不接受该参数")

    if problems:
        print("发现问题：")
        for p in problems:
            print("  -", p)
        return 1

    print(f"检查通过：{checked} 处 progress= 调用点，"
          f"SetDocumentFormat 共 {len(defs)} 个方法签名全部匹配")
    return 0


if __name__ == "__main__":
    sys.exit(main())
