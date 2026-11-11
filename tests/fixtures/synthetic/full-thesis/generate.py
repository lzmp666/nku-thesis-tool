"""full-thesis/ 的统一生成入口：样本清单 + 内存重建。

目录里两份 docx 的关系是「主样本 → 派生样本」：

* `synthetic_complete_thesis.docx` —— 由 `generate_full_thesis.py` 用 python-docx 造。
* `synthetic_object_layout_errors.docx` —— 由 `generate_object_layout_errors.py`
  在主样本字节上派生（一幅图片与一张固定列宽表格改成 600 pt）。

派生只吃**内存里**的主样本字节，不读磁盘文件。主样本万一被改坏，派生结果也会
跟着不同，逐部件重建能一起抓出来；不会出现「源文件变了、派生样本还是旧的，两份
都算通过」这种情况。

`generate_unstyled_thesis.py` 的产物没有提交进仓库（它只用于手工排查），不在
`SAMPLES` 里。

用法（在本目录下）：

    python generate.py --check    # 重建两份并与磁盘逐部件比对，不写盘
    python generate.py            # 重建并写盘（覆盖本目录的已提交样本）
    python generate_full_thesis.py
    python generate_object_layout_errors.py --output <新文件>

依赖 python-docx、Pillow、lxml。重建结果的整文件字节不稳定：python-docx 保存时
ZIP 条目时间戳取自当前时刻，所以 `--check` 比的是逐部件解压后的字节，与
`verify_fixtures.py` 同一口径。
"""
from __future__ import annotations

import argparse
import importlib.util
import sys
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

HERE = Path(__file__).resolve().parent

COMPLETE = "synthetic_complete_thesis.docx"
OBJECT_ERRORS = "synthetic_object_layout_errors.docx"
SAMPLES = (COMPLETE, OBJECT_ERRORS)


def _load(filename):
    """按文件路径加载同目录的生成器（目录名带连字符，不能当包导入）。"""
    module_name = "_full_thesis_" + Path(filename).stem
    if module_name in sys.modules:
        return sys.modules[module_name]
    spec = importlib.util.spec_from_file_location(module_name, HERE / filename)
    if spec is None or spec.loader is None:
        raise RuntimeError("无法加载 %s" % filename)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def build_bytes(name=COMPLETE):
    """返回指定样本的 docx 字节。未知样本名直接报错，不静默回落到主样本。"""
    if name == COMPLETE:
        return _load("generate_full_thesis.py").build_bytes()
    if name == OBJECT_ERRORS:
        source = build_bytes(COMPLETE)
        return _load("generate_object_layout_errors.py").build_bytes(source)
    raise KeyError("未知样本：%s（可选 %s）" % (name, "、".join(SAMPLES)))


def read_parts(data):
    with ZipFile(BytesIO(data)) as package:
        return [(info.filename, package.read(info.filename)) for info in package.infolist()]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="只逐部件比对，不写盘")
    arguments = parser.parse_args()

    failed = 0
    for name in SAMPLES:
        data = build_bytes(name)
        target = HERE / name
        if arguments.check:
            if not target.exists():
                print("缺失：%s" % name)
                failed += 1
                continue
            expected, actual = read_parts(data), read_parts(target.read_bytes())
            if [part for part, _ in expected] != [part for part, _ in actual]:
                print("部件清单不同：%s" % name)
                failed += 1
                continue
            different = [
                part for (part, want), (_, got) in zip(expected, actual) if want != got
            ]
            if different:
                print("与重建结果不一致：%s（%s）" % (name, "、".join(different)))
                failed += 1
                continue
            print("校验通过：%s（%d 个部件逐部件一致）" % (name, len(expected)))
        else:
            target.write_bytes(data)
            print("已生成：%s（%d 字节）" % (name, len(data)))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
