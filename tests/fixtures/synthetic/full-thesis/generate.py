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

依赖 python-docx、Pillow、lxml。python-docx 保存时 ZIP 条目时间戳取当前时刻，
所以最终字节过一遍 `_normalize()`：条目顺序原样保留，时间戳统一成 `FIXED_TIME`。
两份样本因此是整文件字节可复现的，`--check` 直接比整文件字节；字节不同时再逐部件
定位到具体是哪个部件变了。`verify_fixtures.py` 仍按逐部件比对，不用重写 ZIP。
"""
from __future__ import annotations

import argparse
import importlib.util
import sys
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

HERE = Path(__file__).resolve().parent

# 与 format-samples/generate.py 用同一个时间常量：ZIP 可存的最早时间，避免取当前时刻。
FIXED_TIME = (2000, 1, 1, 0, 0, 0)

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
        return _normalize(_load("generate_full_thesis.py").build_bytes())
    if name == OBJECT_ERRORS:
        source = build_bytes(COMPLETE)
        return _normalize(_load("generate_object_layout_errors.py").build_bytes(source))
    raise KeyError("未知样本：%s（可选 %s）" % (name, "、".join(SAMPLES)))


def _normalize(data):
    """把 ZIP 条目的写入时间换成 FIXED_TIME，部件内容与顺序原样保留。

    python-docx 保存时用当前时刻当条目时间戳，同一个生成器在不同分钟跑出的字节
    就不同，哈希没法当判据。这里只重写容器的头部字段，解压后的部件一字不改。
    不排序条目：磁盘样本按 python-docx 的写入顺序排列，排序会改顺序，反而对不上。
    """
    buffer = BytesIO()
    with ZipFile(BytesIO(data)) as source, ZipFile(buffer, "w") as target:
        for info in source.infolist():
            item = ZipInfo(info.filename, date_time=FIXED_TIME)
            item.compress_type = ZIP_DEFLATED
            item.external_attr = info.external_attr
            target.writestr(item, source.read(info.filename))
    return buffer.getvalue()


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
            actual = target.read_bytes()
            if data == actual:
                print("校验通过：%s（整文件字节一致，%d 字节）" % (name, len(data)))
                continue
            # 整文件不等时定位到部件：时间戳已固定，剩下的差异只可能来自内容。
            expected, onDisk = read_parts(data), read_parts(actual)
            if [part for part, _ in expected] != [part for part, _ in onDisk]:
                print("部件清单不同：%s" % name)
                failed += 1
                continue
            different = [
                part for (part, want), (_, got) in zip(expected, onDisk) if want != got
            ]
            if different:
                print("与重建结果不一致：%s（%s）" % (name, "、".join(different)))
            else:
                print("与重建结果不一致：%s（部件相同，差异在 ZIP 容器头部）" % name)
            failed += 1
        else:
            target.write_bytes(data)
            print("已生成：%s（%d 字节）" % (name, len(data)))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
