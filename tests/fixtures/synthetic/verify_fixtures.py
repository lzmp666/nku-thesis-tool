"""逐部件校验已提交的合成样本与生成器产物是否一致。

各生成器的 --check 是整文件逐字节比对，适合在同一台机器上复现；一旦 Python 或
zlib 版本变化，deflate 流可能不同，逐字节比对就会误报。本脚本只比较：

  * ZIP 条目名与顺序
  * 每个条目解压后的字节（CRC 由 zipfile 顺带校验）

不比较压缩流、时间戳、外部属性、注释。因此在 ubuntu-latest / windows-2022
上都能得到同样结论，适合放进 CI。生成器仍然只改自己目录下的样本。

用法（从仓库根目录）：
    python3 tests/fixtures/synthetic/verify_fixtures.py
    python3 tests/fixtures/synthetic/verify_fixtures.py --self-test
"""
import argparse
import importlib.util
import sys
import zipfile
from io import BytesIO
from pathlib import Path

HERE = Path(__file__).resolve().parent

# (目录, 生成器文件, 样本清单来源, 构建函数, 构建函数是否接受样本名)
GENERATORS = [
    ("document", "generate.py", "SAMPLES", "build_bytes", True),
    ("formula", "generate.py", "SAMPLES", "build", True),
    ("demo", "generate_samples.py", "SAMPLES", "build_bytes", True),
    ("demo", "generate_checked_sample.py", "OUTPUT", "build_bytes", False),
]


def load_generator(directory, filename):
    path = HERE / directory / filename
    module_name = "_fixture_gen_" + directory + "_" + Path(filename).stem
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载生成器：{path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def read_parts(data):
    """返回 [(条目名, 解压后字节)]；顺序也是被比较的一部分。"""
    with zipfile.ZipFile(BytesIO(data)) as z:
        return [(info.filename, z.read(info.filename)) for info in z.infolist()]


def diff_parts(expected, actual):
    """返回差异说明列表，空列表表示逐部件一致。"""
    problems = []
    expected_names = [name for name, _ in expected]
    actual_names = [name for name, _ in actual]
    if expected_names != actual_names:
        problems.append(
            "条目名/顺序不同：生成器 %s，仓库 %s"
            % (expected_names, actual_names)
        )
        return problems
    for (name, want), (_, got) in zip(expected, actual):
        if want != got:
            offset = next(
                (i for i, (a, b) in enumerate(zip(want, got)) if a != b),
                min(len(want), len(got)),
            )
            problems.append(
                "%s 内容不同（生成器 %d 字节 / 仓库 %d 字节，首个差异在 %d）"
                % (name, len(want), len(got), offset)
            )
    return problems


def rezip(data, level, timestamp):
    """换一个压缩级别和时间戳重建 ZIP，用来模拟另一台机器上的产物。"""
    parts = read_parts(data)
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, compresslevel=level) as z:
        for name, payload in parts:
            info = zipfile.ZipInfo(name, date_time=timestamp)
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, payload)
    return buf.getvalue()


def self_test():
    """证明这套比较既容忍压缩流差异，又能抓出内容改动。"""
    module = load_generator("demo", "generate_samples.py")
    original = module.build_bytes("known_issues_sample.docx")
    expected = read_parts(original)
    checks = []

    rebuilt = rezip(original, level=1, timestamp=(2025, 6, 1, 12, 0, 0))
    byte_same = rebuilt == original
    parts_same = diff_parts(expected, read_parts(rebuilt)) == []
    checks.append(
        ("换 deflate 级别/时间戳后逐部件仍一致", (not byte_same) and parts_same)
    )

    parts = list(expected)
    index = next(i for i, (name, _) in enumerate(parts) if name == "word/document.xml")
    payload = parts[index][1].replace("宋体".encode("utf-8"), "楷体".encode("utf-8"))
    parts[index] = ("word/document.xml", payload)
    caught = diff_parts(expected, parts) != []
    checks.append(("改一个部件的文字会被发现", caught))

    missing = [pair for pair in expected if pair[0] != "word/styles.xml"]
    checks.append(("少一个部件会被发现", diff_parts(expected, missing) != []))

    for label, ok in checks:
        print(("PASS  " if ok else "FAIL  ") + label)
    return 0 if all(ok for _, ok in checks) else 1


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--self-test", action="store_true", help="验证本脚本自身的判据，不读样本"
    )
    args = parser.parse_args()
    if args.self_test:
        return self_test()

    problems = []
    sample_count = 0
    part_count = 0
    for directory, filename, names_attr, builder_name, takes_name in GENERATORS:
        module = load_generator(directory, filename)
        source = getattr(module, names_attr)
        names = list(source) if isinstance(source, (dict, list, tuple)) else [source]
        builder = getattr(module, builder_name)
        for name in names:
            sample_count += 1
            try:
                expected = builder(name) if takes_name else builder()
            except Exception as exc:  # 生成器自身出错也要明确报出来
                problems.append("%s/%s 生成失败：%s" % (directory, name, exc))
                continue
            target = HERE / directory / name
            if not target.exists():
                problems.append("%s/%s 缺失" % (directory, name))
                continue
            expected_parts = read_parts(expected)
            part_count += len(expected_parts)
            for issue in diff_parts(expected_parts, read_parts(target.read_bytes())):
                problems.append("%s/%s：%s" % (directory, name, issue))

    if problems:
        for issue in problems:
            print("不一致：" + issue)
        print("校验失败：%d 处" % len(problems))
        return 1
    print(
        "校验通过：%d 个生成器、%d 份样本、%d 个 ZIP 部件逐部件一致"
        % (len(GENERATORS), sample_count, part_count)
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
