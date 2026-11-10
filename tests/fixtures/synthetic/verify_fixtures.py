"""校验已提交的合成样本：逐部件重建 + 结构断言 + 清单一致性。

三类判据分开算，不混为一谈：

  * 逐部件重建：导入零依赖生成器在内存里重建样本，比较 ZIP 条目名/顺序与
    每个条目解压后的字节，忽略 deflate 流、时间戳、外部属性、注释。这样在
    Python / zlib 版本不同的机器上不会因压缩流差异误报。
  * 结构断言：format-samples/ 的重建依赖 python-docx，干净检出无法重建，
    因此只调用它零依赖的 verify() 逐项核对 OOXML 维度。这**不是**逐字节
    复现，也不等于格式合规结论。
  * 清单一致性：manifest.samples / 生成器清单 / 磁盘上的 .docx 三方比对，
    抓「新增样本忘写 manifest」「孤儿 docx」「生成器没覆盖」这类漏项。

用法（从仓库根目录）：
    python3 tests/fixtures/synthetic/verify_fixtures.py
    python3 tests/fixtures/synthetic/verify_fixtures.py --self-test
"""
import argparse
import importlib.util
import json
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

# (目录, 生成器文件, 结构校验函数)；只做维度断言，不在内存重建 docx
STRUCTURE_VERIFIERS = [
    ("format-samples", "generate.py", "verify"),
]

# (目录, 是否比对生成器清单)
# full-thesis 的两份 docx 由命令行 --output 指定，生成器没有固定清单常量，
# 因此只比 manifest ↔ 磁盘；这一点在输出里明说，不假装也查了生成器。
LIST_SPECS = [
    ("document", True),
    ("formula", True),
    ("demo", True),
    ("format-samples", True),
    ("full-thesis", False),
]


def load_generator(directory, filename):
    path = HERE / directory / filename
    module_name = "_fixture_gen_" + directory + "_" + Path(filename).stem
    if module_name in sys.modules:
        return sys.modules[module_name]
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


def diff_lists(declared, generated, disk, check_generated=True):
    """纯函数：比 manifest 声明、生成器清单、磁盘文件三个集合。"""
    problems = []
    for name in sorted(disk - declared):
        problems.append("未登记样本：%s（磁盘有，manifest 无）" % name)
    for name in sorted(declared - disk):
        problems.append("manifest 声明缺失：%s" % name)
    if check_generated:
        for name in sorted(disk - generated):
            problems.append("生成器未覆盖：%s" % name)
        for name in sorted(generated - disk):
            problems.append("生成器声明但磁盘无：%s" % name)
    return problems


def generator_samples():
    """返回 {目录: 生成器能产出的样本文件名集合}。"""
    result = {}
    for directory, filename, names_attr, _, _ in GENERATORS:
        source = getattr(load_generator(directory, filename), names_attr)
        names = list(source) if isinstance(source, (dict, list, tuple)) else [source]
        result.setdefault(directory, set()).update(Path(str(n)).name for n in names)
    fs = load_generator("format-samples", "generate.py")
    result.setdefault("format-samples", set()).update(n + ".docx" for n in fs.SAMPLES)
    return result


def manifest_samples(directory):
    path = HERE / directory / "manifest.json"
    if not path.exists():
        return None
    data = json.loads(path.read_text(encoding="utf-8"))
    samples = data.get("samples")
    if not isinstance(samples, list):
        return None
    return {s["file"] if isinstance(s, dict) else s for s in samples}


def check_lists():
    """返回 (问题描述列表, 磁盘 docx 总数)。"""
    problems = []
    generated = generator_samples()
    disk_total = 0
    for directory, compare_generated in LIST_SPECS:
        folder = HERE / directory
        disk = {p.name for p in folder.glob("*.docx")}
        disk_total += len(disk)
        declared = manifest_samples(directory)
        if declared is None:
            problems.append("%s/manifest.json 缺少 samples 清单" % directory)
            continue
        owned = generated.get(directory, set())
        for issue in diff_lists(declared, owned, disk, compare_generated):
            problems.append("%s：%s" % (directory, issue))
    return problems, disk_total


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
    """证明比较判据既不误报也不漏报；不读任何样本文件。"""
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
    checks.append(("改一个部件的文字会被发现", diff_parts(expected, parts) != []))

    missing = [pair for pair in expected if pair[0] != "word/styles.xml"]
    checks.append(("少一个部件会被发现", diff_parts(expected, missing) != []))

    one = {"a.docx"}
    checks.append(("清单三方一致时无问题", diff_lists(one, one, one) == []))
    checks.append((
        "manifest 漏声明会被发现",
        any("未登记样本" in p for p in diff_lists(set(), one, one)),
    ))
    checks.append((
        "manifest 声明了不存在的样本会被发现",
        any("声明缺失" in p for p in diff_lists({"a.docx", "b.docx"}, one, one)),
    ))
    checks.append((
        "生成器漏覆盖会被发现",
        any("生成器未覆盖" in p for p in diff_lists(one, set(), one)),
    ))
    checks.append((
        "关闭生成器比对时不误报",
        diff_lists(one, set(), one, check_generated=False) == [],
    ))

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

    structure_count = 0
    for directory, filename, func_name in STRUCTURE_VERIFIERS:
        module = load_generator(directory, filename)
        structure_count += len(module.SAMPLES)
        for issue in getattr(module, func_name)(HERE / directory):
            problems.append("%s：%s" % (directory, issue))

    list_problems, disk_total = check_lists()
    problems.extend(list_problems)

    if problems:
        for issue in problems:
            print("不一致：" + issue)
        print("校验失败：%d 处" % len(problems))
        return 1
    print(
        "校验通过：逐部件重建 %d 个生成器 / %d 份样本 / %d 个部件；"
        "结构校验 %d 组 / %d 份样本；清单一致性 %d 个目录 / %d 份 docx"
        % (len(GENERATORS), sample_count, part_count,
           len(STRUCTURE_VERIFIERS), structure_count, len(LIST_SPECS), disk_total)
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
