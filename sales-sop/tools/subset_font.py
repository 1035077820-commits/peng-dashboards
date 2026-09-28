"""为销售跟单平台生成 HarmonyOS Sans SC 的 woff2 子集。

字集范围：ASCII + 拉丁补充 + CJK 标点 + 全角符号 + GB2312 一二级汉字(6763字)。
覆盖日常中文姓名、公司名、产品名，缺失字符由系统字体兜底。
"""
import os
import subprocess
import sys

SRC_DIR = r"\\李晓\公司\平面设计\VI文件\捷展\2024\2024新vi\1.捷展QuickShow\品牌字体\中文字体"
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "fonts")

WEIGHTS = [
    ("HarmonyOS_Sans_SC_Light.ttf", "harmonyos-sans-sc-300.woff2"),
    ("HarmonyOS_Sans_SC_Medium.ttf", "harmonyos-sans-sc-500.woff2"),
    ("HarmonyOS_Sans_SC_Bold.ttf", "harmonyos-sans-sc-700.woff2"),
]


def charset():
    chars = set()

    # ASCII 可打印
    for c in range(0x20, 0x7F):
        chars.add(chr(c))
    # 拉丁补充 / 常用符号
    for c in range(0xA0, 0x100):
        chars.add(chr(c))
    for c in range(0x2000, 0x2070):
        chars.add(chr(c))
    for c in range(0x20A0, 0x20C0):
        chars.add(chr(c))
    for c in range(0x2100, 0x2140):
        chars.add(chr(c))
    for c in range(0x2190, 0x21A0):
        chars.add(chr(c))
    for c in range(0x2200, 0x2230):
        chars.add(chr(c))
    for c in range(0x2460, 0x24A0):
        chars.add(chr(c))
    for c in range(0x25A0, 0x25C0):
        chars.add(chr(c))
    # CJK 标点 + 全角
    for c in range(0x3000, 0x3040):
        chars.add(chr(c))
    for c in range(0xFF00, 0xFFF0):
        chars.add(chr(c))
    # GB2312 全部可解码字符
    for hi in range(0xA1, 0xF8):
        for lo in range(0xA1, 0xFF):
            try:
                chars.add(bytes([hi, lo]).decode("gb2312"))
            except UnicodeDecodeError:
                pass
    chars.discard("\ufffe")
    return "".join(sorted(chars))


def main():
    out_dir = os.path.abspath(OUT_DIR)
    os.makedirs(out_dir, exist_ok=True)
    text = charset()
    print(f"charset size: {len(text)}")

    txt_path = os.path.join(out_dir, "_charset.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(text)

    total = 0
    for src_name, out_name in WEIGHTS:
        src = os.path.join(SRC_DIR, src_name)
        if not os.path.exists(src):
            print(f"  ! missing {src_name}")
            continue
        dst = os.path.join(out_dir, out_name)
        cmd = [
            sys.executable, "-m", "fontTools.subset", src,
            f"--text-file={txt_path}",
            "--output-file=" + dst,
            "--flavor=woff2",
            "--layout-features=*",
            "--no-hinting",
            "--desubroutinize",
            "--drop-tables+=DSIG",
            "--name-IDs=*",
        ]
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode != 0:
            print(f"  ! failed {src_name}: {r.stderr[-500:]}")
            continue
        size = os.path.getsize(dst)
        total += size
        print(f"  {out_name:38s} {size/1024:8.1f} KB")

    os.remove(txt_path)
    print(f"total: {total/1024/1024:.2f} MB")


if __name__ == "__main__":
    main()
