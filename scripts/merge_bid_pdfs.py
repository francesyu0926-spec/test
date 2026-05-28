# -*- coding: utf-8 -*-
"""Merge bid PDFs in a fixed order."""

import sys
from pathlib import Path

ORDER = [
    "封面.pdf",
    "投标函及投标函附录.pdf",
    "法定代表人身份证明或授权委托书.pdf",
    "投标保证金.pdf",
    "资格审查资料.pdf",
    "施工组织设计.pdf",
    "项目管理机构表.pdf",
    "拟分包项目情况表.pdf",
    "已标价工程量清单.pdf",
    "原件的复印件.pdf",
    "其它材料.pdf",
]


def merge_pdfs(src_dir: Path, out_name: str) -> Path:
    src_dir = src_dir.resolve()
    out_path = src_dir / f"{out_name}.pdf"

    import fitz

    out_doc = fitz.open()
    merged = []
    for name in ORDER:
        p = src_dir / name
        if not p.exists():
            raise FileNotFoundError(p)
        src = fitz.open(p)
        out_doc.insert_pdf(src)
        src.close()
        merged.append(name)

    out_doc.save(out_path)
    out_doc.close()
    print(f"Merged {len(merged)} files -> {out_path}")
    print(f"Size: {out_path.stat().st_size} bytes")
    return out_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python merge_bid_pdfs.py <dir> [output_basename]")
        sys.exit(1)
    directory = Path(sys.argv[1])
    basename = sys.argv[2] if len(sys.argv) > 2 else "山西吉锐水利水电工程有限公司投标文件"
    merge_pdfs(directory, basename)
