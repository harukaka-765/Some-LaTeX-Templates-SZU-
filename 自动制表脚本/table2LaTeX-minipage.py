#!/usr/bin/env python3
"""
table2latex.py — 将 xlsx / csv / tsv 转换为符合规范的 LaTeX 三线表片段

依赖：pandas, openpyxl, chardet
安装：pip install pandas openpyxl chardet

用法：
    python table2latex.py data.xlsx
    python table2latex.py data.xlsx -o ./output/

列名约定：
    "#"后接变量名，单位用"()"标注，全角与半角小括号均可，即：
    "#变量名(单位)"  或  "#变量名（单位）"  →  列头显示为 变量名/单位
    "#变量名"                           →  列头显示为 变量名（无单位）
    若在 "#" 前加 "~" ，则将该列视为纯文本列
    其他格式                            →  报错退出，指出问题列名
    注意 "#" 后紧接变量名，中间没有空格

xlsx 约定：
    每个 sheet 的名称写表格的名称（即LaTeX表格的caption）
    文件名只能含有英文和数字，不支持中文

输出：
    每个 sheet 生成一个独立的 .tex 文件，命名为 文件名_sheet编号.tex
    文件内容为 minipage 环境片段，供 \\input{} 嵌入主文档
    导言区需加载：\\usepackage{tabularray}  \\UseTblrLibrary{booktabs, siunitx}
                  \\usepackage{caption}
"""

import re
import sys
import argparse
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("请安装依赖：pip install pandas openpyxl chardet")

try:
    import chardet
except ImportError:
    sys.exit("请安装依赖：pip install chardet")


# ── 列名解析 ──────────────────────────────────────────────────────────────────

# 匹配 #变量名(单位) 或 #变量名（单位） 或 #变量名
# 若在 #前加 ~，则将该列视为纯文本列
# 括号同时支持全角（）和半角()
COL_PATTERN = re.compile(r'^(~?)#([^(（]+)(?:[（(]([^)）]+)[）)])?$')   


def parse_col_name(col: str) -> tuple[str, str, bool]:
    """
    解析列名，返回 (变量名, 单位, 文本列)
      #变量名(单位)  → (变量名, 单位, False)
      #变量名        → (变量名, "", False)
      ~#变量名(单位) → (变量名, 单位, True)
      ~#变量名       → (变量名, "", True)
      其他           → 报错退出，指出问题列名
    """
    m = COL_PATTERN.match(str(col).strip())
    if not m:
        sys.exit(f'列名约定错误！问题列名：{col}')
    force_text = m.group(1) == "~"
    return m.group(2).strip(), (m.group(3) or "").strip(), force_text


# ── LaTeX 特殊字符转义 ────────────────────────────────────────────────────────

def latex_escape(s: str) -> str:
    """转义文本列中的 LaTeX 特殊字符，反斜杠必须最先处理"""
    s = s.replace("\\", r"\textbackslash{}")
    for ch, repl in [
        ("&",  r"\&"),
        ("%",  r"\%"),
        ("$",  r"\$"),
        ("#",  r"\#"),
        ("_",  r"\_"),
        ("{",  r"\{"),
        ("}",  r"\}"),
        ("~",  r"\textasciitilde{}"),
        ("^",  r"\textasciicircum{}"),
    ]:
        s = s.replace(ch, repl)
    return s


# ── 列类型判断 ────────────────────────────────────────────────────────────────

def is_numeric_col(series: pd.Series) -> bool:
    """
    严格判断：列内所有非空值均为数字才返回 True。
    只要出现任何文字，整列降级为文本列（c 列）。
    """
    non_null = series.dropna()
    if len(non_null) == 0:
        return False
    try:
        pd.to_numeric(non_null, errors="raise")
        return True
    except (ValueError, TypeError):
        return False


def get_table_format(series: pd.Series) -> str:
    """
    推断 siunitx S 列的 table-format 字符串。
    例：整数部分最多 2 位、小数部分最多 3 位 → '2.3'
    含负数时加前缀 '-'，例：'-2.3'
    仅在确认是数值列后调用。
    """
    int_digits = 1
    dec_digits = 0
    has_negative = False

    for v in series.dropna():
        f = float(v)
        if f < 0:
            has_negative = True
        s = f"{abs(f)}"
        parts = s.split(".")
        int_digits = max(int_digits, len(parts[0]))
        if len(parts) > 1:
            stripped = parts[1].rstrip("0") or "0"
            dec_digits = max(dec_digits, len(stripped))

    fmt = f"{int_digits}.{dec_digits}" if dec_digits > 0 else str(int_digits)
    return f"-{fmt}" if has_negative else fmt


# ── 单元格格式化 ──────────────────────────────────────────────────────────────

def format_cell(val, is_numeric: bool, dec_digits: int) -> str:
    """
    将单元格值转为 LaTeX 字符串。
    - 空值：数值列用 {-}（防止 siunitx 将 - 解析为负号），文本列用 -
    - 数值：按统一小数位格式化
    - 文本：转义 LaTeX 特殊字符
    """
    if pd.isna(val):
        return "{-}" if is_numeric else "-"
    if is_numeric:
        f = float(val)
        return f"{f:.{dec_digits}f}" if dec_digits > 0 else str(int(f))
    return latex_escape(str(val))


# ── DataFrame → LaTeX 片段 ────────────────────────────────────────────────────

def df_to_latex(df: pd.DataFrame, caption: str, label: str) -> str:
    cols = list(df.columns)

    # 1. 解析列名（同时校验格式，出错直接退出）
    parsed = [parse_col_name(c) for c in cols]

    # 2. 判断列类型，推断数值格式
    col_meta = []
    for i, c in enumerate(cols):
        varname, unit, force_text = parsed[i]
        numeric = False if force_text else is_numeric_col(df[c])
        if numeric:
            fmt = get_table_format(df[c])
            dec = int(fmt.split(".")[-1]) if "." in fmt else 0
        else:
            fmt = ""
            dec = 0
        col_meta.append({"numeric": numeric, "fmt": fmt, "dec": dec})

    # 3. 构建 colspec
    col_specs = [
        f"S[table-format={m['fmt']}]" if m["numeric"] else "c"
        for m in col_meta
    ]
    colspec = " ".join(col_specs)

    # 4. 构建表头
    # S 列的表头必须用 {} 包裹，防止 siunitx 将其当作数字解析
    header_cells = []
    for i, (varname, unit, _) in enumerate(parsed):
        text = f"{varname}/{unit}" if unit else varname
        header_cells.append(f"{{{text}}}" if col_meta[i]["numeric"] else text)
    header_row = "    " + " & ".join(header_cells) + r" \\"

    # 5. 构建数据行
    data_rows = []
    for _, row in df.iterrows():
        cells = [
            format_cell(row[c], col_meta[i]["numeric"], col_meta[i]["dec"])
            for i, c in enumerate(cols)
        ]
        data_rows.append("    " + " & ".join(cells) + r" \\")

    # 6. 拼装完整片段
    lines = [
        r"% 导言区需加载：",
        r"% \usepackage{tabularray}",
        r"% \UseTblrLibrary{booktabs, siunitx}",
        r"% \usepackage{caption}  % 提供 \captionof 命令",
        "",
        r"\noindent",
        r"\vspace{0.5em}",        
        r"\begin{minipage}{\linewidth}",
        r"  \centering",
        f"  \\captionof{{table}}{{{caption}}}",
        f"  \\label{{{label}}}",
        f"  \\begin{{tblr}}{{colspec = {{{colspec}}}}}",
        r"    \toprule[1pt]",
        header_row,
        r"    \midrule[0.5pt]",
        *data_rows,
        r"    \bottomrule[1pt]",
        r"  \end{tblr}",
        r"\vspace{0.5em}",          
        r"\end{minipage}",
    ]

    return "\n".join(lines)


# ── 文件读取 ──────────────────────────────────────────────────────────────────

def detect_encoding(path: Path) -> str:
    """用 chardet 自动检测文件编码，检测失败时回退到 utf-8"""
    raw = path.read_bytes()
    result = chardet.detect(raw)
    return result["encoding"] or "utf-8"


def load_sheets(path: Path) -> list[dict]:
    """
    统一入口，返回 list of {df, caption, sheet_name}
    - xlsx：从 Sheet 名称读 caption
    - csv/tsv：caption 使用数据表文件名
    """
    suffix = path.suffix.lower()
    stem = path.stem

    if suffix in (".xlsx", ".xls"):
        try:
            import openpyxl
        except ImportError:
            sys.exit("请安装依赖：pip install openpyxl")

        wb = openpyxl.load_workbook(path)
        sheets = []
        for i, sheet_name in enumerate(wb.sheetnames, start=1):  
            caption = latex_escape(sheet_name)                    
            df = pd.read_excel(path, sheet_name=sheet_name, header=0)
            sheets.append({
                "df": df,
                "caption": caption,
                "sheet_name": f"sheet{i}",                        
            })
        return sheets

    elif suffix == ".csv":
        enc = detect_encoding(path)
        df = pd.read_csv(path, sep=None, engine="python", encoding=enc)
        return [{"df": df, "caption": latex_escape(stem), "sheet_name": "sheet1"}]

    elif suffix in (".tsv", ".txt"):
        enc = detect_encoding(path)
        df = pd.read_csv(path, sep="\t", encoding=enc)
        return [{"df": df, "caption": latex_escape(stem), "sheet_name": "sheet1"}]

    else:
        sys.exit(f"不支持的文件格式：{suffix}（支持 xlsx / xls / csv / tsv）")


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="将数据表（xlsx/csv/tsv）转换为 LaTeX 三线表片段"
    )
    parser.add_argument("file",
                        help="输入文件路径")
    parser.add_argument("-o", "--output", default=".",
                        help="输出目录（默认：当前目录）")
    args = parser.parse_args()

    path = Path(args.file)
    if not path.exists():
        sys.exit(f"文件不存在：{path}")

    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    sheets = load_sheets(path)
    stem = path.stem

    for s in sheets:
        label = f"tab:{stem}_{s['sheet_name']}"
        out_name = f"{stem}_{s['sheet_name']}.tex"
        latex = df_to_latex(s["df"], s["caption"], label)

        out_path = out_dir / out_name
        out_path.write_text(latex, encoding="utf-8")
        print(f"✅ 已生成：{out_path}")


if __name__ == "__main__":
    main()