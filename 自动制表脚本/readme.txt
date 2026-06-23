table2latex.py — 将 xlsx / csv / tsv 转换为符合学术规范的 LaTeX 三线表片段

-float后缀的脚本生成的表格使用浮动体机制插入
-minipage后缀的脚本生成的表格使用小页面机制插入

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
    文件内容为 table 环境片段，供 \\input{} 嵌入主文档
    导言区需加载：\\usepackage{tabularray}  \\UseTblrLibrary{booktabs, siunitx}