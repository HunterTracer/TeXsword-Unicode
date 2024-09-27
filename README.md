# TeXsword-Unicode

TeXsword-Unicode is a macro package for Microsoft Word that **allows one can use Unicode Characters in TeXsword (such as Greek letters, Chinese characters, etc.)** and use LaTeX in Microsoft Word. The original version of TeXsword can be found at https://sourceforge.net/projects/texsword/.

## Installation

1. Download `texsword.dotm` in this repository and copy it into the StartUp directory of MSWord (probably, in `C:\Documents and Settings\<your_user_name>\Application Data\Microsoft\Word\STARTUP` or `C:\Users\<your_user_name>\AppData\Roaming\Microsoft\word\STARTUP`).
2. Download `Ghostscript` from https://www.ghostscript.com/releases/gsdnld.html and install. Append the directory containing `gswin64c.exe` (if 64-bit system) or `gswin32c.exe` (if 32-bit system) to the environment variable `PATH`.
3. Open Microsoft Word, click `TeXsword options`. Make sure that `latex executable name` is `xelatex` and `dvipng name` is `gswin64c` (if 64-bit system) or `gswin32c` (if 32-bit system).

## Testing

You can use the following code to test whether TeXsword is OK.

```
\documentclass{article}
\usepackage{xeCJK}
\pagestyle{empty}
\begin{document}
	\section{前言}
	这是一些文字。
	
	\section{关于数学部分}
	数学、中英文皆可以混排。You can intersperse math, Chinese and English (Latin script) without adding extra environments.
	
	這是繁體中文。
\end{document}
```

## Difference vs Original TeXsword

1. The original `DumpStringToTextFile` in `texsword_util` results in ANSI encoding which is rewritten to support UTF-8 encoding.
2. In `texsword_runlatex`, use `xelatex` to compile .tex and then generate the .pdf file to support UTF-8 encoding..
3. In `texsword_runlatex`, use `Ghostscript` to generate the .bbx file and then convert .pdf to .png.
