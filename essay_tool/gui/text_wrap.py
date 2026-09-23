"""日志区的中英混排换行（用户 2026-09-15 要求）。

`tk.Text` 自带的两种换行方式都不合用：

* `wrap="word"` 只在空格处断行——中文句子没有空格，整句被当成一个"单词"，
  一行放不下时它会在**任意字符**处硬折。实测日志里出现的就是
  `[12:10:01] 第 13` / `段｜目录：标签下方出现正文；……`——看着像排版坏了；
* `wrap="char"` 能写满一行，但会**把标点甩到行首**、**把英文单词拆开**。

所以自己按行宽算断开位置，只做三条必要的控制（用户原话："一行写满，再换
下一行，只做必要的换行控制"）：

1. 一行写满再换下一行；
2. **收尾类标点不留行首**（。，、；：！？）】》…）——它跟着上一个片段一起
   挪到下一行，宁可让上一行短一点；同理**开引号/开括号不留行尾**；
3. **英文单词不拆开**（字母/数字/`@._'-/` 组成的连续串视为整体）——放不下
   就整词挪到下一行；单词本身比一整行还宽时才按字符硬拆（无路可退）。

`measure` 是"这段文字有多宽"的函数（真实使用时传 `tkinter.font.Font.measure`），
这样本模块不依赖 Tk，可以单独测。
"""

from __future__ import annotations

from collections.abc import Callable

# 不能出现在**行首**的字符（收尾类标点、后括号、后引号、省略号等）
_NO_LINE_START = "。，、；：！？）〕］｝〉》」』】〗”’〞…‥—～%‰℃·："
# 不能出现在**行尾**的字符（开括号、开引号）
_NO_LINE_END = "（〔［｛〈《「『【〖“‘"
# 组成"一个英文单词"的字符（数字、字母、以及常见的词内符号）
_WORD_EXTRA = "@._'’-/"


def _is_word_char(ch: str) -> bool:
    return ch.isascii() and (ch.isalnum() or ch in _WORD_EXTRA)


def tokenize(text: str) -> list[str]:
    """切成"不可再分的最小片段"：英文单词整块、其余单字符。

    注意标点不算词：`word，` 会切成 `word` + `，`，这样"不拆单词"和
    "标点不落行首"两条规则互不打架。
    """
    tokens: list[str] = []
    i = 0
    total = len(text)
    while i < total:
        ch = text[i]
        if _is_word_char(ch):
            j = i
            while j < total and _is_word_char(text[j]):
                j += 1
            tokens.append(text[i:j])
            i = j
        else:
            tokens.append(ch)
            i += 1
    return tokens


def wrap_text(text: str, measure: Callable[[str], float], width: float) -> list[str]:
    """把一段文字按 `width`（像素）折成多行。返回行列表（不含换行符）。"""
    if width <= 0:
        return [text]
    lines: list[str] = []
    current = ""
    current_w = 0.0

    for token in tokenize(text):
        token_w = measure(token)
        if not current:
            current, current_w = token, token_w
            continue
        if current_w + token_w <= width:
            current += token
            current_w += token_w
            continue

        # —— 放不下了，准备换行，先处理两条避让规则
        if token and token[0] in _NO_LINE_START:
            # 收尾标点不能单独去下一行：跟本行一起留下（宁可略微超宽）
            current += token
            current_w += token_w
            continue
        if len(token) == 1 and token in _NO_LINE_END:
            # 开括号不能留在行尾：把它连同后面的内容一起挪下去
            lines.append(current.rstrip())
            current, current_w = token, token_w
            continue

        lines.append(current.rstrip())
        # 单词比整行还宽：只能按字符硬拆（无路可退）
        if token_w > width:
            piece = ""
            piece_w = 0.0
            for ch in token:
                ch_w = measure(ch)
                if piece and piece_w + ch_w > width:
                    lines.append(piece)
                    piece, piece_w = ch, ch_w
                else:
                    piece += ch
                    piece_w += ch_w
            current, current_w = piece, piece_w
        else:
            current, current_w = token, token_w

    if current.strip():
        lines.append(current.rstrip())
    return lines or [""]
