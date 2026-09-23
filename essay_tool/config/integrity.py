"""随程序分发的官方文件的"出厂校验"：文件被改动或损坏时拒绝使用。

哪些文件：封面要从三份官方 docx 里抽（论文写作模板、课程考核封面、十星封面
范例）；《规范》docx 与十星格式要求 PDF 用来查阅和"另存官方版"。这些文件装在
安装目录里，用户有能力改它们——改坏了封面会排错，"另存官方版"还会把一份
被动过的文件当成官方版递出去。

做法：`integrity_hashes.py` 里存着这几份文件出厂时的 SHA-256（由
`tools/update_integrity.py` 生成），用之前重算一遍对比，不一致就抛
`IntegrityError`。**不按文件名分角色**——只看"这份文件的哈希在不在出厂清单里"，
所以十星封面范例文件名里的届数每年换，也只要重跑那个脚本更新清单。

这不是防恶意攻击（程序本身也能被改），只是防误改、防损坏。
"""

from __future__ import annotations

import hashlib
import pathlib

from essay_tool.config.integrity_hashes import OFFICIAL_SHA256


# 官方文件缺失或被改动时给用户看的话（用户 2026-09-20 定稿；缺失与校验失败共用）
OFFICIAL_FILE_PROBLEM = "/reference内丢失对应文件，可能已被改动或者损坏。"


class IntegrityError(Exception):
    """官方文件与出厂版本不一致。文案面向用户，GUI 直接显示。"""


def sha256_of(path: str | pathlib.Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def verify_official(path: str | pathlib.Path) -> pathlib.Path:
    """校验通过就把路径原样返回（方便串在 `Document(...)` 里）；否则抛
    `IntegrityError`。文件不存在时抛的是普通 `FileNotFoundError`，不算校验失败。
    """
    path = pathlib.Path(path)
    if sha256_of(path) not in OFFICIAL_SHA256:
        raise IntegrityError(OFFICIAL_FILE_PROBLEM)
    return path
