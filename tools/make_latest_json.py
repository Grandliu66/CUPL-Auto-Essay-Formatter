"""生成「检查更新」读取的 `latest.json`（项目根目录）。

内容 = `CHANGELOG.txt` 里**最近一次**更新的正文 + `packaging/download_links.json` 里的下载链接。
版本号、日期取自 `CHANGELOG.txt` 最上面那一条的标题行（`V1.1.4 2026/10/11`），并且**必须**与
`essay_tool/__init__.py` 的 `__version__` 一致——不一致就报错、不生成（防止忘了写日志或版本号
写错，让所有用户收到一条假的"有新版本"）。生成后会用软件里读取这份文件的同一套校验再读一遍。

发版顺序（重要）：**先把新安装包传到网盘、确认分享链接可用，最后才更新并推送 latest.json**。
反过来用户会先收到"有新版本"的提示，点进去却下载不到。

用法：
    python tools/make_latest_json.py            # 生成 latest.json
    python tools/make_latest_json.py --check    # 只检查现有 latest.json 与当前版本/日志是否一致
"""

from __future__ import annotations

import json
import pathlib
import re
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from essay_tool import __version__                              # noqa: E402
from essay_tool.update_check import parse_latest                # noqa: E402

CHANGELOG = ROOT / "CHANGELOG.txt"
LINKS = ROOT / "packaging" / "download_links.json"
OUT = ROOT / "latest.json"

_HEADER = re.compile(r"^V(\d+\.\d+\.\d+)\s+(\d{4})/(\d{1,2})/(\d{1,2})\s*$")


def latest_changelog_entry() -> tuple[str, str, str]:
    """返回 (版本号, YYYY-MM-DD, 正文)——`CHANGELOG.txt` 最上面那一条。"""
    lines = CHANGELOG.read_text(encoding="utf-8-sig").splitlines()
    start = next((i for i, line in enumerate(lines) if _HEADER.match(line.strip())), None)
    if start is None:
        raise SystemExit("CHANGELOG.txt 里找不到形如 “V1.1.4 2026/10/11” 的标题行")
    version, year, month, day = _HEADER.match(lines[start].strip()).groups()
    end = next((i for i in range(start + 1, len(lines)) if _HEADER.match(lines[i].strip())), len(lines))
    body = "\n".join(lines[start + 1:end]).strip()
    if not body:
        raise SystemExit(f"CHANGELOG.txt 里 V{version} 这一条没有正文")
    return version, f"{int(year):04d}-{int(month):02d}-{int(day):02d}", body


def build() -> dict:
    version, date, body = latest_changelog_entry()
    if version != __version__:
        raise SystemExit(f"版本号对不上：CHANGELOG.txt 最上面是 V{version}，"
                         f"essay_tool/__init__.py 是 {__version__}。先把日志写好、版本号改对再生成。")
    links = json.loads(LINKS.read_text(encoding="utf-8"))
    data = {"version": version, "date": date, "changelog": body, "downloads": links}
    info = parse_latest(json.dumps(data, ensure_ascii=False))
    if info is None:
        raise SystemExit("生成的内容过不了软件里的校验（版本号/日期格式，或下载链接不是白名单域名的 https 地址）")
    if len(info.downloads) != len(links):
        dropped = [item["url"] for item in links if item["url"] not in {d.url for d in info.downloads}]
        raise SystemExit(f"这些下载链接会被软件丢弃（不是 https 或不在白名单域名里）：{dropped}")
    return data


def render(data: dict) -> str:
    return json.dumps(data, ensure_ascii=False, indent=2) + "\n"


def main(argv: list[str]) -> int:
    text = render(build())
    if "--check" in argv:
        current = OUT.read_text(encoding="utf-8") if OUT.exists() else None
        if current == text:
            print("latest.json 与当前版本、日志、链接一致。")
            return 0
        print("latest.json 已过期或不存在——重新运行本脚本生成。")
        return 1
    OUT.write_text(text, encoding="utf-8", newline="\n")
    print(f"已生成 {OUT}（版本 {json.loads(text)['version']}）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
