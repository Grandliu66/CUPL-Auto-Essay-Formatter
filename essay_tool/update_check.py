"""检查更新：只读一个很小的 `latest.json`，告诉用户有没有新版本、更新了什么、去哪下载。

**只通知，不下载、不安装**：下载链接只是在用户的浏览器里打开对应页面（网盘、GitHub）。
本模块不上传任何东西——没有文档、没有用户数据，请求里只有一个普通的 GET（带一个
`User-Agent: CUPL-UTAF/版本号`，服务器能看到的就是任何网页访问都会带的 IP）。联网**只发生在**
检查更新这一件事上，用户协议第一条已经写明。

`latest.json` 的格式（由 `tools/make_latest_json.py` 从 `CHANGELOG.txt` 与
`packaging/download_links.json` 生成）：

    {"version": "1.1.4", "date": "2026-10-11", "changelog": "多行文字",
     "downloads": [{"name": "夸克网盘", "url": "https://pan.quark.cn/s/…"}, …]}

多个源按顺序试（国内能读到的在前），第一个读成功且格式合格的就用；全都失败返回 `None`
（**静默**：自动检查时不打扰用户；手动检查由界面决定怎么说）。

**这个文件来自网络，当它不可信**：
  * 最多读 256 KB，超时 6 秒；
  * 版本号必须是 `数字.数字.数字`；日志/名称限长；
  * 下载链接**只留 https 且域名在白名单里的**（见 `ALLOWED_HOSTS`），其余丢掉——
    即使文件被人改了，软件也不会借用户的浏览器去打开别的地方；
  * 一个能用的链接都没有就当这份文件无效。
"""

from __future__ import annotations

import json
import re
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field

# 按顺序试：Gitee（国内多数网络能访问）→ GitHub。地址一旦随软件发出去就改不了，
# 所以写稳定的仓库地址；要换下载链接改 `latest.json` 里的 downloads，不用改这里。
SOURCES: tuple[str, ...] = (
    "https://gitee.com/grandliu66/cupl-utaf-auto-update-check/raw/master/latest.json",
    "https://raw.githubusercontent.com/Grandliu66/CUPL-Auto-Essay-Formatter/main/latest.json",
)

# 下载链接允许打开的域名（含其子域）。
ALLOWED_HOSTS: tuple[str, ...] = (
    "pan.quark.cn", "pan.baidu.com", "github.com", "gitee.com",
)

TIMEOUT_SECONDS = 6
MAX_BYTES = 256 * 1024
_MAX_CHANGELOG_CHARS = 8000
_MAX_NAME_CHARS = 40
_MAX_LINKS = 8
_VERSION_RE = re.compile(r"^\d{1,3}\.\d{1,3}\.\d{1,3}$")
_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


@dataclass(frozen=True)
class DownloadLink:
    name: str
    url: str


@dataclass(frozen=True)
class LatestInfo:
    version: str
    date: str
    changelog: str
    downloads: tuple[DownloadLink, ...] = field(default_factory=tuple)

    def display_date(self) -> str:
        return self.date.replace("-", "/")


def version_tuple(text: str) -> tuple[int, ...]:
    """`"1.1.4"` → `(1, 1, 4)`。不合格的写法抛 ValueError（调用方当"无法比较"处理）。"""
    if not _VERSION_RE.match(text or ""):
        raise ValueError(f"不是 x.y.z 形式的版本号：{text!r}")
    return tuple(int(part) for part in text.split("."))


def is_newer(latest: str, current: str) -> bool:
    try:
        return version_tuple(latest) > version_tuple(current)
    except ValueError:
        return False


def is_allowed_url(url: str) -> bool:
    """只放行 https + 白名单域名（含子域）；用户名密码、非常规端口一律不要。"""
    try:
        parts = urllib.parse.urlsplit(url)
    except ValueError:
        return False
    if parts.scheme != "https" or not parts.hostname:
        return False
    if parts.username or parts.password or parts.port not in (None, 443):
        return False
    host = parts.hostname.lower()
    return any(host == allowed or host.endswith("." + allowed) for allowed in ALLOWED_HOSTS)


def parse_latest(raw: bytes | str) -> LatestInfo | None:
    """把读到的内容解析成 `LatestInfo`；任何一处不合格都返回 None（不抛异常）。"""
    try:
        text = raw.decode("utf-8-sig") if isinstance(raw, bytes) else raw
        data = json.loads(text)
    except (UnicodeDecodeError, ValueError):
        return None
    if not isinstance(data, dict):
        return None
    version, date, changelog = data.get("version"), data.get("date", ""), data.get("changelog", "")
    if not isinstance(version, str) or not _VERSION_RE.match(version):
        return None
    if not isinstance(date, str) or (date and not _DATE_RE.match(date)):
        return None
    if not isinstance(changelog, str):
        return None
    links: list[DownloadLink] = []
    raw_links = data.get("downloads", [])
    if not isinstance(raw_links, list):
        return None
    for item in raw_links[:_MAX_LINKS]:
        if not isinstance(item, dict):
            continue
        name, url = item.get("name"), item.get("url")
        if not isinstance(name, str) or not isinstance(url, str) or not name.strip():
            continue
        if is_allowed_url(url):
            links.append(DownloadLink(name.strip()[:_MAX_NAME_CHARS], url))
    if not links:
        return None                                    # 没有一个能用的链接：这份文件对用户没用
    return LatestInfo(version, date, changelog.strip()[:_MAX_CHANGELOG_CHARS], tuple(links))


def _fetch(url: str, current_version: str) -> bytes | None:
    request = urllib.request.Request(
        url, headers={"User-Agent": f"CUPL-UTAF/{current_version}", "Accept": "application/json,*/*"})
    try:
        with urllib.request.urlopen(request, timeout=TIMEOUT_SECONDS) as response:   # 默认会跟随跳转
            body = response.read(MAX_BYTES + 1)
    except (urllib.error.URLError, OSError, ValueError):
        return None
    return body if len(body) <= MAX_BYTES else None


def fetch_latest(current_version: str, sources: tuple[str, ...] | None = None) -> LatestInfo | None:
    """依次试各个源，返回第一份合格的 `LatestInfo`；全部失败返回 None。**会阻塞**，
    界面要在后台线程里调。`sources` 只给测试用。"""
    for url in (SOURCES if sources is None else sources):
        body = _fetch(url, current_version)
        if body is None:
            continue
        info = parse_latest(body)
        if info is not None:
            return info
    return None
