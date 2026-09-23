"""GUI 里几个"记住我选过什么"的小设置——目前只有备份相关的几项。

这跟预设（排版规则，Excel/JSON、内置 vs 用户自建）不是一回事：这里存的是
"软件本身怎么表现"的偏好（要不要自动备份、备份放哪、留几份），所以单独
一份文件、单独一个 dataclass，不跟 PresetConfig 混在一起。落盘位置固定用
`user_data_dir()`——这是每用户可写目录，跟预设的用户目录同一个道理。
"""

from __future__ import annotations

import json
import pathlib
from dataclasses import asdict, dataclass, fields

from essay_tool.config.paths import user_data_dir

SETTINGS_FILE_NAME = "settings.json"


@dataclass
class GuiSettings:
    backup_enabled: bool = True
    # None = 用 paths.import_backup_dir() 那个默认位置；填了字符串就是用户
    # 自己用「设置备份路径……」选的目录，覆盖默认值。
    backup_dir: str | None = None
    backup_auto_prune: bool = False
    # backup_keep_unit 决定 backup_keep_count 这个数字怎么解释：
    # "条" = 只留最近 N 份备份（按修改时间）；"天"/"周"/"月" = 删掉比 N
    # 个单位更早的备份。见 essay_tool.engine.loader.prune_backups。
    # 默认 4 周（用户 2026-09-14 拍板的默认值，勾选"自动删除过早备份"时
    # 第一次看到的就是这个）。
    backup_keep_count: int = 4
    backup_keep_unit: str = "周"


def _settings_path() -> pathlib.Path:
    return user_data_dir() / SETTINGS_FILE_NAME


def load_settings() -> GuiSettings:
    """读不到/读出来的东西不对，都退回默认值，不让设置文件本身变成故障点。"""
    path = _settings_path()
    if not path.exists():
        return GuiSettings()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return GuiSettings()
    if not isinstance(data, dict):
        # 文件里是数组/字符串/数字（被手工改坏了）：`data.items()` 会抛
        # AttributeError，而这里是**启动期**调用，异常会一路冒到崩溃弹窗。
        # 上面那句 docstring 承诺”读出来的东西不对都退回默认值”，照做。
        return GuiSettings()
    known = {f.name for f in fields(GuiSettings)}
    filtered = {k: v for k, v in data.items() if k in known}
    # ⚠️ 只按字段名过滤不够——只要字段名对得上，**类型不管对不对都会被塞进去**
    # （比如 `backup_dir` 被手改成数字、`backup_keep_count` 被改成字符串）。
    # `GuiSettings` 是个普通 dataclass，赋值时不做类型检查，坏类型的字段会
    # 原样构造成功，直到真被用到的地方（拼路径、比大小……）才炸，而且每次
    # 启动都读同一份坏文件，就成了”每次一开程序就崩”（2026-09-22 审计揪出的
    # 真 bug）。这里按每个字段”应该是什么类型”逐个校验，不对的就丢掉、退回
    # 该字段的默认值，不让一个坏字段拖累整个设置文件读不进来。
    defaults = GuiSettings()
    checked: dict = {}
    for name, value in filtered.items():
        default_value = getattr(defaults, name)
        if name == "backup_dir":
            ok = value is None or isinstance(value, str)
        elif name in ("backup_enabled", "backup_auto_prune"):
            ok = isinstance(value, bool)
        elif name == "backup_keep_count":
            ok = isinstance(value, int) and not isinstance(value, bool)
        elif name == "backup_keep_unit":
            ok = isinstance(value, str)
        else:
            ok = isinstance(value, type(default_value)) if default_value is not None else True
        checked[name] = value if ok else default_value
    return GuiSettings(**checked)


def save_settings(settings: GuiSettings) -> None:
    path = _settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(asdict(settings), ensure_ascii=False, indent=2),
                     encoding="utf-8")
