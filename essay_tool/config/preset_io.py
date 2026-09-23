"""预设读写的统一入口——默认存盘格式是 JSON，Excel 只是"边缘产品"。

设计意图（用户 2026-09-14 定的方向）：软件默认从 JSON 读入配置、把配置存
成 JSON（配置列表在 GUI 里就是一份份 JSON 文件）；Excel 仅用于"导出成
Excel 给人可视化查看/修改，改完再导回来"这条边缘路径，不是默认存盘格式。

`load_preset()` 按扩展名分派，引擎侧（`pipeline.run()` / `template.
build_template()`）和 GUI 侧统一只调这一个函数，不用关心某条预设具体是
从 JSON 还是 Excel 来的——这样"预设文件"这个概念对调用方是透明的，以后
要加别的存储格式也只用改这一处。
"""

from __future__ import annotations

import json
import pathlib

from essay_tool.config import excel_io  # noqa: F401  兼容旧的 import 位置
from essay_tool.config.schema import PresetConfig
from essay_tool.config.serialize import preset_from_dict, preset_to_dict


class PresetIOError(Exception):
    """预设文件读写失败（格式不认识、JSON 内容缺字段……）。"""


def load_preset(path: str | pathlib.Path) -> PresetConfig:
    """按扩展名分派：`.json` 走本地 JSON 存盘，`.xlsx` 走 Excel（边缘产品）。"""
    path = pathlib.Path(path)
    suffix = path.suffix.lower()
    if suffix == ".json":
        return load_preset_json(path)
    if suffix == ".xlsx":
        return excel_io.load_preset(path)
    raise PresetIOError(f'不认识的预设文件格式："{path.suffix or path.name}"（应为 .json 或 .xlsx）')


def load_preset_json(path: str | pathlib.Path) -> PresetConfig:
    path = pathlib.Path(path)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise PresetIOError(f"打不开预设文件：{exc}") from exc
    except json.JSONDecodeError as exc:
        raise PresetIOError(f'"{path.name}" 不是合法的 JSON：{exc}') from exc
    try:
        return preset_from_dict(data)
    except (KeyError, ValueError) as exc:
        raise PresetIOError(f'"{path.name}" 内容有问题：{exc}') from exc
    except (TypeError, AttributeError, IndexError) as exc:
        # 不只是「缺字段/值不认识」：整个 data 是列表而不是字典、某一节填成
        # 了数字/字符串、某个列表短了一截……都会抛这几种（审计 2026-09-15
        # 核实：以前它们会以原始异常的形式冒出 GUI，用户看到的是
        # `AttributeError: 'list' object has no attribute 'get'` 这种天书）。
        raise PresetIOError(
            f'"{path.name}" 结构不对（{type(exc).__name__}）：{exc}') from exc


def save_preset_json(cfg: PresetConfig, path: str | pathlib.Path) -> pathlib.Path:
    path = pathlib.Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    data = preset_to_dict(cfg)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def export_to_excel(cfg: PresetConfig, out_path: str | pathlib.Path) -> pathlib.Path:
    """把 cfg 另存成一份 Excel（边缘产品：给用户可视化查看/修改用）。"""
    from essay_tool.config.excel_writer import write_preset
    return write_preset(cfg, out_path)


def import_from_excel(xlsx_path: str | pathlib.Path, json_path: str | pathlib.Path) -> pathlib.Path:
    """从 Excel 读一份预设，转存成 JSON（"从 Excel 导入配置文件"按钮走这个）。"""
    cfg = excel_io.load_preset(xlsx_path)
    return save_preset_json(cfg, json_path)
