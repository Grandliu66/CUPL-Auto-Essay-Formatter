"""把用户提交的 docx/txt 变成一份可以放心修改的工作副本。

docx 直接复制（保留 docProps、修订记录等元数据，不解压重建），
txt 新建一个 docx，把每一行原样拷进去（后续阶段再识别标题、结构）。
"""

from __future__ import annotations

import json
import pathlib
import shutil
import time

from docx import Document

TXT_ENCODINGS = ("utf-8-sig", "utf-8", "gbk")


class LoadError(Exception):
    pass


def _read_text(path: pathlib.Path) -> str:
    last_err = None
    for enc in TXT_ENCODINGS:
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError as e:
            last_err = e
    raise LoadError(f"无法识别文本编码（试过 {'/'.join(TXT_ENCODINGS)}）：{path}") from last_err


def _same_file(a: pathlib.Path, b: pathlib.Path) -> bool:
    """两个路径是不是同一个文件（拿不到信息时当作"不是"）。"""
    try:
        return a.samefile(b)
    except OSError:
        return False


def load_as_working_copy(source: str | pathlib.Path, work_dir: str | pathlib.Path) -> pathlib.Path:
    """返回工作副本 docx 的路径。docx 源文件原样复制；txt 源文件新建 docx 逐行写入。"""
    source = pathlib.Path(source)
    work_dir = pathlib.Path(work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)

    suffix = source.suffix.lower()
    if suffix == ".docx":
        target = work_dir / source.name
        if _same_file(source, target):
            # 源文件**本来就在工作目录里**（GUI 的工作目录恒为 `工作区\work`，
            # 而"导出空白模板"的默认落地目录也是那儿）：`shutil.copy2(source, target)`
            # 就成了"复制到自己"，Windows 上直接 WinError 32
            # （PermissionError: 另一个程序正在使用此文件），用户看到的只是
            # 状态栏一句"生成论文失败：（WinError 32）…"（2026-09-16 审计实测复现）。
            #
            # 这里改成一个另起的名字，不能就地当工作副本：管线**全程会改这份
            # 副本**，而"从不修改用户的原始文件"是这个工具的硬约定
            # （见模块开头），导入的文件就在 work 里也一样得守着。
            for n in range(0, 1000):
                tag = "" if n == 0 else str(n)
                candidate = work_dir / f"{source.stem}-工作副本{tag}{source.suffix}"
                if not candidate.exists():
                    target = candidate
                    break
            else:
                # 名字用尽时必须报错：以前没有这个 else，循环跑完 target 仍是
                # 原来的 `work_dir / source.name`（＝源文件本身），“复制到自己”
                # 的 WinError 32 又回来了（审计 2026-09-16 报的）。
                raise LoadError(
                    f"工作目录里已经有 1000 份「{source.stem}-工作副本」，"
                    f"请先清理 {work_dir} 再试（工具不会就地修改你导入的文件）")
        shutil.copy2(source, target)
        return target

    if suffix == ".txt":
        text = _read_text(source)
        doc = Document()
        for line in text.splitlines():
            doc.add_paragraph(line)
        target = work_dir / (source.stem + ".docx")
        doc.save(target)
        return target

    raise LoadError(f"不支持的文件类型：{source.suffix}（只支持 .docx 和 .txt）")


_MANIFEST_NAME = ".auto_backup_manifest.json"


def _load_backup_manifest(backup_dir: pathlib.Path) -> set[str]:
    """读这份目录下"哪些文件是我们自己建的"名单，读不到就当作空的。

    这份名单是"自动清理只删自己建的备份，绝不动用户放在这个目录里的其他
    文件"这条安全规则的依据——**不能靠文件名规律或者修改时间去猜"这个
    文件像不像我们建的"，猜错了会误删用户自己的东西**（用户设置了自定义
    备份路径、那个目录里本来就有别的文件时，这条尤其要紧）。
    """
    path = backup_dir / _MANIFEST_NAME
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return set(data.get("managed_files", []))
    except (OSError, json.JSONDecodeError, AttributeError):
        return set()


def _save_backup_manifest(backup_dir: pathlib.Path, names: set[str]) -> None:
    path = backup_dir / _MANIFEST_NAME
    try:
        path.write_text(
            json.dumps({"managed_files": sorted(names)}, ensure_ascii=False, indent=2),
            encoding="utf-8")
    except OSError:
        pass


def backup_import(source: str | pathlib.Path, backup_dir: str | pathlib.Path) -> pathlib.Path:
    """导入时把原文件复制一份到 backup_dir，返回实际落盘的路径。

    这是"用户手滑删了自己的原始文件"时的安全网，跟 `load_as_working_copy()`
    是两回事：那个是"处理流程要在复制品上动手"，这个是"万一你把原件删了，
    这里还有一份"——两者职责不同，谁都不依赖谁。

    **每次导入都单独存一份，不判断是不是同一份文件、不覆盖已有备份**
    （2026-09-14 用户拍板：不比大小、不猜是不是同一份，重名就一律加序号）。
    同名时在文件名后面加" (1)"" (2)"这样的序号，逐个递增，直到找到没被
    占用的文件名为止。这意味着反复导入同一份没改过的文件也会越攒越多份
    备份——这正是要配合下面 `prune_backups()` 的自动清理一起用的。

    复制完之后，把这个文件名记进 `_MANIFEST_NAME` 这份清单里——`prune_backups()`
    只会删清单里记录过的文件，不会动这个目录里本来就有的、不是我们建的
    文件（用户拿一个已经在用的文件夹当备份路径时，这条尤其要紧）。
    """
    source = pathlib.Path(source)
    backup_dir = pathlib.Path(backup_dir)
    backup_dir.mkdir(parents=True, exist_ok=True)

    target = backup_dir / source.name
    if target.exists():
        stem, suffix = source.stem, source.suffix
        n = 1
        while True:
            candidate = backup_dir / f"{stem} ({n}){suffix}"
            if not candidate.exists():
                target = candidate
                break
            n += 1

    # ⚠ 用 copy 不用 copy2：copy2 会把**源文件的修改时间**一起搬过来，而
    # `prune_backups()` 正是按修改时间判断备份新旧的。搬过来之后，一份很久
    # 以前改过的论文，今天刚做出来的备份会被当成"过旧备份"立刻删掉——勾着
    # "自动删除过早备份（1 周）"导入一份 30 天前改过的论文，唯一那份刚做的
    # 备份会被这一轮清理直接清掉（2026-09-14 实测复现）。这里要表达的是
    # "这份备份是什么时候做的"，所以保留复制那一刻的时间。
    shutil.copy(source, target)

    names = _load_backup_manifest(backup_dir)
    names.add(target.name)
    _save_backup_manifest(backup_dir, names)

    return target


# "自动删除过早备份"里数量输入框旁边那个单位下拉框的换算表：填的数字乘上
# 这里对应的秒数，就是"超过这么久没动过就删"的界限。"月"按 30 天算，不是
# 精确的日历月份——这是个人工具的定期清理功能，没必要为这个再引入日期库
# 算精确月份。
_PRUNE_UNIT_SECONDS = {"天": 86400, "周": 86400 * 7, "月": 86400 * 30}


def prune_backups(backup_dir: str | pathlib.Path, amount: int, unit: str = "条") -> int:
    """清理备份目录里过旧的备份，返回实际删除的份数。

    `unit` 决定 `amount` 怎么解释：
    - "条"：只留最近 amount 份（按修改时间排序），多的从最旧的开始删；
    - "天"/"周"/"月"：删掉修改时间比 amount 个该单位更早的备份，不看总数量。

    **只会动 `backup_import()` 自己记在清单（`_MANIFEST_NAME`）里的文件**，
    这个目录里本来就有的、或者用户自己后来手动放进来的其他文件，一律不碰，
    不管它们的修改时间是不是"看起来很旧"。这是专门防"用户把备份路径设成
    一个已经在用的文件夹，结果自动清理把他自己的东西删了"这种情况——只
    按文件名规律或者修改时间去猜"这是不是我们建的备份"是不安全的，必须
    有一份明确记录。

    两种模式都按文件的**修改时间**判断，不是创建时间——`backup_import()`
    每次导入都新建一份、不覆盖刷新旧文件，而且**故意用 `shutil.copy` 而不是
    `copy2`**（copy2 会连源文件的修改时间一起搬过来，"很久以前改过的论文"
    的备份一诞生就会被判成过旧），所以备份文件的修改时间就是"这份备份是
    什么时候做的"。**改 `backup_import()` 里那个复制方式会直接推翻这里的
    判断依据**，两处要一起看。只看 backup_dir 直属的文件，不递归子目录。
    `amount <= 0` 时什么都不删（调用方应该在启用这个功能前自己校验，这里
    只是兜底，不主动报错）。
    """
    backup_dir = pathlib.Path(backup_dir)
    if amount <= 0 or not backup_dir.exists():
        return 0

    managed = _load_backup_manifest(backup_dir)
    existing_names = {p.name for p in backup_dir.iterdir() if p.is_file()}
    managed &= existing_names   # 清掉指向已经不存在的文件的陈旧记录，顺手瘦身
    files = [backup_dir / name for name in managed]

    if unit == "条":
        files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        stale = files[amount:]
    else:
        seconds = _PRUNE_UNIT_SECONDS.get(unit)
        if seconds is None:
            return 0
        cutoff = time.time() - amount * seconds
        stale = [p for p in files if p.stat().st_mtime < cutoff]

    for p in stale:
        p.unlink()
        managed.discard(p.name)
    _save_backup_manifest(backup_dir, managed)
    return len(stale)
