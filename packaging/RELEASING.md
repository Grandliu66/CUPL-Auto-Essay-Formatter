# 发版流程（维护者用）

「检查更新」读的是仓库根目录的 `latest.json`（版本号、日期、更新日志、下载链接）。软件里写死了两个读取地址（`essay_tool/update_check.py` 的 `SOURCES`），按顺序试，第一份合格的就用：

1. Gitee：`https://gitee.com/grandliu66/cupl-utaf-auto-update-check/raw/master/latest.json`（国内多数网络能访问；这是一个**独立的小仓库**，只放 `latest.json` 和说明，手动同步，不镜像本仓库）
2. GitHub：`https://raw.githubusercontent.com/Grandliu66/CUPL-Auto-Essay-Formatter/main/latest.json`（就是本仓库根目录的那份）

这两个地址随软件发出去后就改不了，要换只能发新版本；下载链接则可以随时改（改 `packaging/download_links.json` 再重新生成）。该功能只用标准库，不涉及 `requirements*.txt`。

## 每次发版

1. 写好 `CHANGELOG.txt` 最上面一条（顺序：新增功能、优化、问题修复）和软件内的 `CHANGELOG_TEXT`（`essay_tool/gui/main_window.py`），改 `essay_tool/__init__.py` 的版本号。
2. 跑 `Adversarial Automatic Test\run_aat.py` 和 `tools\check_update_check.py`，再 `powershell -ExecutionPolicy Bypass -File packaging\build.ps1` 打包（会自检）。
3. **把安装包传到网盘，确认分享链接可用**（链接变了就改 `packaging/download_links.json`）。
4. `python tools/make_latest_json.py` 生成 `latest.json`（版本号与日志对不上会报错）。
5. 提交并推送到 GitHub（含 `latest.json`），创建 GitHub Release 并上传安装包。
6. 把同一份 `latest.json` 推到 Gitee 小仓库（`master` 分支）。顺序要紧：**最后**才推 `latest.json`，否则用户会先收到“有新版本”，点进去却下载不到。
7. 在国内网络（校园网、手机流量）各打开一次 Gitee 的 raw 地址，确认能读到。
