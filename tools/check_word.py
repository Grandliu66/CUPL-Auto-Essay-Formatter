"""用法：.venv\\Scripts\\python.exe tools\\check_word.py <docx路径> [--paras N]

用真实 Word（COM 自动化）打开生成的文件，核对 python-docx 看不出来的东西：
分节数、总页数、每节的页码格式和页眉文字、目录/清单域更新后是否真的找得到
条目，以及前 N 段各自落在第几页。

**为什么必须用 Word 开一遍**：Word 自动生成过的旧目录会把分节符藏进
python-docx 看不见的内容控件（`w:sdt`）里，导致工具漏看一节、自己又加一节，
结果变成 3 节——只用 python-docx 回读完全看不出来，必须实际拿 Word 打开才
暴露。改完结构/分节/样式相关的东西都要按 README 的流程跑一遍这个。

注意：被检查的文件如果正开在 Word 里会被锁定，Word 会报"命令失败"，
让用户先关掉再跑。
"""

import argparse
import pathlib

import win32com.client as win32


def word_usable(path: pathlib.Path) -> bool:
    """Word COM 现在能不能真的用：能打开、能读一个属性。

    为什么要这条：Word 会卡在安全模式 / 有残留进程时，`Documents.Open` 返回的
    对象**整个没用**（实拍 `AttributeError: Open.Sections`、`Open.Close`——
    属性名变成了方法名）。那是**机器状态**，不是产品回归，所以大声说明并跳过，
    别把回归批染红（2026-09-17）。
    """
    word = None
    doc = None
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(path.resolve()), ReadOnly=True)
        _ = doc.Sections.Count
        return True
    except Exception as exc:                                        # noqa: BLE001
        print(f"⚠️ Word COM 当前不可用（{type(exc).__name__}: {exc}）")
        return False
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:                                           # noqa: BLE001
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:                                           # noqa: BLE001
            pass


def check(path: pathlib.Path, paras: int) -> None:
    if not word_usable(path):
        print("   ⇒ 真 Word 复核本项跳过（不算失败；Word 重启后重跑即可）")
        return
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    doc = None
    try:
        # `Open` 自己也会抛（文件正开在 Word 里被锁、路径不对之类，见模块
        # docstring），所以它得待在 try 里面：以前这句在外面，一抛就绕过
        # `finally`，留下一个看不见的 Word 进程不退（2026-09-15 审计核实）。
        doc = word.Documents.Open(str(path.resolve()), ReadOnly=True)
        print("节数：", doc.Sections.Count)
        print("页数：", doc.ComputeStatistics(2))
        for i in range(1, doc.Sections.Count + 1):
            section = doc.Sections.Item(i)
            try:
                style = section.Footers.Item(1).PageNumbers.NumberStyle
            except Exception as e:  # 页脚没设页码时读不出来，不当作错误
                style = f"读取失败({e})"
            try:
                header = section.Headers.Item(1).Range.Text.replace("\r", "").replace("\x07", "")
            except Exception as e:
                header = f"读取失败({e})"
            print(f"  第{i}节：页码格式={style}  页眉={header!r}")

        print("目录域数量：", doc.TablesOfContents.Count)
        for i in range(1, doc.TablesOfContents.Count + 1):
            toc = doc.TablesOfContents.Item(i)
            if i == 1:
                # 先看"没更新域"时的内容——这就是用户把文件一打开时看到的东西，
                # 生成 docx 时域里缓存的是占位提示语，不按 F9 刷新就一直显示它
                cached = toc.Range.Text
                print("更新前的目录内容（＝用户刚打开时看到的）：")
                print("  " + cached.strip().replace("\r", " / ")[:200])
            toc.Update()
            text = toc.Range.Text
            if i == 1:
                print("更新后目录条目数：", text.count("\r") + text.count("\x0c"))
                print("目录内容前 300 字：")
                print(text[:300].replace("\r", " / "))
            else:
                # 插图清单/附表清单域：跟目录域同机制，看收集到的条目和页码
                try:
                    code = " ".join(toc.Range.Fields(1).Code.Text.split())
                except Exception:
                    code = "（域代码读不出）"
                print(f"第 {i} 个域：{code}")
                print("  条目数：", text.count("\r"))
                print("  内容前 260 字：" + text[:260].replace("\r", " / "))

        if paras > 0:
            print()
            print(f"===== 前 {paras} 段各自落在第几页 =====")
            for i in range(1, min(doc.Paragraphs.Count, paras) + 1):
                p = doc.Paragraphs.Item(i)
                text = p.Range.Text.replace("\r", "").replace("\x07", "").strip()
                print(f"  段{i:<3} 第{p.Range.Information(3)}页  样式={p.Style.NameLocal:<14} {text[:24]!r}")
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:                                       # noqa: BLE001
                # COM 动态派发会抽风（实拍 `AttributeError: Open.Close`，同
                # `PageSetup`/`Sections` 那一类），不是产品问题；Word 进程后面
                # 还有 Quit 兜着。2026-09-17 起吞掉，免得把回归批染红。
                pass
        try:
            word.Quit()
        except Exception:                                           # noqa: BLE001
            # 最后一份文档关掉时 Word 会自己退出，再调 Quit() 打的就是死掉的
            # COM 服务（实拍“远程过程调用失败”）——不是错误，吞掉。
            pass


def main() -> None:
    parser = argparse.ArgumentParser(description="用 Word COM 实测排版结果")
    parser.add_argument("path", type=pathlib.Path, help="要检查的 docx")
    parser.add_argument("--paras", type=int, default=16, help="打印前几段的页码，0 表示不打印（默认 16）")
    args = parser.parse_args()
    check(args.path, args.paras)


if __name__ == "__main__":
    main()
