# -*- coding: utf-8 -*-
"""生成 SMART 发票助手使用教程（Word + TXT）。"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

OUT_DIR = Path(__file__).resolve().parents[1]


def _setup(doc: Document) -> None:
    style = doc.styles["Normal"]
    style.font.name = "微软雅黑"
    style.font.size = Pt(11)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    for i in range(1, 4):
        hs = doc.styles[f"Heading {i}"]
        hs.font.name = "微软雅黑"
        hs.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
        hs._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")


def h(doc, text, level=1):
    doc.add_heading(text, level=level)


def p(doc, text, bold=False):
    para = doc.add_paragraph()
    run = para.add_run(text)
    run.bold = bold
    return para


def bullets(doc, items):
    for it in items:
        doc.add_paragraph(it, style="List Bullet")


def numbered(doc, items):
    for it in items:
        doc.add_paragraph(it, style="List Number")


def table(doc, headers, rows):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, x in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = x
        for para in cell.paragraphs:
            for r in para.runs:
                r.bold = True
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            t.rows[ri + 1].cells[ci].text = str(val)
    doc.add_paragraph()


TXT_LINES: list[str] = []


def t(line: str = "") -> None:
    TXT_LINES.append(line)


def build() -> tuple[Document, str]:
    doc = Document()
    _setup(doc)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("SMART 单票 Invoice 助手")
    r.bold = True
    r.font.size = Pt(22)
    r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run("使用教程（业务同事 + 技术人员）").font.size = Pt(14)

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        "版本：v0.3　新旧模板兼容完整复核版\n"
        "日期：2026-08-20\n"
        "本期范围：只生成按柜号命名的 Invoice PDF\n"
        "明确不做：自动发送邮件"
    )

    t("SMART 单票 Invoice 助手 — 使用教程")
    t("=" * 48)
    t("版本：v0.3　新旧模板兼容完整复核版　日期：2026-08-20")
    t("本期范围：只生成按柜号命名的 Invoice PDF；明确不做自动发送邮件。")
    t()

    h(doc, "1. 这个软件做什么")
    p(
        doc,
        "把已经确认的 SMART 账单 Excel，按集装箱号拆成一张张 Invoice PDF。"
        "每张票对应一个柜号，版式对齐客户常用的单票 Invoice。"
        "业务同事不需要安装 Python，双击安装包里的 exe 即可。",
    )
    t("1. 这个软件做什么")
    t("把已经确认的 SMART 账单 Excel，按集装箱号拆成一张张 Invoice PDF。")
    t("业务同事不需要安装 Python，双击安装包里的 exe 即可。")
    t()

    h(doc, "2. 总包里有什么")
    table(
        doc,
        ["文件夹/文件", "给谁", "做什么"],
        [
            ["请先看_使用教程.docx / .txt", "所有人", "本教程"],
            ["01_安装包\\SMART发票助手.exe", "业务同事", "双击即可使用"],
            ["01_安装包\\使用说明.txt", "业务同事", "一页纸速查"],
            ["02_源码", "技术人员", "开发、改规则、重新打包"],
            ["版本说明.txt", "所有人", "本包版本与检查结果"],
        ],
    )
    t("2. 总包里有什么")
    t("01_安装包\\SMART发票助手.exe  — 给业务同事，双击即可")
    t("01_安装包\\使用说明.txt     — 一页纸速查")
    t("02_源码                    — 给技术人员")
    t("请先看_使用教程.docx / .txt — 本教程")
    t()

    h(doc, "3. 业务同事：安装与第一次使用", 1)
    h(doc, "3.1 电脑需要满足什么", 2)
    bullets(
        doc,
        [
            "Windows 10 或 Windows 11（64 位）",
            "已安装 Microsoft Excel，或已安装 WPS 表格（二选一即可）",
            "不需要安装 Python，不需要联网才能生成 PDF",
        ],
    )
    t("3. 业务同事：安装与第一次使用")
    t("3.1 电脑需要：Windows 10/11；Excel 或 WPS 表格二选一；不用装 Python。")
    t()

    h(doc, "3.2 怎么打开", 2)
    numbered(
        doc,
        [
            "把总包解压到一个固定目录，例如 D:\\SMART发票助手。不要解压到临时目录后直接从压缩软件窗口里双击。",
            "进入 01_安装包 文件夹。",
            "双击 SMART发票助手.exe。第一次打开会自解压，可能要等 5–10 秒，属正常。",
            "若杀毒软件拦截，请选择「允许运行」。本程序会调用本机 Excel/WPS 导出 PDF，部分杀毒软件可能误报。",
        ],
    )
    t("3.2 怎么打开：解压到固定目录 → 进入 01_安装包 → 双击 exe。第一次可能要等几秒。")
    t()

    h(doc, "3.3 登录页（不会发邮件）", 2)
    p(
        doc,
        "打开后先进入登录页。请选择 QQ邮箱 或 腾讯企业邮，填写邮箱账号和客户端授权码，再点「进入系统」。",
    )
    bullets(
        doc,
        [
            "本期登录只是进入系统用，程序不会发送任何邮件。",
            "授权码不是邮箱登录密码。QQ：设置 → 账户 → POP3/SMTP 服务 → 生成授权码。腾讯企业邮：邮箱设置里开启 SMTP 并生成授权码。",
            "账号必须包含 @，三项都填才能进入。",
            "程序不会在本地记住授权码。下次打开需要重新填写。",
        ],
    )
    t("3.3 登录页：选 QQ / 腾讯企业邮，填账号和授权码。本期不会发邮件，也不会记住授权码。")
    t()

    h(doc, "3.4 主界面四步", 2)
    numbered(
        doc,
        [
            "待处理 Excel：点「浏览」，选一个文件夹。请把已确认账单 .xlsx 放进这个文件夹。不要选单个文件，要选文件夹。",
            "结果保存位置：点「浏览」，另选一个空一点的文件夹。生成的 PDF 会写到这里。",
            "导出引擎：有 Excel 选 Microsoft Excel 或「自动」；只有 WPS 就选 WPS 表格。不要在没装 Excel 时硬选 Excel，否则可能弹出微软商店。",
            "点「开始生成 Invoice」。进度条和「03 进度」会滚动显示正在处理哪个文件。完成后点「打开结果」。",
        ],
    )
    t("3.4 主界面：选 Excel 文件夹 → 选结果文件夹 → 选 Excel 或 WPS → 开始生成。")
    t()

    h(doc, "3.5 生成过程中可以做什么", 2)
    bullets(
        doc,
        [
            "「中断任务」：弹出确认框，确认后停止后续文件。已经生成好的 PDF 会保留。",
            "「切换账号」：若正在生成，也会先确认是否中断。",
            "窗口缩小时，整页可以滚动；「03 进度」可以收起/展开。",
        ],
    )
    t("3.5 可中断任务（已生成的 PDF 保留）。窗口缩小后可滚动查看进度。")
    t()

    h(doc, "4. 结果文件夹里有什么")
    table(
        doc,
        ["内容", "说明"],
        [
            ["柜号.pdf", "例如 TIIU6724883.pdf，这就是要给客户的单票 Invoice"],
            ["多个 Excel 时的子文件夹", "按账单文件名分子目录，避免不同账单同柜号互相覆盖"],
            ["_logs\\report.xlsx", "校验报告：每个柜的金额、状态、失败原因"],
            ["_work\\filled_xlsx", "填好的中间 Excel（给导出 PDF 用，一般不用打开）"],
        ],
    )
    p(doc, "同一个柜号再次生成时，会先删除旧 PDF 再写新文件。")
    t("4. 结果：柜号.pdf；多 Excel 会分子文件夹；_logs\\report.xlsx 是校验报告。")
    t("同柜号再次生成会覆盖旧 PDF。")
    t()

    h(doc, "5. Excel 应该长什么样")
    p(doc, "程序支持两种已确认账单，会自动识别，不必手工切换。")
    h(doc, "5.1 新模板（推荐）", 2)
    p(doc, "表头需同时包含：APPOINTMENT ID、Total Volume、CBM、DESTINATION、ISA、QTY、RATE、AMOUNT。")
    bullets(
        doc,
        [
            "柜号格式：4 个字母 + 7 个数字，例如 TIIU6724883。合并单元格续行会归属当前柜。",
            "账单名优先取表头 #DS-日期-SMART；日期取自账单名，例如 #DS-06052026-SMART → 06/05/2026。",
            "「最终结算金额」如果是隐藏列里的文字反馈，不会阻断出账；可见且写了非数字时会整批停止。",
            "底部 Total Amount / 总计 / 合计 之后的行不会并入最后一个柜。",
        ],
    )
    h(doc, "5.2 旧模板", 2)
    p(
        doc,
        "部分早期账单把柜号和箱型、体积写在同一格（例如 ABCD1234567-45HC  80.5），"
        "并把 CBM 写在目的地后面（例如 FTW1 - 6.98）。程序会拆出真正的 11 位柜号、体积和 CBM。"
        "识别条件：有「合计金额」「invoice编号」列，且没有独立的 Total Volume / CBM 列。",
    )
    t("5. Excel：新模板要有 APPOINTMENT ID / Total Volume / CBM / DESTINATION / ISA / QTY / RATE / AMOUNT。")
    t("旧模板（柜号和体积写在同一格）也会自动识别。")
    t()

    h(doc, "6. 金额怎么算")
    bullets(
        doc,
        [
            "每个柜：把该柜所有明细行的原始金额加总，最后四舍五入到分。不会用「先每行四舍五入再相加」。",
            "源表金额是唯一出账金额。车架费/板费等规则只做审计，发现不一致会写进报告，但不会改客户账单数字。",
            "隐藏的「汇总」表默认不参与校验，避免手工汇总值拦住整批出账。",
        ],
    )
    t("6. 金额：柜内明细相加后四舍五入到分。费用规则只审计，不改源表金额。")
    t()

    h(doc, "7. 常见问题")
    table(
        doc,
        ["现象", "怎么办"],
        [
            ["双击没窗口", "先等 10 秒。仍没有，看同目录「启动失败日志.txt」，发给技术人员。"],
            ["提示没有 Excel", "把已确认 .xlsx 放进所选文件夹，不要放 .xls 或临时文件 ~$ 开头。"],
            ["弹出微软商店要装 Excel", "改选「WPS 表格」或「自动」。程序不会主动强制安装 Excel。"],
            ["批次失败、一个 PDF 都没有", "打开结果目录 _logs\\report.xlsx，看批次说明。常见原因：表头缺列、可见的最终结算金额不是数字。"],
            ["最后一个柜金额偏大", "检查 Excel 底部是否有总计/备注被当成明细。程序会在 Total Amount/总计处停表。"],
            ["窗口缩小看不到进度", "滚动主界面，或展开「03 进度」。"],
            ["杀毒软件删除 exe", "加入信任区后重新解压。不要在压缩包内部直接运行。"],
            ["要不要装 Python", "业务同事不要装。技术人员如需改源码，见第 8 节。"],
        ],
    )
    t("7. 常见问题：没窗口看启动失败日志；没 Excel 文件检查是否放对文件夹；批次失败看 _logs\\report.xlsx。")
    t()

    h(doc, "8. 技术人员：用源码运行和重新打包")
    h(doc, "8.1 环境", 2)
    p(doc, "建议 Windows + Python 3.11（64 位）。不要用 3.13 打给同事的 exe，兼容性不如 3.11。")
    p(doc, "在 02_源码 目录执行：", bold=False)
    p(doc, "python -m pip install -r requirements.txt")
    p(doc, "python run_app.py")
    p(doc, "命令行：python -m src --input ./input --output ./output --logs ./logs")
    h(doc, "8.2 重新打 exe", 2)
    p(doc, "python -m pip install pyinstaller")
    p(doc, "python -m PyInstaller --noconfirm --clean SMART_Invoice_App.spec")
    p(doc, "完成后把 dist\\SMART发票助手.exe 和 使用说明.txt 一起发给同事。")
    h(doc, "8.3 不要打进公开包的文件", 2)
    bullets(
        doc,
        [
            "config\\config.local.yaml（可能含邮箱授权码）",
            "config\\user_prefs.yaml",
            "_build_deps*、build*、dist、tmp、work、logs、output、__pycache__",
            "客户真实账单（input 里的业务文件）；样例请用 samples",
        ],
    )
    h(doc, "8.4 测试", 2)
    p(doc, "python -m unittest discover -s tests -p \"test_*.py\"")
    t("8. 技术人员：Python 3.11 + pip install -r requirements.txt + python run_app.py")
    t("打包：python -m PyInstaller --noconfirm --clean SMART_Invoice_App.spec")
    t("不要把 config.local.yaml（授权码）打进公开包。")
    t()

    h(doc, "9. 本期明确不做")
    bullets(
        doc,
        [
            "不会自动发邮件、不会抄送客户。",
            "不会改客户源 Excel。",
            "不会在 exe 旁边自动创建 input/output 文件夹，路径必须自己用「浏览」选择。",
        ],
    )
    t("9. 本期不做：发邮件、改源 Excel、自动建文件夹。")
    t()

    h(doc, "10. 本包检查记录（2026-08-20）")
    bullets(
        doc,
        [
            "单元测试（新旧模板布局、隐藏结算列）通过。",
            "样例账单 #DS-06052026-SMART 解析 17 柜，合计 67788.41，TIIU6724883 = 4862.71。",
            "input 目录现有 9 份已确认账单均可解析（含旧模板 BasicInvoice）。",
            "用本机 Microsoft Excel 实导出 #DS-07142026-SMART(SAV) 共 8 个 PDF，成功。",
            "安装包为单文件 exe，内置配置和发票模板；源码包已去掉授权码、构建缓存和客户账单。",
        ],
    )
    t("10. 检查记录：单元测试通过；样例 17 柜合计 67788.41；9 份 input 均可解析；Excel 实导出 8 个 PDF 成功。")
    t("安装包为单文件 exe；源码包不含授权码、构建缓存、客户账单。")

    return doc, "\n".join(TXT_LINES) + "\n"


def main() -> None:
    import sys

    out_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else OUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    doc, txt = build()
    docx_path = out_dir / "请先看_使用教程.docx"
    txt_path = out_dir / "请先看_使用教程.txt"
    doc.save(docx_path)
    txt_path.write_text(txt, encoding="utf-8")
    print(docx_path)
    print(txt_path)


if __name__ == "__main__":
    main()
