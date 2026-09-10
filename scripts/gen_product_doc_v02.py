# -*- coding: utf-8 -*-
"""生成不含步骤四发信的产品审核 Word 文档。"""
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from pathlib import Path

doc = Document()
style = doc.styles["Normal"]
style.font.name = "微软雅黑"
style.font.size = Pt(11)
style._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
for i in range(1, 4):
    hs = doc.styles[f"Heading {i}"]
    hs.font.name = "微软雅黑"
    hs.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    hs._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")


def h(t, lv=1):
    doc.add_heading(t, level=lv)


def p(t, bold=False):
    para = doc.add_paragraph()
    r = para.add_run(t)
    r.bold = bold


def bullets(items):
    for it in items:
        doc.add_paragraph(it, style="List Bullet")


def table(headers, rows):
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


title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = title.add_run("SMART 单票 Invoice 助手")
r.bold = True
r.font.size = Pt(22)
r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub.add_run("产品界面与使用流程设计说明（供审核）").font.size = Pt(14)

meta = doc.add_paragraph()
meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
meta.add_run(
    "版本：V0.2\n"
    "适用范围：仅「生成单票 Invoice PDF」\n"
    "明确不做：步骤四发送邮件\n"
    "适用对象：无 Python 环境的业务同事（小白用户）\n"
    "交付形态：Windows 独立 .exe（双击运行）\n"
    "编写日期：2026-08-07\n"
    "状态：待审核"
)

h("1. 文档目的", 1)
p(
    "本文档描述面向最终用户的产品形态：exe 启动后如何填写邮箱信息、如何选择 Excel 文件夹与结果保存路径、"
    "如何一键生成按柜号命名的 Invoice PDF。"
    "供业务审核；审核通过后再开发界面并打包。"
)
p("本期明确不做：步骤四「发送账单 Invoice 给客户」（含抄送、SMTP 实发等）。", bold=True)

h("2. 用户画像与约束", 1)
table(
    ["项", "说明"],
    [
        ["用户", "业务同事，不熟悉命令行与 Python，按界面点选即可"],
        ["环境", "Windows 10/11；已安装 Microsoft Excel；无 Python"],
        ["交付", "解压后双击 .exe，无需安装 Python 或依赖库"],
        ["本期目标", "选文件夹 → 自动生成单票 Invoice PDF"],
        ["本期不做", "自动发邮件 / Outlook / 谷歌邮箱对接"],
    ],
)

h("3. 产品目标（一句话）", 1)
p(
    "把「已确认账单 Excel」自动拆成按柜号命名的 Invoice PDF；"
    "全程图形界面，零命令行；客户电脑无需 Python。"
)

h("4. 总体使用流程（小白视角）", 1)
bullets(
    [
        "解压软件包 → 双击「SMART发票助手.exe」",
        "第 1 屏：选择邮箱类型 → 输入邮箱账号 → 输入授权码 → 点「进入」",
        "第 2 屏：选择「待处理 Excel 文件夹」和「结果保存文件夹」→ 点「开始生成」",
        "等待进度完成 → 点「打开结果文件夹」查看 PDF",
        "（本期无「发送邮件」按钮）",
    ]
)
p(
    "说明：第 1 屏保留邮箱/授权码录入，便于统一账号身份、以及以后若重启步骤四可复用；"
    "本期进入系统后不会发信，也不调用谷歌/Outlook。",
    bold=True,
)

h("5. 界面设计", 1)

h("5.1 启动页 / 账号页", 2)
p("双击 exe 后首先进入本页。")
table(
    ["控件", "说明", "示例"],
    [
        ["标题", "软件名称", "SMART 单票 Invoice 助手"],
        [
            "邮箱类型（下拉）",
            "降低小白理解成本",
            "QQ邮箱 / 腾讯企业邮 / 网易企业邮 / 阿里云企业邮 / 其他",
        ],
        ["邮箱账号", "完整邮箱地址", "xxx@qq.com"],
        ["授权码", "客户端授权码（不是登录密码），密码框可显示/隐藏", "••••••••"],
        ["记住账号", "下次自动带出类型与账号", "默认勾选"],
        ["记住授权码", "涉及安全，建议默认不勾选", "默认不勾选"],
        ["【进入】", "校验非空后进入主界面", ""],
        ["帮助小字", "授权码如何获取（一句话）", ""],
    ],
)
p("校验：", bold=True)
bullets(
    [
        "类型/账号/授权码为空 → 提示「请填写完整」",
        "账号需含 @",
        "选「其他」时展开 SMTP 服务器、端口（高级，可折叠）",
        "本期不在此页验证能否发信，只做格式校验后进入",
    ]
)

h("5.2 主界面（仅生成，无发信）", 2)
table(
    ["控件", "说明"],
    [
        ["当前账号（只读）", "显示已录入邮箱；可「切换账号」回第 1 屏"],
        ["待处理 Excel 文件夹", "输入框 +【浏览…】；内含一个或多个已确认 xlsx"],
        ["结果保存文件夹", "输入框 +【浏览…】；生成 PDF 与对账报告"],
        ["【开始生成 Invoice】", "主按钮：扫描文件夹并按柜生成 PDF"],
        ["进度/日志", "白话提示进度、成功柜数、失败原因"],
        ["【打开结果文件夹】", "生成后一键打开"],
    ],
)
p("主界面不出现：发送邮件、收件人、抄送人、邮件预览。", bold=True)

p("结果目录建议：", bold=True)
bullets(
    [
        "结果路径/账单文件名/*.pdf（每个柜号一个 PDF）",
        "结果路径/账单文件名/report.xlsx（对账报告，可抽查）",
        "多个 Excel 时按账单分子文件夹，避免同柜号覆盖",
    ]
)

p("生成前检查：", bold=True)
bullets(
    [
        "路径为空 → 提示先选择文件夹",
        "文件夹内无 xlsx → 提示放入已确认 Excel",
        "未安装 Excel → 提示需安装 Microsoft Excel（用于导出 PDF）",
        "对账失败 → 该账单不产出 PDF，日志红字说明",
    ]
)

h("6. 页面线框（文字版）", 1)
h("6.1 账号页", 2)
p(
    "┌──────────────────────────────────────┐\n"
    "│     SMART 单票 Invoice 助手          │\n"
    "│  请选择邮箱类型并填写账号、授权码      │\n"
    "│                                      │\n"
    "│  邮箱类型  [ QQ邮箱 ▼ ]               │\n"
    "│  邮箱账号  [ ________________ ]       │\n"
    "│  授权码    [ **************** ]       │\n"
    "│  ☑ 记住账号   ☐ 记住授权码            │\n"
    "│           [ 进入 ]                    │\n"
    "│  提示：本期仅用于生成PDF，不会发邮件   │\n"
    "└──────────────────────────────────────┘"
)

h("6.2 主界面", 2)
p(
    "┌──────────────────────────────────────────────┐\n"
    "│ SMART助手   账号：xxx@qq.com  [切换账号]      │\n"
    "├──────────────────────────────────────────────┤\n"
    "│ 待处理Excel文件夹  [ D:\\账单\\待处理   ][浏览]  │\n"
    "│ 结果保存文件夹    [ D:\\账单\\结果     ][浏览]  │\n"
    "│                                              │\n"
    "│            [ 开始生成 Invoice ]               │\n"
    "│                                              │\n"
    "│ 进度日志：正在处理… 已完成 17 个柜号           │\n"
    "│            [ 打开结果文件夹 ]                 │\n"
    "└──────────────────────────────────────────────┘"
)

h("7. 小白友好原则", 1)
bullets(
    [
        "不出现 Python、命令行、pip",
        "错误提示用白话中文",
        "大按钮、少选项；本期不做发信，避免干扰",
        "软件包附一页纸《使用说明》",
    ]
)

h("8. 软件包交付（无 Python）", 1)
table(
    ["内容", "说明"],
    [
        ["SMART发票助手.exe + 依赖目录", "PyInstaller 目录版（推荐，比单文件更稳）"],
        ["templates/", "发票空白模板"],
        ["使用说明.pdf", "如何选文件夹、如何看结果"],
    ],
)
p("用户电脑前提：", bold=True)
bullets(
    [
        "Windows 10/11",
        "已安装 Microsoft Excel（导出 PDF 需要）",
        "本期即使填了邮箱，也不要求能连外网发信；生成 PDF 可离线完成",
    ]
)

h("9. 与现有能力对应", 1)
table(
    ["功能", "状态"],
    [
        ["按柜解析 / 对账 / 填模板导出 PDF", "已有"],
        ["账号页 + 自选路径 GUI", "待按本文实现"],
        ["打成 exe", "待做"],
        ["步骤四发邮件", "本期不做"],
    ],
)

h("10. 待审核确认事项", 1)
table(
    ["编号", "问题", "建议默认"],
    [
        ["Q1", "既不发信，账号页是否仍要保留？", "按你的要求保留；页上注明「本期不发信」"],
        ["Q2", "是否记住账号直达主界面？", "记住账号则直达，可切换"],
        ["Q3", "是否允许记住授权码？", "默认否"],
        ["Q4", "单文件 exe 还是目录版？", "目录版更稳"],
        ["Q5", "默认结果是否按账单分子文件夹？", "是"],
    ],
)

h("11. 审核通过后实施顺序", 1)
bullets(
    [
        "实现账号页 + 主界面（仅生成）",
        "联调自选文件夹生成 PDF",
        "PyInstaller 打包 exe + 一页纸说明",
        "业务试用",
    ]
)

h("12. 总结", 1)
p(
    "本期产品 = exe 双击 → 填写邮箱类型/账号/授权码进入 → 选择 Excel 文件夹与结果文件夹 → 生成单票 Invoice PDF。"
    "步骤四发邮件明确不做。"
    "请审核本流程；确认后按此实现界面与打包。"
)

out = Path(__file__).resolve().parents[2] / "SMART发票助手_产品界面与使用流程_供审核V0.2.docx"
doc.save(out)
print("saved", out)
