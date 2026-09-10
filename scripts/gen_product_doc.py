# -*- coding: utf-8 -*-
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
    "版本：V0.1（需求提案）\n"
    "适用对象：无 Python 环境的业务同事（小白用户）\n"
    "交付形态：Windows 独立 .exe（双击运行）\n"
    "编写日期：2026-08-07\n"
    "状态：待审核"
)

h("1. 文档目的", 1)
p(
    "本文档描述「SMART 单票 Invoice 助手」面向最终用户的产品形态："
    "如何用 exe 启动、如何登录邮箱、如何选择文件夹、如何抄送多人发信。"
    "供业务/产品审核界面与流程是否符合预期；审核通过后再进入界面开发与打包。"
)

h("2. 用户画像与约束", 1)
table(
    ["项", "说明"],
    [
        ["用户", "对账/客服等业务同事，不熟悉命令行与 Python"],
        ["环境", "Windows 电脑；已安装 Microsoft Excel；无 Python"],
        ["交付", "解压文件夹后双击 .exe 即可，无需安装依赖"],
        ["邮箱", "支持 QQ 邮箱 / 国内常见企业邮（SMTP + 客户端授权码）"],
        ["核心任务", "选文件夹 → 生成单票 PDF → 按固定格式发邮件（可多抄送）"],
    ],
)

h("3. 产品目标（一句话）", 1)
p(
    "把「已确认账单 Excel」自动拆成按柜号命名的 Invoice PDF，"
    "并按固定邮件模板发给客户，支持选择多个抄送人；全程图形界面，零命令行。"
)

h("4. 总体使用流程（小白视角）", 1)
bullets(
    [
        "第一次：解压软件包 → 双击「SMART发票助手.exe」",
        "第 1 屏：选择邮箱类型 → 输入邮箱账号 → 输入授权码 → 进入系统",
        "第 2 屏（主界面）：选择「待处理 Excel 文件夹」与「结果保存文件夹」→ 点「开始生成」",
        "生成完成后：填写收件人、添加多个抄送邮箱 → 点「发送邮件」→ 确认后发出",
        "下次打开：可记住邮箱类型与账号；路径可记忆上次选择",
    ]
)
p("流程：启动 exe → 邮箱登录页 → 主功能页（路径 + 生成 + 发信）→ 完成提示", bold=True)

h("5. 界面设计（页面说明）", 1)

h("5.1 启动页 / 邮箱登录页", 2)
p("用户双击 exe 后首先看到本页。目的：完成发件邮箱配置，再进入业务功能。")
table(
    ["控件", "说明", "示例/默认"],
    [
        ["标题", "软件名称 + 简短说明", "SMART 单票 Invoice 助手"],
        [
            "邮箱类型（下拉）",
            "必选，降低小白填 SMTP 的难度",
            "QQ邮箱 / 腾讯企业邮 / 网易企业邮 / 阿里云企业邮 / 华为云 / 其他(自定义)",
        ],
        ["邮箱账号（输入框）", "完整邮箱地址", "1395970278@qq.com"],
        ["授权码（密码框）", "客户端授权码，非登录密码；可显示/隐藏", "••••••••"],
        ["记住账号", "勾选后下次自动填账号与邮箱类型", "默认勾选"],
        ["记住授权码（可选）", "涉及安全，需产品拍板；默认不勾选更安全", "默认不勾选"],
        ["【进入系统】按钮", "校验非空后进入主界面", ""],
        ["帮助小字", "如何获取 QQ/企业邮授权码（一句话说明）", ""],
    ],
)
p("校验规则：", bold=True)
bullets(
    [
        "邮箱类型、账号、授权码任一为空 → 提示「请填写完整」不进入",
        "账号需包含 @",
        "选择「其他(自定义)」时，额外显示：SMTP 服务器、端口、是否 SSL（可折叠）",
        "登录页不真实发信；进入后发信时再连接 SMTP",
    ]
)

h("5.2 主界面 — 路径与生成区（上半部分）", 2)
p("只出现「文件夹」和「大按钮」，不出现命令行。")
table(
    ["控件", "说明"],
    [
        ["当前发件邮箱（只读）", "显示已登录邮箱，旁有「切换账号」返回登录页"],
        ["待处理 Excel 文件夹", "输入框 +【浏览…】；放一个或多个「已确认」xlsx"],
        ["结果保存文件夹", "输入框 +【浏览…】；PDF 与对账报告写到这里"],
        ["【开始生成 Invoice】", "扫描 Excel 文件夹 → 按柜生成 PDF"],
        ["进度/日志区", "白话显示：处理进度、成功柜数、失败原因"],
        ["打开结果文件夹", "生成后一键打开资源管理器"],
    ],
)
p("结果目录建议：", bold=True)
bullets(
    [
        "结果保存路径/账单文件名/*.pdf（每个柜号一个 PDF）",
        "结果保存路径/账单文件名/report.xlsx（对账报告）",
        "多个 Excel 时按账单分子文件夹，避免同柜号覆盖",
    ]
)
p("生成前检查：", bold=True)
bullets(
    [
        "路径为空 → 提示先选择",
        "文件夹里没有 xlsx → 提示把已确认 Excel 放进该文件夹",
        "未安装 Excel → 明确提示（导出 PDF 依赖本机 Excel）",
        "对账失败 → 该账单不产出 PDF，红字说明原因",
    ]
)

h("5.3 主界面 — 发信区（下半部分）", 2)
p("对应步骤四：按固定格式发信，支持多个抄送人。")
table(
    ["控件", "说明"],
    [
        ["选择要发送的账单", "下拉列表：已生成账单（账单名、PDF 个数）"],
        ["收件人 To", "可改；测试时可填内部同事邮箱"],
        ["抄送人 CC", "支持多个：【添加抄送】动态行，或分号粘贴一批"],
        ["邮件预览", "只读：主题、正文、附件数量"],
        ["【发送邮件】", "默认先弹确认框"],
        ["以后自动发送", "勾选后下次跳过确认（本地记住）"],
    ],
)

h("5.4 邮件内容格式（固定）", 2)
table(
    ["项", "规则"],
    [
        ["发件人", "登录页填写的邮箱"],
        ["收件人 To", "主界面填写"],
        ["抄送 CC", "0～N 个"],
        ["主题", "{账单名}账单invoice，如 #DS-06052026-SMART账单invoice"],
        [
            "正文",
            "Hello, SMART team. + 账单名 + 账单invoice已附上，请查收，谢谢！ + Best Regards + 落款",
        ],
        ["附件", "该账单下全部 {柜号}.pdf"],
        ["附件过大", "超过约 20MB 自动拆多封，界面提示"],
    ],
)

h("6. 页面线框（文字版）", 1)
h("6.1 登录页", 2)
p(
    "┌──────────────────────────────────────┐\n"
    "│     SMART 单票 Invoice 助手          │\n"
    "│  首次使用请选择邮箱并填写授权码        │\n"
    "│                                      │\n"
    "│  邮箱类型  [ QQ邮箱 ▼ ]               │\n"
    "│  邮箱账号  [ ________________ ]       │\n"
    "│  授权码    [ **************** ]       │\n"
    "│  ☑ 记住账号   ☐ 记住授权码            │\n"
    "│           [ 进入系统 ]                │\n"
    "│  小字：授权码不是登录密码，获取方法…   │\n"
    "└──────────────────────────────────────┘"
)

h("6.2 主界面", 2)
p(
    "┌──────────────────────────────────────────────┐\n"
    "│ SMART助手   发件：xxx@qq.com [切换账号]       │\n"
    "├──────────────────────────────────────────────┤\n"
    "│ 【生成 Invoice】                              │\n"
    "│ 待处理Excel文件夹  [ D:\\账单\\input     ][浏览] │\n"
    "│ 结果保存文件夹    [ D:\\账单\\output    ][浏览] │\n"
    "│              [ 开始生成 Invoice ]             │\n"
    "│ 进度日志…                                     │\n"
    "├──────────────────────────────────────────────┤\n"
    "│ 【发送邮件】                                  │\n"
    "│ 选择账单  [ #DS-06052026-SMART ▼ ]            │\n"
    "│ 收件人To  [ client@xxx.com              ]     │\n"
    "│ 抄送CC    [ a@x.com ] [+添加抄送]             │\n"
    "│           [ b@y.com ] [删除]                  │\n"
    "│ 预览：主题/正文/附件数量                      │\n"
    "│ ☐ 以后自动发送     [ 发送邮件 ]               │\n"
    "└──────────────────────────────────────────────┘"
)

h("7. 小白友好设计原则", 1)
bullets(
    [
        "不出现 Python、命令行、pip（自定义邮箱除外才显示服务器）",
        "错误用白话：「文件夹里没有 Excel」「授权码不对」",
        "发送默认二次确认；自动发送需主动勾选",
        "授权码默认不永久保存",
        "软件包附《一页纸使用说明》",
    ]
)

h("8. 软件包交付形态（无 Python）", 1)
table(
    ["内容", "说明"],
    [
        ["SMART发票助手.exe + 依赖目录", "PyInstaller 目录版（比单文件更稳）"],
        ["templates/", "发票空白模板"],
        ["使用说明.pdf", "一页纸操作说明"],
        ["用户自选文件夹", "不必强制使用软件内 input/output"],
    ],
)
p("用户电脑前提：Windows 10/11；已装 Excel；发信需能上网。", bold=True)

h("9. 与现有能力对应", 1)
table(
    ["产品功能", "状态"],
    [
        ["按柜生成 PDF / 对账阻断", "已有"],
        ["SMTP 发信", "已有"],
        ["登录页选邮箱 + 授权码", "待按本文做界面"],
        ["自选输入/输出路径", "待做"],
        ["界面多抄送", "待做"],
        ["打成 exe", "待做"],
    ],
)

h("10. 安全与合规", 1)
bullets(
    [
        "授权码仅存本机（可选），不上传云端",
        "对账未通过禁止发信",
        "日志不打印完整授权码",
    ]
)

h("11. 待审核确认事项", 1)
table(
    ["编号", "问题", "建议默认"],
    [
        ["Q1", "是否记住登录直接进主界面？", "记住账号则直达，可切换账号"],
        ["Q2", "是否允许记住授权码？", "默认否"],
        ["Q3", "收件人 To 默认值？", "可配置，界面可改"],
        ["Q4", "落款姓名默认？", "设置项，可自填"],
        ["Q5", "多抄送录入方式？", "动态多行 + 分号粘贴"],
        ["Q6", "是否要「生成并发送」一键？", "先分两步，后续可加"],
        ["Q7", "单文件 exe 还是目录版？", "目录版更稳"],
    ],
)

h("12. 审核通过后实施顺序", 1)
bullets(
    [
        "改 GUI：登录页 + 路径选择 + 多抄送",
        "联调生成与发信",
        "PyInstaller 打包 + 一页纸说明书",
        "业务试用后定稿",
    ]
)

h("13. 总结", 1)
p(
    "以小白可双击使用为目标：先选邮箱并填授权码，再选 Excel 文件夹与结果文件夹生成 PDF，"
    "最后配置收件人与多个抄送人，按步骤四固定格式发信。"
    "请审核界面与默认规则；确认后按第 12 节实现与打包。"
)

out = Path(__file__).resolve().parents[2] / "SMART发票助手_产品界面与使用流程_供审核V0.1.docx"
doc.save(out)
print("saved", out)
