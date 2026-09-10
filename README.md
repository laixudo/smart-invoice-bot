# SMART 单票 Invoice 助手

已确认 SMART 账单 Excel → 按柜号生成单票 Invoice PDF。
**本期不含发邮件。** 业务同事双击 exe 即可使用，无需安装 Python。

当前版本：**v0.3**（新旧模板兼容完整复核版，2026-08-20）

## 业务同事怎么用

解压安装包，双击 `SMART发票助手.exe`。详见 `使用说明.txt` 和《使用教程》。

需要：Windows 10/11，以及 **Microsoft Excel 或 WPS 表格**（二选一）。

## 开发机运行

建议 Python 3.11：

```bat
cd smart_invoice_bot
python -m pip install -r requirements.txt
python run_app.py
```

命令行批量生成：

```bat
python -m src --input ./input --output ./output --logs ./logs
```

## 打包 exe

```bat
cd smart_invoice_bot
python -m PyInstaller --noconfirm --clean SMART_Invoice_App.spec
```

生成文件：`dist\SMART发票助手.exe`

## 目录说明

| 路径 | 用途 |
|------|------|
| `src/` | 程序源码 |
| `config/config.yaml` | 默认配置（模板字段、容差、公司抬头） |
| `templates/invoice_blank.xlsx` | 发票 Excel 模板 |
| `samples/` | 本地回归样例（含业务数据，不提交到仓库） |
| `tests/` | 单元测试 |
| `scripts/` | 打包/文档脚本 |
| `使用说明.txt` | 给业务同事的一页纸说明 |

不要把 `input/`、`samples/`、`config/config.local.yaml`、`config/user_prefs.yaml`
提交到仓库或打进发布包；这些目录可能含客户资料或授权码。
