# 银行流水自动生成手工日记账工具

面向财务同事的 Windows GUI 工具（PySide6），用于将银行流水自动映射并追加写入手工日记账模板，最终可打包为单个 `.exe` 直接使用。

## 目录结构

```text
codex-test-repo/
├─ main.py
├─ requirements.txt
├─ ui/
│  └─ main_window.py
├─ services/
│  ├─ excel_service.py
│  └─ mapping_service.py
├─ utils/
│  ├─ helpers.py
│  └─ logger.py
└─ tests/
   └─ smoke_test.py
```

## 安装依赖

```bash
pip install -r requirements.txt
```

## 启动方式

```bash
python main.py
```

## 打包为 exe（推荐）

```bash
pyinstaller --noconfirm --clean --noconsole --onefile --name 银行流水自动填充工具 main.py
```

如需图标：

```bash
pyinstaller --noconfirm --clean --noconsole --onefile --name 银行流水自动填充工具 --icon assets/app.ico main.py
```

## 功能说明

- 选择银行流水文件（`.xls` / `.xlsx`）
- 选择模板文件（`.xlsx`）
- 自动识别银行流水正式表头（按关键字段命中数最高行）
- 自动识别模板目标页签（兼容轻微差异）
- 数据清洗（日期、金额、文本）
- 字段映射与借贷对调
- 按模板已有最大序号继续递增
- 会计期间按历史格式（`YYYY-MM` / `YYYYMM` / `YYYY年M月`）生成
- 追加写入（不覆盖原模板）并尽量继承上一行样式
- GUI 日志 + 成功/失败弹窗 + 一键打开输出目录

## 打包与部署注意事项

1. **目标平台**：请在 Windows 环境执行 PyInstaller 打包，保证 Qt 运行库一致。
2. **常见问题**：
   - 杀毒软件误报：建议加入企业白名单；首次运行可改用目录模式 `--onedir` 排查。
   - Qt 插件缺失：升级到匹配版本的 `PySide6` + `pyinstaller`，必要时执行 `--collect-all PySide6`。
   - 输出失败（文件占用）：关闭已打开的模板/输出 Excel 后重试。
   - 权限不足：避免输出到受保护目录（如 `C:\Windows`），推荐桌面或文档目录。
3. **同事使用建议**：将 `dist/银行流水自动填充工具.exe` 发给同事，双击即可运行，无需安装 Python。

## 已做兼容处理

- 银行流水列名兼容（空格、大小写、中英文括号、别名）
- 银行流水存在说明区/余额区时自动定位真实表头
- 过滤空白行、合计行，仅处理交易日期有效的记录
- 金额支持千分位、空格、括号负数
- 文本清洗避免将 `None/nan` 写入 Excel
- 模板页签名轻微差异兼容
- 模板列名轻微差异兼容
- 模板中异常序号值（文本/公式/空值）会自动跳过并正确找最大值

## 测试

```bash
python tests/smoke_test.py
```

测试覆盖：
- 表头识别
- 页签识别
- 映射规则（备注兜底摘要）
- 借贷金额对调
- 序号递增
- 会计期间生成
- 输出文件生成


## 在 VSCode 中测试（Windows）

1. 用 VSCode 打开项目根目录 `codex-test-repo`。
2. 打开终端（`Ctrl+``），创建并激活虚拟环境：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

3. 安装依赖：

```powershell
pip install -r requirements.txt
```

4. 在 VSCode 右下角选择解释器为 `.venv`（或按 `Ctrl+Shift+P` 执行 `Python: Select Interpreter`）。
5. 先跑逻辑烟雾测试：

```powershell
python tests/smoke_test.py
```

6. 再启动 GUI 手工验证：

```powershell
python main.py
```

7. 建议在 VSCode 中重点验证：
   - 选择 `.xls/.xlsx` 银行流水是否能识别表头；
   - 模板页签是否能自动识别；
   - 输出文件是否生成到目标目录；
   - 金额借贷是否对调写入；
   - 序号是否在历史最大值后递增。

8. 若需调试，直接按 `F5`，选择 `Python: Current File`（建议当前文件设为 `main.py`）。
