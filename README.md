# PDF 处理工具

一个整合多个 PDF 常用操作的桌面/命令行工具，基于 Python 与 `pypdf` 实现，
支持 **Windows** 与 macOS。

## 功能一览

| 功能 | 说明 |
| --- | --- |
| 拆分 | 拆分 PDF（每页、页码范围、指定页） |
| 合并 | 合并多个 PDF，可调整顺序 |
| 删除页 | 删除 PDF 的指定页码 |
| doc转PDF | doc/docx 批量转 PDF 并按清单顺序拼接 |
| 转A4 | 批量将 PDF 页面等比缩放、居中到标准 A4 |

- Windows 图形界面：`gui_app.py`（tkinter），打包为单文件 exe
- 命令行：`main.py <功能> [参数]`，跨平台可用

> 拆分页交互：页码范围框支持混写（如 `1-5`、`1,3,5`、`2,4-6,8`，留空 = 整本）。
> 勾选“是否合并成一个文件”时输出 1 个 PDF；不勾选则把所选页拆成多个单页 PDF。

## Windows 桌面版（推荐）

双击运行的单文件 GUI 程序“**PDF 处理工具.exe**”，无需安装 Python，
五个功能都通过对话框选文件/目录完成，界面底部有运行日志。

### 获取 exe（GitHub Actions 自动构建）

代码推送后，GitHub 的 Windows runner 会自动打包并上传产物：

1. 打开仓库的 **Actions** 页面 → 选择 **Build Windows exe** 工作流
2. 在最近一次运行记录（绿色 ✓）底部 **Artifacts** 下载
   `PDF处理工具-windows`
3. 解压得到 `PDF处理工具.exe`，双击即可使用

也可以在 Actions 页面手动触发（Run workflow → 分支选 main）。

> 说明：doc转PDF 功能在 Windows 上通过 Word 的 COM 接口转换，
> 运行该功能需要目标电脑安装 **Microsoft Word**。
> 其余四个功能（拆分/合并/删除页/转A4）不依赖任何外部软件。

> **Windows 7 兼容版**：默认 exe 用 Python 3.11 构建，最低要求 Windows 8.1。
> 需要在 **Win7** 运行时，请下载 **Build Windows exe (Win7)** 工作流的产物
> `PDF处理工具-windows-win7`（用最后一个支持 Win7 的 Python 3.8 构建）。
> 若 Win7 仍提示缺少 `vcruntime140.dll` / `api-ms-win-crt-*.dll`，
> 请安装“Microsoft Visual C++ 2015-2022 运行库(x64)”及系统更新 **KB2999226**。

### Windows 本地手动打包

在装有 Python 3.9+ 的 Windows 电脑上，进入项目目录双击
`build_windows.bat`，产物生成在 `dist\PDF处理工具.exe`。

等价命令：

```bat
pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --clean pdf_tool.spec
```

## macOS 桌面版（.dmg）

绿色图标的 `PDF处理工具.app`，拖入“应用程序”即可使用，界面与 Windows 版一致。

### 获取 dmg（GitHub Actions 自动构建）

1. 打开仓库的 **Actions** 页面 → 选择 **Build macOS dmg** 工作流
2. 在最近一次运行记录（绿色 ✓）底部 **Artifacts** 下载
   `PDF处理工具-macos`
3. 解压得到 `PDF处理工具.dmg`，双击挂载后把 `PDF处理工具` 拖进 Applications

> 首次打开提示“无法验证开发者”时：右键点 App → 打开；
> 或到 系统设置 → 隐私与安全性 里点击“仍要打开”。
> （如需彻底去除提示，可用 Apple Developer ID 对产物签名/公证。）

### macOS 本地手动打包

```bash
bash build_dmg.sh      # 需要 Python3 + 网络；产物在 dist_mac/PDF处理工具.dmg
```

## 目录结构

```
pdf_processor/
├── gui_app.py               # 图形界面（Windows/macOS 通用）
├── main.py                  # 命令行整合入口
├── pdf_split.py             # 拆分模块
├── merge_pdfs.py            # 合并模块（含顺序调整 GUI 逻辑）
├── pdf_delete_pages.py      # 删除页码模块
├── doc2pdf_merge.py         # doc/docx 转 PDF 并拼接（Windows: Word COM）
├── pdf_to_a4.py             # PDF 批量标准化为 A4
├── requirements.txt         # 运行依赖（pypdf；Windows 额外 pywin32）
├── pdf_tool.spec            # Windows PyInstaller 打包配置（8.1+）
├── pdf_tool_win7.spec       # Windows 7 兼容打包配置（Python 3.8）
├── requirements-win7.txt    # Win7 构建的依赖锁定
├── pdf_tool_mac.spec        # macOS PyInstaller 打包配置（.app）
├── build_windows.bat        # Windows 手动打包脚本
├── build_dmg.sh             # macOS 手动打包脚本（产出 .dmg）
├── make_icon.py             # 纯 Python 生成 app.ico / app.icns
├── assets/                  # 应用图标
└── .github/workflows/
    ├── build-windows.yml    # 自动构建 Windows exe
    └── build-macos.yml      # 自动构建 macOS dmg
```

## 环境要求

- Python 3.9+，运行依赖见 `requirements.txt`：

```bash
pip install -r requirements.txt
```

- doc转PDF 需要本机安装 Microsoft Word（Windows 走 COM；
  macOS 走 AppleScript，两种平台都无需再装 Ghostscript）

## 命令行用法（跨平台）

所有功能通过 `main.py` 调用：

```bash
python main.py <功能> [参数]
```

运行 `python main.py -h` 查看命令列表。

### 1. 拆分（split）

```bash
python main.py split <PDF文件> [-r 页码范围] [-p 指定页] [-o 输出目录]
```

不指定 `-r` / `-p` 时按每页拆分。

| 参数 | 说明 |
| --- | --- |
| `-r, --range` | 按页码范围拆分，如 `1-5` |
| `-p, --pages` | 按指定页拆分，如 `1,3,5` |
| `-o, --output` | 输出目录，默认与输入 PDF 同目录 |

输出命名：每页 `<原名>_第N页.pdf`；页码范围 `<原名>_第起始-结束页.pdf`；
指定页 `<原名>_自定义.pdf`。

### 2. 合并（merge）

```bash
python main.py merge <PDF1> <PDF2> ... [-o 输出.pdf]
```

不跟文件时打开图形窗口选择文件并调整顺序后合并。默认输出 `merged.pdf`。

### 3. 删除页（delete）

```bash
python main.py delete <PDF文件> -d <页码> [-r <页码范围>] [-o 输出.pdf]
```

`-d` 与 `-r` 可混写（如 `-d 1,3`、`-r 2-4`、`-d 3 -r 5-7`）。
不指定 `-o` 时输出为 `<原名>_删除后.pdf`。

### 4. doc/docx 转 PDF 并合并（doc2pdf）

```bash
# 合并为一个 PDF（默认）
python main.py doc2pdf <清单文件> [输出.pdf]

# 不合并：每份文档单独转成一个 PDF，放入输出目录
python main.py doc2pdf <清单文件> <输出目录> --no-merge
```

清单文件（如 `order.txt`）每行一个文件名，按行顺序处理；空行与
`#` 开头的注释行忽略；相对路径相对于清单所在目录；`.pdf` 文件直接使用，
其余按 Word 文档处理（需要本机安装 Word）。

> GUI 版“doc转PDF”页也有“是否合并成一个文件”复选框：
> 勾选输出单个 PDF，不勾选则每份文档单独输出到“转换结果”目录
> （重名文件自动追加 `_2/_3…` 避免覆盖）。

### 5. 批量标准化为 A4（a4）

```bash
python main.py a4 [PDF目录] [-o 输出目录] [--overwrite]
```

页面内容保持原比例并居中到 A4 画布（595.28 × 841.89 pt），不会拉伸变形。
目录缺省时使用当前目录；输出缺省为输入目录下的 `A4标准版` 子目录，
同名输出已存在时跳过，`--overwrite` 可覆盖。

## 注意事项

- 页码从 1 开始计数。
- 超出范围的页码会被忽略（split 自动截断，delete 直接跳过）。
- 仅重命名/上传场景建议先转 A4，压缩文件体积问题请用专业工具处理。
