# PDF 处理工具

一个整合多个 PDF 常用操作的命令行工具，基于 Python 与 `pypdf` 实现。

## 功能一览

| 命令 | 说明 |
| --- | --- |
| `split` | 拆分 PDF（每页、页码范围、指定页） |
| `merge` | 合并多个 PDF |
| `delete` | 删除 PDF 的指定页码 |
| `doc2pdf` | doc/docx 批量转 PDF 并按清单顺序拼接 |

## 目录结构

```
pdf_processor/
├── main.py                # 整合入口（统一调用入口）
├── pdf_split.py           # 拆分模块
├── merge_pdfs.py          # 合并模块
├── pdf_delete_pages.py    # 删除页码模块
└── doc2pdf_merge.py       # doc/docx 转 PDF 并合并模块
```

## 环境要求

- Python 3，依赖包 `pypdf`
- 建议在 conda 虚拟环境 `web` 中使用：

```bash
conda activate web
```

- `doc2pdf` 命令额外需要：
  - macOS 系统（通过 AppleScript 调用 Microsoft Word）
  - 已安装 Microsoft Word
  - 已安装 Ghostscript（`gs` 命令，用于最终拼接）

## 通用用法

所有功能都通过 `main.py` 调用，格式为：

```bash
python main.py <命令> [参数]
```

运行 `python main.py` 或 `python main.py -h` 可查看命令列表。

---

## 1. 拆分 PDF（split）

```bash
python main.py split <PDF文件> [-r 页码范围] [-p 指定页] [-o 输出目录]
```

不指定 `-r` / `-p` 时，按每页拆分。

| 参数 | 说明 |
| --- | --- |
| `-r, --range` | 按页码范围拆分，如 `1-5` |
| `-p, --pages` | 按指定页拆分，如 `1,3,5` |
| `-o, --output` | 输出目录，默认与输入 PDF 同目录 |

### 示例

```bash
# 按每页拆分
python main.py split 报告.pdf

# 拆分第 1 到 5 页
python main.py split 报告.pdf -r 1-5

# 只提取第 2、4、6 页
python main.py split 报告.pdf -p 2,4,6 -o /tmp/out
```

输出文件命名规则：

- 每页：`<原名>_第N页.pdf`
- 页码范围：`<原名>_第起始-结束页.pdf`
- 指定页：`<原名>_自定义.pdf`

---

## 2. 合并 PDF（merge）

```bash
python main.py merge <PDF1> <PDF2> ... [-o 输出.pdf]
```

按参数给出的顺序依次合并，输出默认名为 `merged.pdf`。

### 示例

```bash
python main.py merge 封面.pdf 正文.pdf 附录.pdf -o 完整文档.pdf
```

---

## 3. 删除指定页码（delete）

```bash
python main.py delete <PDF文件> -d <页码> [-r <页码范围>] [-o 输出.pdf]
```

`-d` 与 `-r` 可同时使用，也可混写（如 `-d 1,3`、`-r 2-4`、`-d 3 -r 5-7`）。
不指定 `-o` 时，输出为 `<原名>_删除后.pdf`。

### 示例

```bash
# 删除第 3、5 页
python main.py delete 报告.pdf -d 3,5

# 删除第 2 到 4 页
python main.py delete 报告.pdf -r 2-4

# 组合删除并指定输出
python main.py delete 报告.pdf -d 1,10 -r 20-25 -o 精简版.pdf
```

> 说明：该功能复用 `split`（拆分所有页）+ `merge`（合并剩余页）实现。

---

## 4. doc/docx 转 PDF 并合并（doc2pdf）

```bash
python main.py doc2pdf <清单文件> [输出.pdf]
```

需要一个顺序清单文件（如 `order.txt`），每行一个文件名，按行顺序拼接。规则：

- 每行一个文件名，支持空行
- 以 `#` 开头的行是注释，会被忽略
- 相对路径相对于清单文件所在目录解析
- 清单中的 `.pdf` 文件直接使用（跳过转换）
- 其余扩展名按 Word 文档（doc/docx）处理，先转成 PDF 再拼接

### 示例

`order.txt`：

```
# 会议材料
01-封面.docx
02-议程.docx
03-附件目录.pdf
04-报告正文.docx
```

```bash
python main.py doc2pdf order.txt 会议材料.pdf
```

输出默认名为 `merged.pdf`。

---

## 注意事项

- 页码从 1 开始计数。
- 超出范围的页码会被忽略（`split` 自动截断到有效范围，`delete` 直接跳过）。
- `doc2pdf` 依赖 macOS 上的 Microsoft Word 和 Ghostscript，缺少任一环境时请改用本机其他转换方式。
