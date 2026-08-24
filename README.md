# TOPCON 数据解析与图片复制工具

用于解析 TOPCON 导出的数据文件。工具可以读取记录清单和 `HIST0010` 图片索引，
按指定 ID 查找关联信息，递归定位 JPG/PNG 图片，将匹配的数据复制到独立目录，
并生成 Excel 结果表。

## 日常使用（推荐）

首次使用或依赖有变化时执行：

```powershell
uv sync
```

启动图形界面：

```powershell
uv run python main.py
```

在界面中：

1. 选择 TOPCON 数据目录。
2. 图片不在同一目录时，单独选择图片根目录；否则留空。
3. 将待查询的 ID 粘贴到输入框，或从 TXT、CSV、XLSX 文件导入。
4. 选择结果输出目录。
5. 点击“查询、拷贝图片并生成 Excel”。

每次执行会创建独立的 `TOPCON核查_时间戳` 文件夹：

```text
TOPCON核查_20260824_153000_000000/
├─ 核查反馈.xlsx
└─ 图片/
   ├─ ID-1_2016-01-02_女/
   │  ├─ IM000001.JPG
   │  └─ IM000002.JPG
   └─ ID-2_2018-03-04_男/
      └─ IM000010.PNG
```

图片复制采用安全策略：内容相同的文件跳过，同名但内容不同的文件自动添加序号，
不会覆盖已有文件。记录子目录使用“ID_出生日期_性别”的格式，性别统一为
“男/女”；缺失值使用“出生日期未知/性别未知”。文件复制使用 `copy2`，保留原文件修改时间。

### ID 前导零

ID 始终按文本处理。推荐从 Excel 以文本形式复制，保留前导零。界面提供
“忽略前导零的唯一匹配”选项，默认关闭；只有候选记录唯一时才允许这种匹配。

## Excel 反馈表

`核查反馈.xlsx` 包含四张工作表：

- `查询汇总`：每个请求 ID 一行，显示匹配状态、关联信息、图片数量和复制数量。
- `图片明细`：每张图片一行，包含检查日期、图片名、源路径、复制状态和目标路径。
- `异常与缺失`：集中列出未找到的 ID、缺少图片索引、图片文件缺失及复制失败。
- `拷贝记录`：逐文件审计源路径、目标路径、判重和复制结果。

ID 列被强制保存为 Excel 文本，避免前导零丢失。日期按 `yyyy-mm-dd` 保存，表头可筛选
并冻结，已找到、需关注和失败状态使用不同颜色。

## 命令行批量使用

```powershell
uv run python topcon_lookup.py D:\topcon_data `
  --ids-file D:\work\ids.xlsx `
  --image-root D:\topcon_images `
  --copy-to D:\result\images `
  --output D:\result\核查反馈.xlsx
```

也可以直接传入少量 ID：

```powershell
uv run python topcon_lookup.py data --ids "000001,000002" --output output\核查反馈.xlsx
```

## 完整数据解析

项目使用 `uv` 管理 Python 3.13 环境，依赖由 `pyproject.toml` 和 `uv.lock`
统一锁定，无需手工安装。

```powershell
uv run python parse_topcon.py data --output-dir output
```

如果本机配置的 uv 缓存目录不可写，可以只为当前终端指定项目缓存：

```powershell
$env:UV_CACHE_DIR = "$PWD\.uv-cache"
uv run python parse_topcon.py data --output-dir output
```

若数据目录包含多个 `.TXT` 文件，可以明确指定记录清单文件：

```powershell
uv run python parse_topcon.py data --patients-file 2026.7.31.TXT --output-dir output
```

### 输出文件

- `output/patients.csv`：记录清单中的 6 个原始字段，转换为 UTF-8 with BOM，
  可直接用 Excel 打开。
- `output/image_index.csv`：`HIST0010` 中每张图片对应的一条索引记录。
- `output/parse_report.json`：记录数、图片扩展名、交叉匹配与格式校验统计。

`image_index.csv` 中已确认的主要列包括：

- `patient_id`、`patient_name`、`sex`、`birth_date`、`registered_date`
- `last_capture_date`：每条记录的最后一次拍摄日期
- `capture_date`、`capture_time`、`capture_source`
- `image_number_in_exam`：同次检查中的 `#1`、`#2` 等序号
- `image_filename`：`IM######.JPG` 或 `IM######.PNG`
- `image_sequence`：从图片名提取的数字序号

无法可靠解释的尾部字段使用带偏移量的列名，并同时保存在
`unknown_tail_hex` 中，避免逆向解析时丢失信息。

> 注意：Excel、CSV 和复制的图片可能包含敏感信息。`output/` 已加入 `.gitignore`，
> 请根据实际数据管理要求妥善保存和处理。

## 运行测试

```powershell
uv run python -m unittest discover -s tests -v
```

## Windows EXE 打包与分发

先同步开发依赖，再使用仓库中的 PyInstaller 配置生成单文件版本：

```powershell
uv sync
uv run pyinstaller --noconfirm --clean ".\TOPCON数据核查工具.spec"
```

产物位于 `dist\TOPCON数据核查工具.exe`。单文件版本只需复制这个 EXE，
不需要同时复制 `_internal` 目录；首次启动可能因解压依赖而稍慢。

如果需要更快启动、也更方便排查缺少组件的问题，可生成目录版：

```powershell
uv run pyinstaller --noconfirm --clean --onedir --windowed `
  --icon ".\assets\topcon_lookup.ico" `
  --name "TOPCON数据核查工具" main.py
```

目录版必须复制整个 `dist\TOPCON数据核查工具\` 文件夹，EXE 和 `_internal`
缺一不可。`dist/`、`build/` 和本地数据目录都不会提交到 Git。
