# SITRAIN 教材智能翻译与校对平台（阶段一）

本目录是独立于 portable 参考包的新平台仓库，基线为 LinguaHaru `v6.8.1`（提交 `b3f73b8a8dc3cb7eeb120ec3e7b25efd5a60680d`，GPL-3.0）。Office 抽取和格式回填直接复用经测试的 LinguaHaru 内核；SITRAIN 业务、数据、校对与 Provider 层位于 `sitrain_app/`。

## 启动

首次双击 `Setup-SITRAIN.bat` 安装隔离依赖，之后双击 `Start-SITRAIN.bat`。浏览器会打开 `http://127.0.0.1:8080`。停止时运行 `Stop-SITRAIN.ps1`，或在服务窗口按 `Ctrl+C`。

数据默认保存在 `sitrain_data/`：SQLite 数据库、不可变原件、内部工作副本、管线中间文件和发布版本彼此分离。可以用 `SITRAIN_DATA_ROOT` 指向其他数据盘。

## 已实现的阶段一能力

- SITRAIN 品牌本地 Web UI；单文件和教材项目两个并列入口。
- 后端格式白名单、上传/解压规模限制、不安全 Office 容器检查。
- `project → chapter → files` 与 standalone task 数据模型，预留 owner/member。
- SQLite WAL 持久任务；服务重启后运行中任务恢复为 queued。
- `.ppt/.doc/.xls` 通过 Microsoft Office COM 转现代内部副本，保留原件并记录风险。
- 复用 LinguaHaru 的 PPTX/DOCX/XLSX 原位回填；TXT/MD/JSON/HTML/CSV/TSV 自有结构保护管线。
- PyMuPDF 数字/混合/扫描 PDF 分流；PP-OCRv6 Small CPU OCR、坐标/置信度保存及可逆遮盖回填。
- Office 内嵌 PNG/JPEG/BMP/WebP/TIFF 进入同一 OCR 管线，校对后在包内原媒体路径回填，不改变 Office 关系、尺寸与裁剪。
- Qwen-MT、DeepSeek、OpenAI-compatible 和自定义 Provider；严格返回键验证、重试、超时与服务端 Key 存储。
- NFKC/casefold/连字符规范化、源术语别名（仅用于识别同一源术语的不同写法）、最长匹配、翻译前占位、翻译后强校验、数字 QA。
- 全局术语库支持下载 `.xlsx` 空模板或现有术语表，并可批量导入；导入不会覆盖已有源术语，会返回新增、跳过与无效行统计。
- 统一在线校对、乐观锁自动保存、确认译文写入项目翻译记忆。
- 原始/工作/发布版本、单文件下载、章节 ZIP 与 SHA-256 交付清单。
- PPT—Word 按序自动映射、低置信度阻止同步、PowerPoint 完整页渲染及 Word 底层媒体替换（保留 Word 对象的裁剪/尺寸/位置）。

## 当前机器上的验收边界

仓库未提供规划指定的四个 SITRAIN 真实回归样本，因此自动测试使用生成的 PPTX 和文本样本验证完整上传、翻译、校对、回填、发布与下载。真实 PPT+Word 映射、旧格式和扫描 PDF 仍需加入规划中的样本后执行材料级验收。

隔离环境已安装 RapidOCR 3.9.2、ONNX Runtime 和 PP-OCRv6 Small，默认使用 CPU。PP-OCRv6 Medium 可在管理设置中作为后续高精度模型扩展；BabelDOC 尚未安装，当前 PDF 使用 PyMuPDF 坐标回填路线。Office 中的 EMF/WMF OCR 仍需先通过 Office 渲染为栅格预览；PPT—Word 同步本身直接支持 EMF 媒体替换。

## 测试

```powershell
.\.venv\Scripts\python.exe -m pytest tests\test_sitrain_phase1.py -q
```

上游 Office 回归套件也应继续运行，例如 `tests/test_corpus_pptx.py`、`tests/test_corpus_docx.py` 和 `tests/test_corpus_xlsx.py`。
