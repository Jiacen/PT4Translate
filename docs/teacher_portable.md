# 教师绿色版试用

教师端是 Windows 绿色文件夹。老师解压后双击 `SITRAIN-Teacher.exe`，浏览器打开 `http://127.0.0.1:9000/`。需已安装 Microsoft Office；翻译需要联网。教材、校对和发布文件保存在同一文件夹下的 `sitrain_data/`，更新程序前必须备份该目录。

当前最新本地发布包为 **v0.1.2-fix2（2026-11-25）**：
`teacher_packages/v0.1.2-fix2/SITRAIN-Teacher-v0.1.2-fix2.zip`。
包含本轮重复导入、PPT 原生预览与格式保真、按文件记忆缩放的修复；网页和 EXE
产品版本显示 0.1.2-fix2，Windows 文件版本为 0.1.2.2。发布包使用全新空数据库及原先批准的共用 Qwen 试用配置。
独立解压副本已验证启动、模拟翻译、重启持久化、校对、完整下载和原生 PPT 渲染，
并验证旧任务中的未使用模板 QA 误报会清除，正文数字确实丢失时仍阻止发布，人工校对和修订号保持不变。
测试没有调用外部翻译 API，也没有改动发出的 ZIP。原始 v0.1.2 ZIP 仍保留，不覆盖。
本次新增 Word 真实页码、图片放大对照、逐项明确确认和持久进度计数。
全部确认后只生成一次同步 Word；重新进入识别不会清空已确认项。
Word 页码通过 `--index-word-images` 子进程在独立临时副本上添加零宽书签，
读取 Word 实际分页后缓存；不改动原件，不以图片顺序冒充页码。
详见 `docs/releases/v0.1.2-fix2.txt`；`teacher_packages/qa012f1b/` 仅为旧版验证副本，不作为交付件。

已有任务升级时，先退出旧版服务并备份数据，仅将修正版 EXE、`_internal` 和说明文件
复制到原来的程序目录；保留原 `sitrain_data`，不以新包空数据库覆盖。重启后刷新网页，
确认左下角为 `V0.1.2-fix2`。已有教材无需重新翻译，未处理完的图片可继续确认。

面向老师的完整操作步骤已写入 [教师绿色版 Pilot 使用手册 PDF](teacher_user_manual.pdf)，也可打开[离线网页版](teacher_user_manual.html)预览。

Pilot 发包时，ZIP 内自带可离线打开的 `使用手册.pdf`。老师先解压整个文件夹，再打开 PDF 按步骤操作。ZIP 也保留 `使用手册.html` 和配套截图，便于在浏览器中阅读。

## 构建运行包

PPT 预览必须调用真正的 PowerPoint 渲染。冻结后的程序通过自身的
`--render-ppt-preview` 子进程入口调用 `sitrain_app.ppt_render`；该分支必须先于
Web 服务、浏览器和数据目录初始化执行，不能把 EXE 当作 Python 去运行外置脚本。
渲染失败时报告错误，不使用缺失背景、图形和项目符号的简化重绘。

SITRAIN 默认仅替换 PPT 文字，保留原字体、字号、手动换行、段落/列表样式、
对象位置、裁剪、母版、图片和嵌入媒体。重建前后执行包结构保真检查；预览使用
`ppt-native-v2` 缓存目录，避免继续读取旧的简化预览。中英文字宽不同，固定字体
字号下的自动折行仍需在校对页检查。

在开发机的项目虚拟环境安装 PyInstaller 6.18 或更新版本，然后运行：

```powershell
.\.venv\Scripts\python.exe -m PyInstaller SITRAIN-Teacher.spec
```

生成目录为 `dist/SITRAIN-Teacher/`。运行目录只包含程序和依赖，不包含开发机的 `sitrain_data/`、`.venv/`、测试教材或 Git 目录。不要把开发目录整体压缩发给老师。

## 短期共用现有 API Key

管理员运行下列命令，会从当前本地数据库读取正在使用的 Qwen-MT Key，生成一份带全新空数据库的通用绿色 ZIP。给每位老师发送同一个 ZIP 即可；每台电脑运行后在自己的文件夹内保存教材和校对记录。命令不会打印 Key，也不会复制现有教材、术语库、项目或历史记录：

```powershell
.\.venv\Scripts\python.exe scripts\issue_teacher_packages.py --output teacher_packages\shared --source-db sitrain_data\sitrain.db Shared
```

ZIP 内的 Key 保存在本地 SQLite 数据库中，技术上可被拿到 ZIP 的人提取。仅向受邀教师发送，监测统一 API 额度，试用结束后轮换 Key。老师之间的数据和术语各自独立；若需要统一术语，可另发一份无凭证的术语表供导入。

## 长期统一翻译入口

如果提供可供老师访问的 HTTPS 地址，可运行 `sitrain_app.relay_server:app`。把上游 Qwen Key 放在服务器环境变量 `SITRAIN_RELAY_UPSTREAM_KEY`，老师凭证 JSON 文件路径放在 `SITRAIN_RELAY_TOKENS_FILE`；对外只开放反向代理的 HTTPS 端口。老师包改用：

```powershell
.\.venv\Scripts\python.exe scripts\issue_teacher_packages.py --output teacher_packages\relay --relay-url https://translate.example.org/v1 teacher01 teacher02
```

输出的 `relay-tokens.SERVER-ONLY.json` 只能放到网关服务器，不得发给老师。老师 ZIP 只含各自可撤销的入口凭证，不含上游 Qwen Key。
