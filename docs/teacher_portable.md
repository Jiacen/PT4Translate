# 教师绿色版试用

教师端是 Windows 绿色文件夹。老师解压后双击 `SITRAIN-Teacher.exe`，浏览器打开 `http://127.0.0.1:9000/`。需已安装 Microsoft Office；翻译需要联网。教材、校对和发布文件保存在同一文件夹下的 `sitrain_data/`，更新程序前必须备份该目录。

## 构建运行包

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
