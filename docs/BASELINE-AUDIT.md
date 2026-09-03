# 基线审计（2026-09-02）

- portable 目录名为 `LinguaHaru-6.8.2-web-portable`，但包内 `version.json` 是 `6.8.1`。
- 上游标签 `v6.8.1` 指向提交 `b3f73b8a8dc3cb7eeb120ec3e7b25efd5a60680d`；本仓库从该标签创建 `sitrain-phase-1` 分支。
- 上游根目录包含 GPL-3.0 `LICENSE`；portable 本身没有根许可证副本，但关于页声明 GPL-3.0。
- portable 原生注册 DOCX/PPTX/XLSX/PDF 等现代格式，不注册 PPT/DOC/XLS；阶段一新增显式 COM 转换层。
- portable 的 OCR/PDF 插件只有清单，运行环境中未安装 RapidOCR/PaddleOCR/BabelDOC。
- 本机发现 PowerPoint 及 Poppler，未发现 LibreOffice；旧 Office 转换采用已安装的 Microsoft Office COM。
- 工作区和桌面未找到规划列出的四个 SITRAIN 回归样本，因此当前不能给出真实教材版式与 PPT—Word 映射的最终验收结论。

## 可复用边界

```text
SITRAIN Web / SQLite / projects / review / providers / terminology / TM
                              │
                              ▼
                    stable translation units
                              │
          ┌───────────────────┼───────────────────┐
          ▼                   ▼                   ▼
 LinguaHaru PPTX        LinguaHaru DOCX      LinguaHaru XLSX
 extract + writeback    extract + writeback  extract + writeback
          │                   │                   │
          └───────────────────┴───────────────────┘
                              ▼
                  immutable published versions
```

PDF/OCR 和 PPT—Word 同步通过 SITRAIN 独立适配层接入，不把项目数据模型耦合进格式处理器。
