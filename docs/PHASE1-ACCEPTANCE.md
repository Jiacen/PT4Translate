# 阶段一验收记录

日期：2026-09-02

## 已通过的自动化验收

- 单文件与教材项目/章节双入口及数据归属。
- TXT 完整上传、mock Provider 翻译、在线校对、TM、发布及下载。
- PPTX 通过 LinguaHaru 内核抽取、人工校对、原位回填、打开验证及页面预览。
- PPTX 内嵌图片通过 PP-OCRv6 Small CPU OCR，保存区域坐标/置信度，校对后写回原媒体关系。
- 不支持扩展名由后端返回 415；Provider API 不返回密钥。
- 上游 PPTX、DOCX、XLSX 结构回归 20 项全部通过。
- 最终测试汇总：26 passed；唯一 warning 来自上游测试夹具重复注入 `notesMaster1.xml`，不是平台输出错误。
- JavaScript 语法检查与 `git diff --check` 通过。
- 浏览器实际检查登录页、工作台、Provider 设置与资源加载；控制台无 error/warn。

## 需要真实材料完成的验收

规划指定的四个 SITRAIN 回归文件不在当前仓库或桌面中，因此以下项目已有实现但尚不能声明“真实教材样本验收通过”：

- `03_User_interface_LE.pptx` + `03_User_interface.docx` 的页面映射、裁剪保留和同步后分页；
- `03_Commissioning_Distributed_IO(cn).pptx` + `.docx` 的中英成品对比；
- 真实 `.ppt/.doc/.xls` 的 Office COM 转换风险报告；
- 文字型、扫描型、混合型真实 PDF 的视觉回填质量；
- 含 EMF/WMF 截图文字的 Office OCR（栅格图片 OCR 已通过，EMF/WMF 目前用于 PPT—Word 媒体同步，不直接 OCR）。

收到这些文件后，应作为固定回归集放入独立的非 Git 测试资料目录，通过网页完整跑一遍并记录每页 QA；不得仅以合成样本替代材料级验收。
