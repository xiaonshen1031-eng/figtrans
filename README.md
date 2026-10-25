# FigTrans

**Copyright © NCEPU-Senior · Author: XiaoN**

把参考图片重建成可继续修改的 PowerPoint PPTX 与 Visio VSDX。MIT 开源。

- [访问网站与观看演示](https://xiaonshen1031-eng.github.io/figtrans/)
- [下载技能 v1.1.0](https://xiaonshen1031-eng.github.io/figtrans/downloads/figtrans-v1.1.0.zip)
- [安装说明](docs/docs/install.md)
- [技能入口](senior-figtrans/SKILL.md)

## 能做到什么

原生文字、形状、曲线和连接关系重建；支持 PPTX、Visio 或同时输出。
包含局部排版/箭头修正、分组检查、保存重开与代表性编辑验收流程。
网站示例使用用户提供的模型架构参考图，展示从参考图到可编辑对象的完整重建过程，实际绘制过程剪辑为 30 秒快放。

- [参考图与重建结果对照、30 秒过程视频](https://xiaonshen1031-eng.github.io/figtrans/#demo)
- [下载示例 PPTX](https://xiaonshen1031-eng.github.io/figtrans/downloads/transformer-recreated.pptx)
- [下载示例 Visio](https://xiaonshen1031-eng.github.io/figtrans/downloads/transformer-recreated.vsdx)

## 使用

将 `senior-figtrans` 文件夹放入自己 Codex 的 skills 目录，然后附上参考图并调用 `$senior-figtrans`。
基础 COM 执行器需要 Windows PowerShell 5.1 和对应的桌面 Office 软件。
网站提供下载与演示，转换在使用者自己的工具环境运行。
照片/独立插画可能保留为单独图像；原生文本公式、OMML 与 MathType 是不同类型。
明确要求 MathType 时，需要可用的本地集成与独立验证。

## 仓库结构

- `senior-figtrans/`：技能源码、辅助脚本与参考文档。
- `docs/`：完整独立静态网站，GitHub Pages 从此目录发布。

## 许可

技能采用 [MIT](LICENSE) 许可。保留 [第三方声明](THIRD_PARTY_NOTICES.md)。
MIT 许可适用于技能代码及本站原创实现。演示参考图由用户提供，参考图及其中内容保留各自权利。
网站使用本地字体回退与同站托管资产，不依赖境外脚本或字体 CDN。
免费托管与默认域名由 GitHub Pages 提供；各地网络可达性仍需按实际线路验证。
