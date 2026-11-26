# FigTrans

**Copyright © NCEPU-Senior · Author: XiaoN**

把参考图片重建成可继续修改的 PowerPoint PPTX 与 Visio VSDX。MIT 开源。

- [访问网站与观看演示](https://xiaonshen1031-eng.github.io/figtrans/)
- [下载技能 v1.2.0](https://xiaonshen1031-eng.github.io/figtrans/downloads/figtrans-v1.2.0.zip)
- [安装说明](docs/docs/install.md)
- [技能入口](senior-figtrans/SKILL.md)

## 能做到什么

可选择 PPTX、VSDX 或两者。重建后的文字可直接修改，曲线保留原生路径，需要随节点移动的连接使用宿主的附着方式。已有源文件可做局部排版和箭头修正。
网站用一张 Transformer 架构图演示重建。拖动对照滑杆可以比较原图和成图。两个示例文件均含 42 个带文字的对象、41 个原生分组和 10 条双端附着连接。

- [参考图与重建结果对照、30 秒过程视频](https://xiaonshen1031-eng.github.io/figtrans/#demo)
- [下载示例 PPTX](https://xiaonshen1031-eng.github.io/figtrans/downloads/transformer-recreated.pptx)
- [下载示例 Visio](https://xiaonshen1031-eng.github.io/figtrans/downloads/transformer-recreated.vsdx)

## v1.2.0

结构审计可识别重复对象 ID、重复附着端点以及连线附着到自身的错误。Visio 专用画布底板从原生重建证据中剔除。发布脚本从技能版本生成下载包及源码镜像，并核对 SHA-256。

## 使用

将 `senior-figtrans` 文件夹放入自己 Codex 的 skills 目录，然后附上参考图并调用 `$senior-figtrans`。
基础 COM 执行器需要 Windows PowerShell 5.1 和对应的桌面 Office 软件。
网站提供下载与演示，转换在使用者自己的工具环境运行。
照片/独立插画可能保留为单独图像；原生文本公式、OMML 与 MathType 是不同类型。
明确要求 MathType 时，需要可用的本地集成与独立验证。

## 仓库结构

- `senior-figtrans/`：技能源码、辅助脚本与参考文档。
- `docs/`：完整独立静态网站，GitHub Pages 从此目录发布。
- `tools/build_release.py`：构建发布文件，使用 `--check` 执行只读校验。

## 许可

技能采用 [MIT](LICENSE) 许可。保留 [第三方声明](THIRD_PARTY_NOTICES.md)。
MIT 许可适用于技能代码及本站原创实现。演示参考图由用户提供，参考图及其中内容保留各自权利。
站点由 GitHub Pages 发布，字体使用本机字体回退。
