# FigTrans

**Copyright © NCEPU-Senior · Author: XiaoN**

把参考图片重建成可继续修改的 PowerPoint PPTX 与 Visio VSDX。MIT 开源。

- [访问网站与观看演示](https://xiaonshen1031-eng.github.io/figtrans/)
- [下载技能 v1.4.0](https://xiaonshen1031-eng.github.io/figtrans/downloads/figtrans-v1.4.0.zip)
- [安装说明](docs/docs/install.md)
- [技能入口](senior-figtrans/SKILL.md)

## 能做到什么

可选择 PPTX、VSDX 或两者。重建后的文字可直接修改，曲线保留原生路径，需要随节点移动的连接使用宿主的附着方式。已有源文件可做局部排版和箭头修正。
网站用 iTransformer 架构图演示重建。两个示例文件均含 42 个带文字的对象、41 个原生分组和 10 条双端附着连接。参考图对应 Liu et al., *iTransformer: Inverted Transformers Are Effective for Time Series Forecasting*, ICLR 2024, Fig. 4。 [ICLR 论文](https://proceedings.iclr.cc/paper_files/paper/2024/file/2ea18fdc667e0ef2ad82b2b4d65147ad-Paper-Conference.pdf#page=4)

- [参考图与重建结果对照、30 秒过程视频](https://xiaonshen1031-eng.github.io/figtrans/#demo)
- [下载示例 PPTX](https://xiaonshen1031-eng.github.io/figtrans/downloads/transformer-recreated.pptx)
- [下载示例 Visio](https://xiaonshen1031-eng.github.io/figtrans/downloads/transformer-recreated.vsdx)

PID 示例按参考图重建箭头轮廓，并按公式字形调整下标、积分号和分数线。可在 PowerPoint 中移动分组后的被控过程模块，查看箭头和进出连线是否跟随。参考图由 TravTigerEE 基于 Arturo Urquizo 的原图修订，参考图及 FigTrans 重建版按 [CC BY-SA 3.0](https://creativecommons.org/licenses/by-sa/3.0/) 提供。

- [对照 PID 参考图与重建结果](https://xiaonshen1031-eng.github.io/figtrans/#example-pid)
- [下载 PID 示例 PPTX](https://xiaonshen1031-eng.github.io/figtrans/downloads/pid-recreated.pptx)
- [Wikimedia Commons 参考图](https://commons.wikimedia.org/wiki/File:PID_en_updated_feedback.svg)

## v1.4.0

在 Windows 桌面 PowerPoint 中，可对已有原生对象做三项精修。将自定义矢量箭头和接收端口编入接收模块，按实际渲染的字形校准公式，再用单独的测试副本移动模块、修改文字并保存重开。结构校验读取保存后的对象关系，核对随动端点和应留在原处的端点。 [PowerPoint 精修流程](senior-figtrans/references/scene-helper.md#refine-an-existing-native-pptx)

## 使用

将 `senior-figtrans` 文件夹放入自己 Codex 的 skills 目录，然后附上参考图并调用 `$senior-figtrans`。
Codex 模型选择器同时提供 **Sol** 和 **Ultra** 时，建议复杂重建优先用 Sol，推理强度选 Ultra。
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
MIT 许可适用于技能代码及本站原创实现。演示参考图及其重建版遵循各示例标注的权利与许可。
站点由 GitHub Pages 发布，字体使用本机字体回退。
