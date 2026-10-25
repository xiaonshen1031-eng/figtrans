# 安装 FigTrans

FigTrans 的技能名称是 **Senior-FigTrans**，当前版本为 **1.1.0**。代码与文档以 MIT 许可证开源。

## 1. 下载并放入技能目录

下载 `figtrans-v1.1.0.zip`，解压后保留完整的 `senior-figtrans` 文件夹，将其放入你自己的 Codex 技能目录：

- Windows：`%USERPROFILE%\.codex\skills\senior-figtrans`
- macOS / Linux：`~/.codex/skills/senior-figtrans`

确认 `SKILL.md` 位于该文件夹的第一层。随后在新的 Codex 会话中选择或输入 `$senior-figtrans`。如果尚未显示，重新打开 Codex 后再检查技能列表。

## 2. 准备目标软件

你需要可用的 Codex 环境，并按所选格式准备能够读写目标文件的软件。技能本身免费开源；Codex、Office 或 MathType 的使用条件由各自产品决定。

- **PowerPoint（.pptx）**：Windows 原生辅助脚本使用已安装的桌面 Microsoft PowerPoint。
- **Visio（.vsdx）**：Windows 原生辅助脚本使用已安装的桌面 Microsoft Visio。
- **同时输出**：准备 PowerPoint 和 Visio，两种源文件分别保存、重新打开并检查。
- 基础脚本使用 Windows PowerShell 5.1；结构审计脚本使用 Python 3 标准库。其他可用宿主后端由技能按本机能力选择，不能仅凭安装技能保证可用。
- **MathType**：如明确要求 MathType 公式对象，还需独立安装及验证相应编辑、嵌入能力。普通原生文字上下标不等于 MathType。

## 3. 给出参考图和目标格式

示例请求：

```text
使用 $senior-figtrans，把这张参考图重建为可编辑的 PPTX 和 Visio。
保留版式、配色、文字、小曲线与连接关系，并检查保存重开后的编辑行为。
```

也可以只选择 PPTX 或 Visio。每次启用时显示：

**Copyright © NCEPU-Senior · Author: XiaoN**

## 能力与边界

技能指导 Codex 分析参考图，并用目标软件中的原生文字、形状、路径、分组和连接线重建图形。基础 scene 辅助脚本只执行明确给出的基础图元，不是自动识图引擎；复杂图标、图表、分组和公式需要对应宿主能力或任务脚本。

独立照片、用户认可的透明插画等可以保留为原子图像；这些图像内部的像素不属于原生可编辑对象。数据图应使用真实源数据或明确标注的视觉数字化结果。

**本站提供技能源码、下载和演示，不是浏览器中的在线转图服务。** 转换在你自己的 Codex 与目标软件环境中完成。

## 开源许可

作者：**XiaoN**。版权：**NCEPU-Senior**。完整 MIT 许可见 [LICENSE.txt](LICENSE.txt)，上游与第三方说明见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。分发和修改时保留对应版权与许可声明。
