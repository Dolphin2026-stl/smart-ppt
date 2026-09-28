# 自由风格模式

| 用户偏好 | ID | 视觉规范 | 推荐场景 |
|---|---|---|---|
| 简约 | minimal | 白底、无衬线、绿色细线、留白 | 工作汇报 |
| 炫酷/科技 | tech-blue | 深蓝底、青色强调、稀疏网格 | 技术展示 |
| 学术 | academic | 暖白底、紫色、衬线标题 | 学术答辩 |
| 红色/节庆 | festive-red | 暖白底、深红文字与金色强调 | 表彰、庆典 |
| 传统 | traditional | 纸色底、墨绿、红色印章意象 | 文化主题 |
| 轻松 | casual | 浅亮背景、圆形装饰、暖色强调 | 课程展示 |

内置风格是可编辑的基础视觉系统，不承诺生成复杂插画。设计优先级：内容层级 → 留白和对齐 → 视觉节奏 → 装饰。保持单页一个主要信息，避免把正文写成堆积的口号。长内容先拆页，不靠缩小字号解决。

```json
{"slides":[
  {"kind":"cover","title":"研究与表达","body":["学术汇报 · 2026"]},
  {"kind":"two-column","title":"方法与验证","body":["方法说明","验证依据"]},
  {"kind":"ending","title":"谢谢","body":["欢迎交流"]}
]}
```
`kind` 支持 cover / agenda / content / two-column / three-column / chart / ending。chart 提供图片槽位，使用 images 数组填入真实图表截图；数据可编辑性不是当前引擎的功能。

```sh
python scripts/style_engine.py examples/demo-deck/content.json output.pptx --style academic
python scripts/validate_output.py output.pptx --report validation.json
```

## 自定义风格

复制 `assets/styles/minimal.json`，修改 `colors` 五个六位 HEX 值（不带 #）、`fonts` 字体与层级、`layouts` 的角色和英寸坐标、`decorations.kind`。保持七种版式键名；支持 line / grid / seal / round / band 装饰。通过 `--style path/to/custom.json` 加载。

槽位 role 为 title / body / picture。页面固定 13.333333 × 7.5 英寸；自定义框必须在页面内。默认封面 44pt、标题 32pt、正文 24pt。脚本不会自动检测或替换系统字体；根据 assets/fonts/mapping.json 选择已安装且有中文字形的字体。指定中文无衬线字体比依赖系统自动回退更可靠。

