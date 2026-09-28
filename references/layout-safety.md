# 布局保护原理与排查

## 继承链

slide placeholder 通过 placeholder idx 对应 slideLayout placeholder；layout 的位置和主题样式还可以继承 slideMaster。idx 可以不连续，不能用列表位置代替。`slide.placeholders[idx]` 按 idx 查找，`layout.placeholders.get(idx)` 才是版式侧对应查找方法。多母版时始终使用当前 slide.slide_layout，不假设第一个母版。

只改变 text_frame 内容不会主动设置位置。对 left/top/width/height 的赋值会在 XML 的 p:spPr 下创建 a:xfrm，覆盖继承。模板里可能原本就有显式值：存在 xfrm 不一定意味着本次操作导致偏移；校验同时比较有效几何和生成基线。

`insert_picture()` 会把图片槽位的 p:sp 替换为 p:pic，保留 idx；不能继续使用旧 shape 对象。该操作可能生成显式变换，因此不能以“有 xfrm”直接判错。实际位置必须与插图前一致。

## 图片比例

保持框的宽高比与保持整张图片可见是两回事。图片槽位通常以裁剪填满框：图片不拉伸，但边缘可能被裁掉。校验采用图片 DPI、crop_left/right/top/bottom 计算可见区域比例，再与框比例比较。不要将正常裁剪误报成变形。

## 常见错位

- 插入图片后手动 resize：覆盖继承。删除相关赋值，重新从原模板生成。
- 在 placeholder 上叠加 textbox：原 placeholder 留在原处，新框破坏层级。通过已有 text_frame 写入。
- idx 当成第几个占位符：尤其在 idx=10、12 的模板出错。用键查找。
- 多母版 index 混用：使用二元索引与明确白名单。
- 用 layout 复制实际示例页：会丢掉 slide 本身的装饰。用 sample-pages 原页克隆。
- 长文本被截断：估算不是渲染，优先拆页/精简，禁止改模板框。
- 字体缺失：文字宽度和换行会变。安装有授权的字体或经用户同意选择替代字体，再重新渲染。

## 排查程序

1. 保留原文件，记录模板、内容计划和输出文件。
2. 查看 layouts.json 及 audit 的基线；位置单位为整数 EMU（914400 EMU = 1 英寸）。
3. 比较占位符 idx 集合、所有框几何与显式变换。
4. 样例页模式比较所有 shape_id 集合与几何，不仅比较 placeholder。
5. 在实际目标软件上逐页渲染，检查换行、溢出、遮挡、logo、图片主体和页序。

## 范围与限制

输出校验能发现结构变化，不证明内容正确，也不能完全证明文字无溢出。未带 audit 时无法证明相对源文件没有改动。示例页克隆使用 python-pptx 私有 XML API，已集中在 template_engine.py；升级 python-pptx 后必须重跑真实模板测试。图表/OLE/跨页链接不做不完整复制。

开发参考入口（未将当前可访问性当作已验证事实）：
- https://python-pptx.readthedocs.io/en/latest/user/placeholders-understanding.html
- https://python-pptx.readthedocs.io/en/latest/user/placeholders-using.html
- https://agentskills.io/specification
