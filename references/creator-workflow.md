# 创作者模式：从素材到可复用模板

## 最少输入

至少提供一张背景图或配色参考图、一张 logo，以及字体偏好。已在上下文给出则不重复询问。字体文件不是必填；不能擅自分发或安装用户字体。

```text
assets/
  backgrounds/       # 封面背景，优先 16:9 PNG/JPEG
  logos/             # 官方透明 PNG 标识，不拉伸
  icons/             # 可选图标，扫描归档，目前不会自动摆放
  palettes/          # 配色参考 PNG/JPEG
  fonts/             # mapping.json；可选有授权的 TTF/OTF
  styles/            # 内置主题，勿混入用户配色截图
```
也可以传入项目之外的独立素材目录。扫描按目录名和文件名分类；无法分类的图像放进 unclassified 清单，不盲目当背景。SVG 请先用可靠转换器转为 PNG；引擎不执行素材中的程序。多个同类素材按路径排序选第一张，并在 `.assets.json` 中记录；需要精确控制时为本次任务建立只含已选素材的独立目录。

## 构建过程

1. 扫描背景、logo、图标、配色参考、字体文件。
2. 优先取 palettes 中第一张图；否则取背景。缩小到 160×160 以内，用 Pillow median-cut 量化为五个代表色，并按像素占比排序。透明像素先合成到白底。
3. 最深颜色作为主色，饱和程度最高的颜色作为强调色；文本和背景保留基础主题的可读配色。抽色不是品牌色识别，用户提供的品牌色值优先。
4. 用空白页初始摆放形状，构建封面、目录、一栏、两栏、三栏、图表图像区、结尾七个版式。
5. 将文字/图片槽位写入真实 slideLayout 的 placeholder；布局坐标只在首次创建时设置并记录。移除构建时的临时页，生成零内容页的可复用 `.pptx` 模板。
6. 输出 `.layouts.json`、`.creation.json` 和 `.assets.json`，之后通过模板引擎生成演示样张。

```sh
python scripts/template_creator.py assets organization-template.pptx --font "Microsoft YaHei"
python scripts/layout_inspector.py organization-template.pptx --output layouts.json --table
python scripts/template_engine.py organization-template.pptx plan.json preview.pptx
```
图表版式目前提供图像占位符和文字区，不会凭空生成数据或可编辑图表；需要原生图表时由宿主按真实数据另行制作并审核。图标与字体文件会扫描入清单但不自动嵌入，报告对此作明确说明。

## 预览

先生成七个版式各一页，用 PowerPoint 或 LibreOffice 打开并导出图片或 PDF；检查完整 logo、背景比例、对比度、字体可用性及中文换行。Windows 上可使用项目的可选 `scripts/render_powerpoint.ps1`，必须安装 PowerPoint。只有经确认的模板才继续生成正式内容。没有渲染器时交付模板与清单，明确未做视觉验证。

背景以保留完整比例的方式放在封面右侧；不铺满文字背后，避免随机图片破坏可读性。可自定义 JSON 中坐标进行新模板创作，创建后进入模板套用阶段则冻结坐标。
