# 完整自定义模板工作流

## Clone 后放入模板

```text
templates/
  university/
    defense-blue.pptx
    defense-red.pptx
  company/
    report.pptx
```

支持任意嵌套目录，所有完整PPTX均独立入选，不修改或裁剪文件。公开仓库不附带任何用户学校/公司模板。

```sh
python scripts/template_library.py list
python scripts/template_library.py select TEMPLATE_ID work.pptx
python scripts/layout_inspector.py work.pptx --output all-pages.json --table
```

list 会重新扫描 templates/，输出 ID、完整页数与相对路径，写入本机 templates/library.json。ID由相对路径生成。select 验证哈希后复制整份原稿。

## 默认保留全部页面

空计划 `{}` 逐字节复制全部内容；不指定 mode 时默认 full-deck。

```json
{"edits":[{"page":0,"texts":{"13":"新的汇报标题"}}]}
```

可在同一页计划替换既有图片并补充逐页讲稿和来源：
```json
{
  "edits": [{
    "page": 2,
    "texts": {"13": "研究流程"},
    "images": {"18": "assets/research-flow.png"},
    "notes": "先介绍研究问题，再说明流程图中的三个步骤。"
  }],
  "sources": [
    {"page": 2, "claim": "流程依据的事实或数据", "title": "来源标题", "url": "https://example.org/source", "accessed": "2026-09-29"}
  ],
  "image_sources": [
    {"page": 2, "shape_id": "18", "title": "图片来源或 AI 原创插画", "url": "https://example.org/image", "license": "许可说明"}
  ]
}
```

`images` 只能指定原页已经存在的图片 shape_id；引擎替换原图片引用并保留其尺寸、位置和裁剪，不创建叠加形状。AI 生成插画的 url 标为 `AI-generated`，在 image_sources 记录生成描述和工具；联网图片必须记录原始页面和许可。没有可验证出处的事实写入报告的“待核实项”，禁止编造引用。

每页 `notes` 是讲稿草案，应自然对应该页内容，扩展图表读法、转场和讲解重点，不重复逐字读标题。未填写的页，Agent 应根据最终页面内容补上讲稿，并保持已有备注除非用户要求改动。成品旁必须有 `信息源报告.md`，按页列出来源、链接和访问日期；无外部引用的页面标注“无外部来源（内容来自用户材料/分析）”。图片授权、事实来源与讲稿备注分开记录。

page 从0开始，shape_id 从探测器读取，13不是通用值。没有出现在计划中的页和内容保持。只改字时仅目标slide XML变化，其他部件不重新序列化。

```sh
python scripts/template_engine.py work.pptx plan.json finished.pptx
python scripts/validate_output.py finished.pptx --report validation.json
```

## 按需求增删改

`{"delete_pages":[3,7]}` 删除原稿第4、8页，其余保留。有保留页链接到待删页时会报错，需先处理链接。

sequence 表示最终**完整**页顺序。例如四页原稿变五页：
```json
{"sequence":[0,1,{"source_page":1,"texts":{"5":"新增页文字"}},2,3]}
```
整数均指原稿页序，重复则复制。不能照抄到54页原稿；未出现在 sequence 的页会删除。sequence 与 delete_pages 不可并用。

复制页保留形状、主题和布局；图表、工作簿、备注等关系部件独立复制，静态图像和母版可共享。复杂对象仍需用目标PowerPoint实际打开验证。texts API 不修改图表数据和任意图片。

## 明确重新组稿时才抽页

`mode: sample-pages` 使用 allowed_pages 与 slides/page/texts，只输出指定原页。`mode: layouts` 使用 allowed_layouts 与 slides/layout/title/body/images，从母版新建页面。两个模式均须显式指定，不是默认整稿编辑方式。

```json
{"mode":"layouts","allowed_layouts":[[0,1]],"slides":[{"layout":[0,1],"title":"研究问题","body":["研究背景"]}]}
```
母版索引用 `[master_index,layout_index]`。占位符使用 idx 键，普通文本框使用 shape_id，两者不能混淆。图片使用 PICTURE 占位符的 insert_picture，不叠加新的框。

## 校验

核对原文件哈希、页数变化、删增页清单、所有改字与保留页；导出图片逐页检查。文字容量只是估算，字体缺失可能改变换行。长文应按需求明确安排续页，不自动覆盖或删掉原页。
