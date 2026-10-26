# 实时对象编辑协议

保留原 129 个工具的输入参数和只读标记，新增 3 个统一工具，总数 132。先调用 `pagepilot_capabilities`，再调用 `pagepilot_objects(document, kind="schema")` 读取本机属性白名单和能力边界。

## 范围与对象

`document` 必须精确对应当前活动文档的完整路径。未保存文档使用上下文返回的名称，不猜测同名文件。`scope` 支持：

- `area=main`：正文。
- `area=header/footer, section=1, type=1`：节及页眉页脚；type 为 1 主要、2 首页、3 偶数页。
- `area=story, type, chain`：`kind=stories` 返回的文本区域和链索引，覆盖脚注、尾注、批注等 Word story。
- `area=shape, index`：正文浮动形状的文本范围。
- `area=page, index`：当前排版的页面范围；分页随编辑变化，页面设置属于节。
- `start/end`：所选 story 内的 Word 字符偏移；需要插入时 start=end。索引从 1 开始，字符位置从 Word 返回值读取。

对象种类包括 range、section、header、footer、table、cell、row、column、border、paragraph、style、field、bookmark、control、footnote、endnote、toc、inline、shape、chart、ole、equation。分页 count 为 1–200，使用 nextStart 继续。cell/row/column/border 使用 `scope.table`，单元格边框再指定物理 `scope.cell` 索引；复杂合并表的行列集合可能被 Word 拒绝，此时使用物理 Cells 枚举。

返回每个对象的句柄、可读属性及 unavailableProperties。范围文本超过 16,000 字符会明确截断，其他对象摘要最多 4,000 字符；精确读取需缩小范围。句柄不是永久 ID：任何正文、格式或其他 story 的变化均要求重新枚举，最多保存 256 个句柄。

## 修改

`pagepilot_object_edit(document, handle, operation, properties, parameters)`：

| 操作 | 对象 / 参数 |
|---|---|
| set | 按 schema 返回的精确属性路径稀疏修改，先验证全部属性 |
| text | range、cell、shape、header/footer、note、bookmark、control；parameters.text |
| delete | 表、行列、域、书签、控件、注释、目录、图片、形状、公式、自定义样式 |
| update / unlink | 域更新或取消链接；目录更新 |
| split_cell / merge_cell | rows/columns；或同一份枚举中的 other_handle |
| row_add / column_add | table 句柄，before 插入位置，省略则追加 |
| autofit | table，behavior 0 固定 / 1 内容 / 2 窗口 |
| columns_set | section，count 1–12 |
| tab_add / tab_clear | paragraph/range；position、alignment、leader |
| dropdown_add/delete/select | 组合框或下拉框；text/value 或 index |
| convert_inline / convert_float | Word 支持转换的图片/形状 |
| replace_image | 既有嵌入或浮动图片，path；保留受支持的布局属性 |
| linearize / build_up / equation_text | 既有 OMath 公式；text 为线性公式 |
| series_set / series_delete | chart；index/name/values/categories，或 index |

页眉页脚提供链接上一节、首页/奇偶页、页码起始和格式；section 提供纸张、方向、页边距、分节、分栏。表格提供行高、跨页、重复标题、单元格内边距、垂直对齐、底纹和逐边边框。对象文本局部修改使用明确范围，禁止整体覆盖包含单元格标记的范围。

`pagepilot_object_create` 支持 field、bookmark、control、footnote/endnote、toc、equation、textbox/shape、picture、table（包括明确单元格内的嵌套表）、section_break/page_break、style。所有参数详见工具 docstring / `tool-inventory.json`。

## 版本与恢复

Word 2010 基础类型可用；复选框同时要求文档兼容模式 ≥14。重复节内容控件要求 Word 2013+ 且兼容模式 ≥15。批注线程、控件颜色/外观要求新版 Word。实际 COM 属性可能因对象类型、文档格式或保护状态不可用，接口会报告，而不是把“工具存在”当作“此对象支持”。

修改前检查只读/保护、文档身份和整个 Flat OPC 的规范化哈希，创建包含未保存编辑的备份，并开启 Word 撤销记录。失败后尝试撤销并再次校验，结果为 completed、not_applied 或 recovery_required。必须检查 status；通信成功不代表文档操作成功。Word 自行变化的 proofing/rsid/段落会话标识不参与哈希。

## 明确边界

这是常用对象编辑面的扩展，不是完整 Word 对象模型。没有任意 COM 执行入口。OLE 提供枚举、尺寸、图标和链接属性，不启动嵌入程序或编辑其内部文档；图表提供类型、标题、图例、样式及数据系列，未覆盖所有轴和数据标签选项；OMath 提供整体线性表达式和排版转换，未暴露所有公式内部节点。未新增 Word 原生文档比较、任务级修订归属、复杂公式节点编辑、SmartArt 或数字一致性核对业务。上游工具的所有边界并未逐项重写或认证；复杂嵌套/合并结构仍受 Word 对象模型限制。
