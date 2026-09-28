# 科研PPT设计系统 v1.0

这是一套可复用的科研汇报工作流：用MD梳理表达关系，检索适合的页面原型，提取可编辑模板，再结合真实材料制作PPT。

## 现在包含什么

- `research-ppt/SKILL.md`：供助手调用的工作流程。
- `references/`：内容关系方法论、41类页面原型、选用规则、设计规范和备注字段约定。
- `assets/template-library.pptx`：84页可编辑模板，含2页备用页。
- `assets/catalog.json`：从简明备注自动生成的模板索引。
- `scripts/template_catalog.py`：读取备注、检索候选、提取页面和检查索引。
- `assets/report-brief.md`：可选的汇报需求输入模板。
- `assets/slide-cards.md`：由AI填写的逐页汇报卡片模板，连接内容规划与模板匹配。
- `references/report-cards.md`：卡片字段、证据状态和使用规则。
- `examples/logic-test-input.md`：可直接测试的虚构科研文案；生成后用 `examples/logic-test-expected.md` 对照验收。

PPT备注统一为九行，最多约200字：模板ID、原型、关系、结构、适用、槽位、注意、分组、来源。打开PPT底部“备注”即可查看。详细理论在MD中，备注只保留选用所需信息。

## 直接使用

在当前项目中可以说：

> 请按科研PPT设计系统处理这些材料，整理成15页组会报告。先判断每页的表达目标与主要关系，匹配模板，再制作可编辑PPT。

技能安装并加载后可以使用：

> 请使用 $research-ppt。我的研究包含三种数据输入，经融合得到预测结果。请推荐模板，说明选择理由，并提取合适的可编辑页面。

只需要梳理思路时说“仅整理逐页逻辑与模板建议”；需要最终报告时明确提供材料、用途及大致页数或时长。技能不会把“先做大纲”变成必须等待确认的流程。

## 带到其他设备或分享

复制整个 `research-ppt` 文件夹即可。它不依赖当前项目的绝对路径，也不需要附带原始参考图片或临时分析目录。

要让Codex发现这个技能，将文件夹放入该设备的个人技能目录 `${CODEX_HOME}/skills`；未设置CODEX_HOME时为 `~/.codex/skills`。新会话加载后使用 `$research-ppt`。也可以不安装，向助手提供该文件夹的 `SKILL.md` 路径，让其按文件说明使用。

`research-ppt-v1.0.zip`包含可迁移的技能文件夹及本说明。模板中的图片和数据区域仍为空白，使用者填写自己的研究材料；来源标签用于追溯版式案例。

## 自动化能做什么

脚本已经实现读取备注、按关系或原型检索，以及保持原生对象提取指定页面。理解研究材料、判断主要关系和实际填充版式由助手与演示文稿工具配合完成。检索分数不是逻辑正确率，不把关键词命中当作最终选页理由。

在 `research-ppt` 文件夹中运行，Python 3标准库即可：

```text
python scripts/template_catalog.py search --relation 汇聚 --query 多源输入 --top 5
python scripts/template_catalog.py search --prototype T09
python scripts/template_catalog.py extract --ids TPL-P09,TPL-N25,TPL-P33 --output selected.pptx
python scripts/template_catalog.py validate
```

提取文件是可编辑页面组合，还没有自动填充研究内容。脚本不会覆盖现有输出。同一报告多次采用同一版式时，可在编辑器中复制该页。

## 以后如何扩充

在技能包的 `assets/template-library.pptx` 中添加或修改模板，按 `references/metadata-schema.md` 填写备注。每张具体模板分配唯一ID，原型ID可以复用。重排页面时保留模板ID。

修改后重新生成索引：

```text
python scripts/template_catalog.py index --output assets/catalog.json --replace
python scripts/template_catalog.py validate
```

新增结构若无法归入现有原型，再更新页面原型文档。不要因为添加一个新例子就新增逻辑类别。修改原件前自行留一份版本备份；重新打包或安装时同步更新完整文件夹。

当前项目 `个人模板/科研PPT总模板库` 中的PPT与技能包内PPT在发布时一致。后续建议以技能包内的模板为维护入口，发布时再同步项目中的使用副本。
