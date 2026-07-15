# 人民法院在线材料助手

一个面向中国网络小额纠纷的 Codex Skill：从聊天记录、订单、支付凭证、平台记录和身份线索中提取事实，先判断应当优先准备民事起诉、报警材料还是人工复核，再生成结构化证据目录和可提交的诉讼材料。

> 本项目用于个人自助整理材料，不提供律师代理，不代替法院、公安机关或律师的专业判断，也不保证立案、胜诉或追回款项。

## 它能做什么

- 自动读取用户已经提供的截图、聊天、订单、账单和身份材料，尽量不重复提问。
- 每次只询问一个真正影响下一步的关键问题，并优先使用简单选择题。
- 区分材料中明确显示、用户确认、合理推测和仍然缺失的信息。
- 初步分流为民事起诉、报警准备或人工复核。
- 按证据来源、证明对象和证明内容分类编号。
- 生成民事起诉状、证据目录、编号证据册、上传清单等 DOCX/PDF 文件。
- 在生成最终版前检查身份、请求、管辖、证据来源、签名日期和平台上传要求。
- 保留原始电子证据，不替用户签名、盖章或直接提交。

## 当前适用范围

V1 仅在以下条件同时满足时生成民事起诉材料：

- 纠纷属于网络商品买卖；
- 买卖双方均为个人；
- 被告能够与他人区分；
- 用户提出了明确的民事请求；
- 基本事实、金额、证据来源和拟提交法院已经确认。

网络服务、借贷、投资返利、虚拟货币、账号租赁、多名受害者、跨境交易或其他复杂情形，仅生成分流、证据保全和待补材料，不直接生成最终民事起诉状。

## 安装

Codex 支持把个人 Skill 放在 `$HOME/.agents/skills`，也可以把项目专用 Skill 放在仓库的 `.agents/skills`。详见 [OpenAI Codex Skills 文档](https://learn.chatgpt.com/docs/customization/overview#skills)。

个人安装：

```bash
mkdir -p "$HOME/.agents/skills"
git clone https://github.com/carrieren0126/prepare-cn-online-court-materials.git \
  "$HOME/.agents/skills/prepare-cn-online-court-materials"
```

项目内安装：

```bash
mkdir -p .agents/skills
git clone https://github.com/carrieren0126/prepare-cn-online-court-materials.git \
  .agents/skills/prepare-cn-online-court-materials
```

## 使用方法

在 Codex 中附上案件材料，然后直接说：

```text
使用 $prepare-cn-online-court-materials 整理这些网络交易材料。
先从文件中提取信息，只问我无法从材料中确定且会影响下一步的问题。
最后生成可供人民法院在线服务使用的 PDF 材料。
```

也可以使用更短的请求：

```text
帮我把这些闲鱼订单、聊天和付款记录整理成证据目录与诉讼材料。
```

Skill 会先盘点文件、提取事实并判断路径。遇到缺失信息时，它会一次只问一个问题；满足最终版条件后，再生成正式材料包。

## 典型输出

民事路径通常包括：

1. 材料准备情况及待确认事项
2. 民事起诉状
3. 证据目录
4. 编号证据材料
5. 人民法院在线服务上传清单
6. 原告身份证明材料

报警路径通常包括事实陈述、损失和付款表、涉案账号线索表、证据目录及证据册。人工复核路径仅生成材料盘点、缺失项和下一步建议。

## 独立运行生成脚本

如需在本地直接运行生成器，请先准备 Python 3，并安装：

```bash
python -m pip install python-docx pillow pypdf reportlab
```

生成匹配的 DOCX/PDF 还需要安装 LibreOffice，并确保系统可以调用 `soffice` 或 `libreoffice`。

复制并填写 [`assets/case-data.example.json`](assets/case-data.example.json)，然后运行：

```bash
python scripts/build_case_packet.py case.json --output-dir packet --mode draft
python scripts/build_case_packet.py case.json --output-dir packet --mode final
```

`final` 模式具有强制门槛：关键事实仍为推测、法院或当事人未确认、存在重大待处理问题、未检查在线平台要求等情况都会阻止生成最终版。

## 证据准备原则

- 保留原始手机、账号、文件和平台可下载记录。
- 截图尽量包含账号主页、时间、完整上下文和连续页面。
- 优先补充支付平台官方账单、电子回单或交易快照。
- 不修改原始证据；标注只放在副本上并注明。
- 未脱敏材料只提交到法院、公安机关等依法有权接收的渠道。
- 最终上传前由本人核对并完成签名或平台电子签名。

## 官方依据

本 Skill 的规则索引优先使用政府、立法机关和法院官方网站，包括：

- [最高人民法院：民事起诉状、答辩状示范文本](https://www.court.gov.cn/zixun/xiangqing/468671.html)
- [最高人民法院：关于民事诉讼证据的若干规定](https://gongbao.court.gov.cn/Details/0c15319f2bdbabb8e398035f775385.html)
- [最高人民法院：人民法院在线诉讼规则](https://gongbao.court.gov.cn/Details/ac5f36e345967c22e0a2ac4fbeb0a6.html)
- [最高人民法院：人民法院在线运行规则](https://www.court.gov.cn/zixun/xiangqing/346471.html)

完整索引和核验方法见 [`references/official-sources.md`](references/official-sources.md)。法律规则、模板和在线服务要求可能调整，生成最终版前仍应重新核验官方页面及小程序中的实时提示。

## 项目结构

```text
prepare-cn-online-court-materials/
├── SKILL.md                         # Skill 工作流与安全边界
├── agents/openai.yaml               # Codex 展示信息和默认提示词
├── assets/                          # 案件数据示例与渲染配置
├── references/                      # 证据、分流、文书和官方来源规则
└── scripts/build_case_packet.py     # DOCX/PDF 材料生成器
```

## 隐私与责任边界

请不要把未脱敏的身份证、住址、手机号、支付账号或完整案件材料提交到无关服务。本项目输出属于基于用户材料生成的自助草稿；是否受理、立案、认定案件性质及裁判结果，以有权机关依法审查为准。
