import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const outputDir = "/Users/zhazhakai777/Documents/博格华纳（宁波）项目开发管理/outputs/项目记录节点整理";
const outputPath = `${outputDir}/项目记录节点清单_开发交付版.xlsx`;

const nodes = [
  ["PR-001","项目","CKO","创建开发项目","开发项目已创建","记录项目编号、项目名称、客户、产品/系统、流程发起人及初始状态。","Sales","必须","关键","否","项目","否","项目记录起点。"],
  ["PR-002","项目","全阶段","编辑项目基本信息","项目信息已更新","逐字段记录修改前值、修改后值、修改原因；敏感字段变更应单独留痕。","PM/授权人员","必须","重要","视字段而定","项目","可选","不可只记录“保存成功”，必须能追溯字段差异。"],
  ["PR-003","阶段","CKO","完成 CKO 默认任务","CKO 任务已完成","记录完成的任务清单、完成人、完成时间、附件及备注。","Sales","必须","关键","是","项目任务","可选","系统默认任务内容变化由配置日志另行记录。"],
  ["PR-004","阶段","CKO→PKO","提交 CKO 并进入 PKO","项目完成 CKO 并进入 PKO","记录原阶段、新阶段、提交人、提交时间及准入检查结果。","Sales","必须","关键","是","项目/阶段实例","否","阶段流转记录。"],
  ["PR-005","阶段","PKO","发起/召开 PKO","PKO 项目启动会已发起/完成","记录会议时间、参会角色、纪要、结论和附件。","PM","必须","重要","否","项目/会议","是","Gate1-Gate6 团队成员参与。"],
  ["PR-006","计划","PKO","首次录入项目关键节点","项目计划关键节点已设置","记录客户 OTS/PPAP/SOP、BW 内部 OTS/PPAP/SOP 及对应 Gate 日期。","PM","必须","关键","是","项目计划","否","BW OTS=Gate4、PPAP=Gate5、SOP=Gate6。"],
  ["PR-007","计划","全阶段","修改项目关键节点","项目计划关键节点已调整","逐节点记录修改前日期、修改后日期、调整原因、是否触发缓冲不足提醒。","PM","必须","关键","视影响而定","项目计划","可选","时间节点属于审批依据时可影响 Gate。"],
  ["PR-008","阶段","PKO","选择标准/简化流程","项目流程分支已确定","记录产品类型、原流程、目标流程、适用 Gate 范围及判断说明。","PM","必须","关键","是","项目流程","否","标准 Gate1-6；简化 Gate1-3。"],
  ["PR-009","阶段","PKO及以后","变更流程分支","项目流程分支已变更","记录修改前后流程、原因、已生成任务处理方式及审批意见。","PM/审批人","必须","关键","是","项目流程","可选","高影响变更，应保留审批依据。"],
  ["PR-010","成员","PKO及以后","新增项目团队成员","项目团队成员已添加","记录成员、角色、部门、覆盖 Gate、加入时间及操作者。","PM","必须","重要","视角色而定","项目成员","否","原型已具备新增事件。"],
  ["PR-011","成员","全阶段","移除/替换团队成员","项目团队成员已移除/替换","记录原成员、新成员、角色、影响任务、交接说明及原因。","PM","必须","重要","视角色而定","项目成员","可选","不能只覆盖当前成员名单。"],
  ["PR-012","成员","全阶段","调整成员角色或 Gate 范围","项目成员职责已调整","记录调整前后角色、部门、Gate 范围和原因。","PM","必须","重要","视角色而定","项目成员","否","用于任务和审批追溯。"],
  ["PR-013","任务","PKO/Gate1-6","创建/批量生成项目级任务","项目阶段任务已创建","记录任务模板、K-PAC、所属 Gate、负责人、部门、计划完成日及创建来源。","PM/系统","必须","重要","是","项目任务","否","仅项目级任务进入项目 Gate 准入。"],
  ["PR-014","任务","全阶段","分配/改派项目级任务","项目任务负责人已调整","记录原负责人、新负责人、原因、计划日期变化及未完成工作交接。","PM","必须","重要","是","项目任务","否","保留每次改派。"],
  ["PR-015","任务","全阶段","提交项目级任务","项目任务已提交待审核","记录任务、提交人、填写内容、附件、提交时间及任务状态。","任务负责人","必须","关键","是","项目任务","是","原型已记录任务提交。"],
  ["PR-016","任务","全阶段","审核通过项目级任务","项目任务审核通过","记录审核人、结论、意见、完成时间和状态变化。","PM/审核人","必须","关键","是","项目任务","可选","任务闭环依据。"],
  ["PR-017","任务","全阶段","驳回项目级任务","项目任务审核驳回","记录审核人、驳回原因、整改要求、责任人和期限。","PM/审核人","必须","关键","是","项目任务","可选","驳回原因必须完整展示。"],
  ["PR-018","任务","全阶段","重新打开/取消项目级任务","项目任务已重开/取消","记录原状态、新状态、原因、审批人和关联整改事项。","PM/授权人员","必须","重要","视任务而定","项目任务","可选","不得物理删除历史任务。"],
  ["PR-019","BOM","PKO/Gate1","创建项目 BOM 容器/L1 总成","L1 总成 BOM 已创建","记录 L1 零件、工作版本、创建人、创建来源及初始状态。","AE","必须","关键","是","L1总成/BOM版本","否","项目容器本身不承载业务版本号。"],
  ["PR-020","BOM","全阶段","从生效版本创建新工作版本","L1 BOM 新工作版本已创建","记录来源生效版本、新版本号、所属阶段、复制时间和创建人。","AE","必须","关键","视影响而定","L1 BOM版本","否","每个 L1 独立升版。"],
  ["PR-021","BOM","全阶段","新增 BOM 行","BOM 行新增","记录零件、层级、父级、数量、APQP 标志、供应方案和起始 Gate。","AE","必须","重要","视影响分析","BOM行","可选","字段级变更记录。"],
  ["PR-022","BOM","全阶段","编辑 BOM 行","BOM 行字段已变更","逐字段记录修改前值、修改后值、操作者、原因和变更状态。","AE","必须","重要","视影响分析","BOM行","可选","数量/层级/图纸等变化均应留痕。"],
  ["PR-023","BOM","全阶段","删除 BOM 行","BOM 行删除待确认","记录被删零件、原父级、关联任务、删除原因及任务关闭/取消建议。","AE","必须","关键","L1或项目级影响时是","BOM行","可选","不得物理删除已生成任务。"],
  ["PR-024","BOM","全阶段","替换零件","BOM 零件替换待确认","记录旧零件、新零件、替换关系、任务继承/重做/关闭结论。","AE/SCM","必须","关键","视影响分析","BOM行/替换关系","可选","保留旧新零件追溯。"],
  ["PR-025","BOM","全阶段","新增/变更/删除供应方案","BOM 供应方案已变更","记录供应商、一供/二供、供应策略、是否开发/APQP、起始 Gate 及前后差异。","AE/SCM","必须","重要","仅项目级影响时是","供应方案","可选","默认只影响子零件开发，不阻塞项目 Gate。"],
  ["PR-026","BOM","全阶段","一供变更或供应商升降级","关键供应方案已调整","记录原供应商、新供应商、供应商类型变化、重评结论和风险。","SCM","必须","关键","仅项目级影响时是","供应方案","可选","高影响，触发供应商开发任务重评。"],
  ["PR-027","BOM","全阶段","AE 提交 L1 BOM 至 SCM","L1 BOM 已提交 SCM 确认","记录被选 L1、工作版本、行数、提交人、时间及备注。","AE","必须","关键","视当前Gate要求","L1 BOM版本","否","批量提交时逐个 L1 建立明细关联。"],
  ["PR-028","BOM","全阶段","SCM 驳回 L1 BOM","L1 BOM 确认被驳回","记录版本、驳回人、原因、整改要求和返回状态。","SCM","必须","关键","视当前Gate要求","L1 BOM版本","可选","保留每次往返。"],
  ["PR-029","BOM","全阶段","SCM 确认并锁定 L1 BOM","L1 BOM 已确认并生效","记录生效版本、原生效版本、确认人、锁定时间、控制码及变更影响数。","SCM","必须","关键","是","L1 BOM版本","否","同一 L1 同时只允许一个生效版本。"],
  ["PR-030","BOM","全阶段","冻结/解冻 BOM 版本","L1 BOM 版本已冻结/解冻","记录版本、原状态、新状态、原因、操作人和审批依据。","SCM/PM","必须","重要","视项目级影响","L1 BOM版本","可选","冻结不删除历史。"],
  ["PR-031","BOM","全阶段","旧版本作废/被新版本替代","L1 BOM 历史版本状态已更新","记录旧版本、新生效版本、作废时间及替代关系。","系统","必须","重要","否","L1 BOM版本","否","系统自动生成，不物理覆盖。"],
  ["PR-032","BOM","全阶段","完成 BOM 变更影响分析","BOM 变更影响分析已完成","记录变更类型、影响对象、项目级/子零件级判断、处理动作、结论人和结论时间。","PM/AE/SCM/AQE","必须","关键","按结论","BOM变更","可选","项目级与子零件级影响必须拆分。"],
  ["PR-033","BOM","全阶段","标记/解除项目级 BOM 阻塞","BOM 变更已标记/解除项目级阻塞","记录阻塞原因、影响 Gate、标记者、解除条件、解除人和关闭证据。","PM/审批人","必须","关键","是","BOM变更/项目Gate","可选","只有明确影响项目审批依据时允许阻塞。"],
  ["PR-034","APQP","Gate1及以后","创建 L1 总成 APQP","L1 总成 APQP 已创建","记录 L1、来源生效 BOM 版本、负责人（默认 AQE）、计划节点和创建人。","PM","必须","关键","是","L1 APQP","否","L1 APQP 是项目 Gate 准入对象。"],
  ["PR-035","APQP","全阶段","调整 L1 APQP 计划/负责人","L1 APQP 已调整","记录字段前后值、负责人变化、原因和影响 Gate。","PM/AQE","必须","关键","是","L1 APQP","可选","不可覆盖历史值。"],
  ["PR-036","APQP","全阶段","提交/审核/驳回/关闭/重开 L1 APQP","L1 APQP 状态已变更","记录原状态、新状态、提交或审核人、意见、附件及时间。","AQE/PM/审核人","必须","关键","是","L1 APQP","可选","按每次状态变化分别生成事件。"],
  ["PR-037","APQP","Gate1及以后","从 BOM 创建/关联子零件开发任务","已关联子零件开发任务","记录子零件任务号、零件、供应方案、来源 BOM 行、负责人和跳转关联。","AQE/SQD/系统","必须","提示","否","子零件开发任务","否","只作为跨模块关联，不并入项目 Gate 完成条件。"],
  ["PR-038","APQP","全阶段","子零件风险同步到项目","子零件开发风险已更新","记录风险类型、任务号、零件/供应商、摘要、同步时间和责任人。","系统/SQD","建议","提示","否","子零件开发任务/风险","可选","只能作为风险提醒，不得描述为 Gate 阻塞原因。"],
  ["PR-039","审批","Gate1-6","发起项目 Gate 审批","已发起项目 Gate 审批","记录 Gate、发起人、审核人/审核组、准入检查快照、提交时间。","PM","必须","关键","是","Gate审批实例","可选","项目级准入检查结果需固化。"],
  ["PR-040","审批","Gate1-6","审批通过","项目 Gate 审批通过","记录每位审核人意见、最终结论、通过时间及下一阶段。","审核人/审核组","必须","关键","是","Gate审批实例","可选","阶段推进依据。"],
  ["PR-041","审批","Gate1-6","带条件通过","项目 Gate 带条件通过","记录条件内容、责任人、完成期限、验证方式、审批人和下一阶段。","审核人/审核组","必须","关键","是","Gate审批/条件事项","可选","条件事项需后续闭环记录。"],
  ["PR-042","审批","Gate1-6","审批拒绝/退回整改","项目 Gate 审批未通过","记录拒绝成员、完整原因、整改项、责任人、期限和项目状态。","审核人/审核组","必须","关键","是","Gate审批实例","可选","原型已有审核记录弹窗，需落库。"],
  ["PR-043","审批","Gate1-6","条件事项/整改事项闭环","审批整改事项已闭环","记录事项、原要求、处理结果、验证人、证据和关闭时间。","责任人/审核人","必须","关键","是","条件/整改事项","是","闭环后才可满足项目级审批条件。"],
  ["PR-044","审批","Gate1-6","重新发起审批","项目 Gate 已重新提交审批","记录上一审批实例、整改摘要、本次审核人和提交时间。","PM","必须","关键","是","Gate审批实例","可选","审批实例不可覆盖。"],
  ["PR-045","阶段","Gate1-6","阶段推进","项目已进入下一 Gate","记录原 Gate、新 Gate、审批结论、推进人和生效时间。","系统/PM","必须","关键","是","项目/阶段实例","否","审批结论与阶段流转可关联但应分别记录。"],
  ["PR-046","审批","Gate4","发起 Gate4 标准审批/临时审批","Gate4 审批已发起","临批需额外记录到期日、临批原因、范围、附件和后续转正式要求。","PM","必须","关键","是","Gate4审批实例","是","原型已有标准/临批分支。"],
  ["PR-047","状态","全阶段","调整项目健康度","项目健康度已调整","记录原健康度、新健康度、原因、操作者和时间。","PM","必须","重要","否","项目健康度","否","原型已具备。"],
  ["PR-048","状态","全阶段","暂停项目","项目已暂停","记录原状态、新状态、暂停原因、附件、预计恢复条件和操作者。","PM/授权人员","必须","关键","是","项目状态","是","使用统一弹窗。"],
  ["PR-049","状态","全阶段","恢复项目","项目已恢复","记录暂停记录、恢复原因、恢复人、恢复时间及遗留风险。","PM/授权人员","必须","关键","否","项目状态","可选","与暂停记录关联。"],
  ["PR-050","状态","全阶段","终止项目","项目已终止","记录终止原因、批准人、附件、未完事项处置和终止时间。","PM/审批人","必须","关键","是","项目状态","是","危险操作需二次确认。"],
  ["PR-051","需求","全阶段","新增/编辑/删除客户要求","客户要求已新增/更新/删除","记录需求标题、内容前后值、发起人、附件、变更原因及操作者。","Sales/PM/授权人员","必须","重要","视项目级影响","客户要求","可选","删除采用逻辑删除并保留历史。"],
  ["PR-052","交接","Hand Over","登记/更新交接信息","项目交接信息已登记/更新","记录移交团队、接收团队、双方成员、交接内容、未关闭事项和附件。","PM/移交人","必须","关键","是","交接项","是","每个交接项独立留痕。"],
  ["PR-053","交接","Hand Over","交接项确认/退回","项目交接项已确认/退回","记录接收人、结论、退回原因、整改要求及时间。","接收人","必须","关键","是","交接项","可选","用于交接闭环。"],
  ["PR-054","阶段","Hand Over→结束","完成 Hand Over 并结束项目","项目交接完成，项目已闭环","记录交接完成情况、未关闭风险、完成确认人、完成时间和最终状态。","PM","必须","关键","是","项目/Hand Over","可选","项目结束节点。"],
  ["PR-055","系统","全阶段","数据导入/批量操作","项目数据已批量导入/更新","记录文件名、批次号、成功/失败数、操作者、错误明细和回滚标识。","授权人员/系统","必须","重要","视对象而定","导入批次","是","适用于 BOM、任务、成员等批量变更。"],
  ["PR-056","系统","全阶段","记录撤销/更正","项目记录已更正","保留原记录，新增更正记录并注明原记录ID、更正原因、审批人和更正后内容。","管理员/授权人员","必须","关键","否","项目记录","可选","项目记录原则上禁止物理删除或覆盖。"]
];

const fields = [
  ["recordId","记录唯一ID","字符串/UUID","是","系统生成","所有事件唯一，不可复用。"],
  ["projectId","项目ID","字符串/UUID","是","业务对象","必须关联项目。"],
  ["projectCode","项目编号快照","字符串","是","业务对象","防止后续名称变化影响查询。"],
  ["eventCode","事件编码","字符串","是","节点清单","例如 PR-029。"],
  ["category","一级分类","枚举","是","节点清单","阶段/计划/成员/任务/BOM/APQP/审批/状态/需求/交接/系统。"],
  ["eventType","动作类型","枚举","是","业务动作","CREATE/UPDATE/SUBMIT/APPROVE/REJECT/TRANSITION等。"],
  ["title","记录标题","字符串","是","标题模板","面向用户展示。"],
  ["description","记录内容","长文本","是","业务动作","关键字段不得用省略号截断。"],
  ["stageBefore","变更前阶段","枚举","否","项目阶段","阶段动作时填写。"],
  ["stageAfter","变更后阶段","枚举","否","项目阶段","阶段动作时填写。"],
  ["statusBefore","变更前状态","字符串/枚举","否","业务对象","状态类动作填写。"],
  ["statusAfter","变更后状态","字符串/枚举","否","业务对象","状态类动作填写。"],
  ["objectType","关联对象类型","枚举","是","业务对象","Project/Task/BomVersion/BomLine/Approval等。"],
  ["objectId","关联对象ID","字符串/UUID","是","业务对象","用于跳转详情。"],
  ["objectCode","关联对象编号快照","字符串","否","业务对象","如任务号、BOM版本号。"],
  ["parentRecordId","上一/关联记录ID","字符串/UUID","否","项目记录","审批重提、恢复、更正等场景使用。"],
  ["changeSet","字段差异集","JSON","否","系统计算","[{field,oldValue,newValue}]，编辑类事件必填。"],
  ["reason","原因/意见","长文本","按节点","用户输入","驳回、暂停、终止、更正等必须填写。"],
  ["gateImpact","是否影响项目Gate","布尔/枚举","是","业务判断","YES/NO/CONDITIONAL。"],
  ["gateImpactReason","项目Gate影响原因","长文本","条件必填","业务判断","仅项目级输入、L1 APQP、项目审批依据可作为阻塞原因。"],
  ["riskOnly","是否仅风险提醒","布尔","是","业务判断","子零件/供应商开发信息通常为 true。"],
  ["actorId","操作人ID","字符串/UUID","是","登录用户","系统事件使用系统账号。"],
  ["actorName","操作人姓名快照","字符串","是","登录用户","展示用快照。"],
  ["actorRole","操作角色快照","字符串","是","登录用户","如 PM/AE/SCM/AQE/SQD。"],
  ["occurredAt","业务发生时间","日期时间","是","系统时间","精确到秒，统一时区。"],
  ["createdAt","记录写入时间","日期时间","是","系统时间","与业务发生时间可不同。"],
  ["source","来源","枚举","是","系统","WEB/API/IMPORT/SYSTEM。"],
  ["requestId","请求/幂等ID","字符串","建议","接口层","避免重复点击产生重复记录。"],
  ["attachmentIds","附件ID集合","JSON/数组","否","文件服务","只存引用，不建议复制文件。"],
  ["approvalInstanceId","审批实例ID","字符串/UUID","否","审批模块","审批类事件必填。"],
  ["bomVersionId","L1 BOM版本ID","字符串/UUID","否","BOM模块","BOM/L1 APQP相关事件填写。"],
  ["bomLineId","BOM行ID","字符串/UUID","否","BOM模块","BOM行/子零件关联事件填写。"],
  ["supplyPlanId","供应方案ID","字符串/UUID","否","BOM模块","供应商方案事件填写。"],
  ["isVisible","前台是否展示","布尔","是","系统规则","审计保留但可控制前台展示。"],
  ["correctionOf","被更正记录ID","字符串/UUID","否","项目记录","更正时填写，不覆盖原记录。"],
  ["metadata","扩展信息","JSON","否","各模块","保存业务快照，不替代结构化字段。"]
];

const enums = [
  ["category","stage","阶段","阶段完成、阶段推进。"],
  ["category","plan","计划","关键节点计划设置和修改。"],
  ["category","team","成员","项目成员及职责变化。"],
  ["category","task","任务","仅项目级任务。"],
  ["category","bom","BOM","L1版本、BOM行、供应方案和影响分析。"],
  ["category","apqp","APQP","L1 APQP；子零件仅记录关联/风险。"],
  ["category","approval","审批","项目Gate审批和整改事项。"],
  ["category","status","状态","项目状态和健康度。"],
  ["category","requirement","需求","客户要求变化。"],
  ["category","handover","交接","Hand Over 信息和确认。"],
  ["category","system","系统","导入、更正等审计事件。"],
  ["eventType","CREATE","创建","新建对象或记录。"],
  ["eventType","UPDATE","更新","字段发生修改，需携带 changeSet。"],
  ["eventType","ASSIGN","分配/改派","任务或成员负责人变化。"],
  ["eventType","SUBMIT","提交","提交确认或审批。"],
  ["eventType","APPROVE","通过","任务/Gate/交接审批通过。"],
  ["eventType","CONDITIONAL_APPROVE","带条件通过","需创建并关联条件事项。"],
  ["eventType","REJECT","驳回/拒绝","必须记录完整原因。"],
  ["eventType","TRANSITION","阶段流转","记录前后阶段。"],
  ["eventType","PAUSE","暂停","项目暂停。"],
  ["eventType","RESUME","恢复","项目恢复。"],
  ["eventType","CANCEL","取消/终止","逻辑取消，不物理删除。"],
  ["eventType","CLOSE","关闭/完成","任务、事项或项目闭环。"],
  ["eventType","FREEZE","冻结","BOM或业务对象冻结。"],
  ["eventType","UNFREEZE","解冻","恢复编辑/流转。"],
  ["eventType","CORRECT","更正","关联原记录，不覆盖原记录。"],
  ["gateImpact","YES","阻塞/影响","必须是项目级条件未满足。"],
  ["gateImpact","NO","不影响","一般记录或跨模块信息。"],
  ["gateImpact","CONDITIONAL","需影响分析","结论完成后更新为 YES/NO。"]
];

const rules = [
  ["R-01","追加写入","项目记录采用 append-only；不允许物理删除或覆盖原记录，更正用新记录关联原记录。","必须"],
  ["R-02","字段差异","编辑类动作必须保存字段级 oldValue/newValue，不能只写“信息已修改”。","必须"],
  ["R-03","审批实例","每次发起、驳回、整改、重提都保留独立审批实例和审核人明细。","必须"],
  ["R-04","BOM版本","项目BOM容器无业务版本；记录必须关联具体 L1 BOM 版本。","必须"],
  ["R-05","批量动作","批量提交/确认需同时有批次记录，并为每个 L1 或对象保留明细关联。","必须"],
  ["R-06","项目Gate准入","只有项目任务、项目级必需输入、L1 APQP、项目审批事项及明确项目级 BOM 影响可阻塞 Gate。","必须"],
  ["R-07","子零件独立","子零件任务、子零件APQP、供应商开发状态不阻塞项目 Gate，只作为关联和风险提醒。","必须"],
  ["R-08","完整展示","项目记录、BOM、审批、任务等关键内容不得用 ... 或 text-overflow: ellipsis 截断。","必须"],
  ["R-09","统一弹窗","原因、确认、驳回、终止等输入使用统一项目弹窗，不使用 alert/confirm/prompt。","必须"],
  ["R-10","幂等","前端重复点击和接口重试不得生成重复记录，建议使用 requestId/idempotencyKey。","建议"],
  ["R-11","权限","记录查询按项目权限控制；后台审计记录不因前台隐藏而删除。","必须"],
  ["R-12","时间与快照","保存 actorId/objectId 的同时保存姓名、角色、编号等展示快照，时间统一到秒和时区。","必须"],
  ["R-13","UI筛选","当前原型有 plan、apqp 事件但筛选按钮未覆盖；开发时应增加“计划、APQP”，或明确归并到现有类别。","建议"],
  ["R-14","跳转","任务号、审批实例、BOM版本、子零件任务号等记录应支持跳转到有权限的业务详情。","建议"]
];

const wb = Workbook.create();
const summary = wb.worksheets.add("使用说明");
const nodeSheet = wb.worksheets.add("项目记录节点");
const fieldSheet = wb.worksheets.add("记录字段定义");
const enumSheet = wb.worksheets.add("枚举与开发规则");

const colors = {
  navy: "#17365D",
  blue: "#2F75B5",
  lightBlue: "#D9EAF7",
  pale: "#F3F7FA",
  gold: "#F4B183",
  green: "#E2F0D9",
  red: "#FCE4D6",
  gray: "#667085",
  line: "#D7DEE8",
  white: "#FFFFFF",
};

function title(sheet, range, text, subtitle) {
  sheet.getRange(range).merge();
  const cell = sheet.getRange(range.split(":")[0]);
  cell.values = [[text]];
  cell.format.fill = colors.navy;
  cell.format.font = { bold: true, color: colors.white, size: 18 };
  cell.format.rowHeight = 34;
  if (subtitle) {
    const row = Number(range.match(/\d+/)[0]) + 1;
    const startCol = range.match(/[A-Z]+/)[0];
    const endCol = range.split(":")[1].match(/[A-Z]+/)[0];
    sheet.getRange(`${startCol}${row}:${endCol}${row}`).merge();
    sheet.getRange(`${startCol}${row}`).values = [[subtitle]];
    sheet.getRange(`${startCol}${row}`).format.fill = colors.lightBlue;
    sheet.getRange(`${startCol}${row}`).format.font = { color: colors.navy, italic: true };
    sheet.getRange(`${startCol}${row}`).format.rowHeight = 28;
  }
}

function styleHeader(range) {
  range.format.fill = colors.blue;
  range.format.font = { bold: true, color: colors.white };
  range.format.verticalAlignment = "center";
  range.format.wrapText = true;
  range.format.rowHeight = 34;
  range.format.borders = { preset: "all", style: "thin", color: colors.line };
}

function styleBody(range) {
  range.format.verticalAlignment = "top";
  range.format.wrapText = true;
  range.format.borders = {
    insideHorizontal: { style: "thin", color: colors.line },
    bottom: { style: "thin", color: colors.line },
  };
}

summary.showGridLines = false;
title(summary, "A1:H1", "项目记录节点清单｜开发交付说明", "依据：项目开发详情6月26日第一版.html + 当前已明确业务规则");
summary.getRange("A4:B9").values = [
  ["交付目的","明确“什么动作需要生成项目记录”以及记录表/API至少保存哪些字段。"],
  ["记录范围","项目开发主流程：项目、阶段、计划、成员、项目任务、L1 BOM、L1 APQP、项目审批、状态、客户要求、交接。"],
  ["跨模块边界","子零件开发任务/APQP/供应商开发只记录关联或风险摘要，不作为项目 Gate 硬阻塞。"],
  ["落库原则","追加写入、不可覆盖；编辑保存字段差异；审批和BOM版本保留完整历史。"],
  ["前台展示","按时间倒序，支持分类筛选、对象跳转、完整内容查看和附件查看。"],
  ["开发验收","任一必记节点触发后，可查询到操作人、时间、对象、前后状态/字段差异、原因和关联附件。"]
];
summary.getRange("A4:A9").format.fill = colors.lightBlue;
summary.getRange("A4:A9").format.font = { bold: true, color: colors.navy };
styleBody(summary.getRange("A4:B9"));
summary.getRange("D4:E8").values = [
  ["指标","数量"],
  ["记录节点总数",nodes.length],
  ["必记节点",nodes.filter(r => r[7] === "必须").length],
  ["建议节点",nodes.filter(r => r[7] === "建议").length],
  ["字段总数",fields.length]
];
styleHeader(summary.getRange("D4:E4"));
styleBody(summary.getRange("D5:E8"));
summary.getRange("D5:D8").format.fill = colors.pale;
summary.getRange("D5:D8").format.font = { bold: true, color: colors.navy };
summary.getRange("A12:H12").merge();
summary.getRange("A12").values = [["开发特别注意"]];
summary.getRange("A12").format.fill = colors.gold;
summary.getRange("A12").format.font = { bold: true, color: "#7F4125" };
summary.getRange("A13:H17").merge();
summary.getRange("A13").values = [[
  "1）当前原型项目记录已有阶段、计划、成员、任务、BOM、APQP、审批、状态事件，但筛选区缺少“计划、APQP”；\n" +
  "2）“项目 Gate 被阻塞”的原因只能来自项目级任务、项目级必需输入、L1 总成 APQP、项目审批事项或明确影响项目审批依据的 L1 BOM 变更；\n" +
  "3）子零件任务未启动、延期、未完成、供应商风险等只允许显示为风险提醒，不得写成项目 Gate 阻塞原因；\n" +
  "4）任何项目记录更正都要新增一条更正记录，不能修改或删除原记录。"
]];
summary.getRange("A13").format.wrapText = true;
summary.getRange("A13").format.verticalAlignment = "top";
summary.getRange("A13:H17").format.fill = colors.red;
summary.getRange("A13:H17").format.borders = { preset: "outside", style: "thin", color: colors.gold };
summary.getRange("A:A").format.columnWidth = 18;
summary.getRange("B:B").format.columnWidth = 58;
summary.getRange("C:C").format.columnWidth = 3;
summary.getRange("D:D").format.columnWidth = 20;
summary.getRange("E:E").format.columnWidth = 12;
summary.getRange("F:H").format.columnWidth = 14;
summary.freezePanes.freezeRows(2);

nodeSheet.showGridLines = false;
title(nodeSheet, "A1:M1", "项目记录节点清单", "“是否影响项目Gate”表示该业务对象本身是否可能进入项目级准入判断；子零件相关节点固定为否。");
const nodeHeaders = [["编号","一级分类","业务阶段","触发动作/节点","记录标题模板","必须记录的内容","操作角色","必记级别","重要度","是否影响项目Gate","关联对象","附件","开发备注"]];
nodeSheet.getRange("A4:M4").values = nodeHeaders;
nodeSheet.getRange(`A5:M${nodes.length + 4}`).values = nodes;
styleHeader(nodeSheet.getRange("A4:M4"));
styleBody(nodeSheet.getRange(`A5:M${nodes.length + 4}`));
nodeSheet.getRange(`A5:A${nodes.length + 4}`).format.font = { bold: true, color: colors.navy };
nodeSheet.getRange(`B5:B${nodes.length + 4}`).format.fill = colors.pale;
nodeSheet.getRange(`H5:J${nodes.length + 4}`).format.horizontalAlignment = "center";
nodeSheet.getRange(`H5:H${nodes.length + 4}`).conditionalFormats.add("containsText", { text: "必须", format: { fill: colors.green, font: { bold: true, color: "#375623" } } });
nodeSheet.getRange(`J5:J${nodes.length + 4}`).conditionalFormats.add("containsText", { text: "否", format: { fill: colors.lightBlue, font: { color: colors.navy } } });
nodeSheet.getRange(`J5:J${nodes.length + 4}`).conditionalFormats.add("containsText", { text: "是", format: { fill: colors.red, font: { bold: true, color: "#9C0006" } } });
nodeSheet.getRange(`A4:M${nodes.length + 4}`).format.rowHeight = 42;
nodeSheet.getRange("A4:M4").format.rowHeight = 40;
const nodeWidths = [12,12,15,27,28,58,18,12,10,18,22,10,48];
nodeWidths.forEach((w, i) => nodeSheet.getRangeByIndexes(0, i, nodes.length + 4, 1).format.columnWidth = w);
nodeSheet.freezePanes.freezeRows(4);
nodeSheet.freezePanes.freezeColumns(3);
nodeSheet.tables.add(`A4:M${nodes.length + 4}`, true, "ProjectRecordNodes").style = "TableStyleMedium2";

fieldSheet.showGridLines = false;
title(fieldSheet, "A1:F1", "项目记录字段定义", "建议使用主记录表 + 审批人明细/附件关联/字段差异 JSON；字段是否必填以具体事件节点为准。");
fieldSheet.getRange("A4:F4").values = [["字段名","中文名称","建议类型","基础必填","数据来源","说明/校验规则"]];
fieldSheet.getRange(`A5:F${fields.length + 4}`).values = fields;
styleHeader(fieldSheet.getRange("A4:F4"));
styleBody(fieldSheet.getRange(`A5:F${fields.length + 4}`));
fieldSheet.getRange(`A5:A${fields.length + 4}`).format.font = { bold: true, color: colors.navy };
fieldSheet.getRange(`D5:D${fields.length + 4}`).format.horizontalAlignment = "center";
[22,24,20,12,22,65].forEach((w, i) => fieldSheet.getRangeByIndexes(0, i, fields.length + 4, 1).format.columnWidth = w);
fieldSheet.getRange(`A4:F${fields.length + 4}`).format.rowHeight = 34;
fieldSheet.freezePanes.freezeRows(4);
fieldSheet.freezePanes.freezeColumns(2);
fieldSheet.tables.add(`A4:F${fields.length + 4}`, true, "ProjectRecordFields").style = "TableStyleMedium2";

enumSheet.showGridLines = false;
title(enumSheet, "A1:H1", "枚举与开发规则", "枚举值用于前后端统一；规则用于开发和测试验收。");
enumSheet.getRange("A4:D4").values = [["字段","代码","中文显示","使用说明"]];
enumSheet.getRange(`A5:D${enums.length + 4}`).values = enums;
styleHeader(enumSheet.getRange("A4:D4"));
styleBody(enumSheet.getRange(`A5:D${enums.length + 4}`));
enumSheet.getRange(`B5:B${enums.length + 4}`).format.font = { bold: true, color: colors.navy };
enumSheet.getRange("F4:H4").values = [["规则编号","规则主题","开发/验收要求"]];
enumSheet.getRange(`F5:H${rules.length + 4}`).values = rules.map(r => [r[0], `${r[1]}（${r[3]}）`, r[2]]);
styleHeader(enumSheet.getRange("F4:H4"));
styleBody(enumSheet.getRange(`F5:H${rules.length + 4}`));
[18,24,26,55,4,14,25,76].forEach((w, i) => enumSheet.getRangeByIndexes(0, i, Math.max(enums.length, rules.length) + 4, 1).format.columnWidth = w);
enumSheet.getRange(`A4:H${Math.max(enums.length, rules.length) + 4}`).format.rowHeight = 35;
enumSheet.freezePanes.freezeRows(4);
enumSheet.tables.add(`A4:D${enums.length + 4}`, true, "RecordEnums").style = "TableStyleMedium2";
enumSheet.tables.add(`F4:H${rules.length + 4}`, true, "DevelopmentRules").style = "TableStyleMedium4";

await fs.mkdir(outputDir, { recursive: true });
const checks = [
  await wb.inspect({ kind: "table", range: "项目记录节点!A1:M12", include: "values,formulas", tableMaxRows: 12, tableMaxCols: 13 }),
  await wb.inspect({ kind: "table", range: "记录字段定义!A1:F12", include: "values,formulas", tableMaxRows: 12, tableMaxCols: 6 }),
  await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula errors" })
];
for (const c of checks) console.log(c.ndjson);

for (const [sheetName, range, filename] of [
  ["使用说明","A1:H17","preview_使用说明.png"],
  ["项目记录节点","A1:M18","preview_项目记录节点.png"],
  ["记录字段定义","A1:F18","preview_字段定义.png"],
  ["枚举与开发规则","A1:H20","preview_枚举规则.png"],
]) {
  const rendered = await wb.render({ sheetName, range, scale: 1.2, format: "png" });
  await fs.writeFile(`${outputDir}/${filename}`, new Uint8Array(await rendered.arrayBuffer()));
}

const file = await SpreadsheetFile.exportXlsx(wb);
await file.save(outputPath);
console.log(`OUTPUT=${outputPath}`);
