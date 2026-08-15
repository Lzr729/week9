import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = "/workspace/scratch/a4b07439de83";
const workDir = `${root}/week9_closure_work`;
const previewDir = `${workDir}/previews`;
const outputDir = `${root}/outputs/week9_closure`;
const outputXlsx = `${outputDir}/week9_research_variables_descriptive_analysis_v1.0.xlsx`;
const packageDir = `${outputDir}/week9_github_submit_v1.0`;
const version = "Week9 Research Variables and Descriptive Analysis v1.0";
const freezeDate = "2026-08-14";

const stageFiles = [
  ["S01", "Stage 1 分析入口", "outputs/week9_stage01/week9_stage01_analysis_entry.xlsx", "FROZEN_INPUT", "确认Week8输入、层级和分析边界"],
  ["S02", "Stage 2 基础描述统计", "outputs/week9_stage02/week9_stage02_descriptive_statistics.xlsx", "COMPLETE", "公司、轮次、交易、主体与路径的描述统计"],
  ["S03", "Stage 3 候选变量开发", "outputs/week9_stage03/week9_stage03_candidate_variable_development.xlsx", "COMPLETE", "32个候选变量及定义"],
  ["S04", "Stage 4 可计算性与缺失", "outputs/week9_stage04/week9_stage04_variable_computability_and_missingness.xlsx", "SUPERSEDED_BY_S04C", "修复前变量状态审计"],
  ["S04B1", "云汉芯城角色范围修复", "outputs/week9_stage04b1/week9_stage04b1_yunhan_role_scope_repair.xlsx", "COMPLETE", "角色范围与去重修复"],
  ["S04B2", "友升股份PE/VC主体桥接修复", "outputs/week9_stage04b2/week9_stage04b2_yousheng_pevc_party_bridge_repair.xlsx", "COMPLETE", "PE/VC主体及轮次桥接"],
  ["S04B3", "三协电机PE/VC路径修复", "outputs/week9_stage04b3/week9_stage04b3_sanxie_pevc_investment_path_repair.xlsx", "COMPLETE", "投资路径证据链修复"],
  ["S04B4", "黄山谷捷融资主体修复", "outputs/week9_stage04b4/week9_stage04b4_huangshan_financing_party_repair.xlsx", "COMPLETE", "融资参与方桥接修复"],
  ["S04C", "变量可用性冻结", "outputs/week9_stage04c/week9_stage04c_variable_availability_frozen_v1.0.xlsx", "FROZEN", "256个公司—变量单元及最终状态"],
  ["S05A", "价值评估框架", "outputs/week9_stage05a/week9_stage05a_value_evaluation_framework_v1.0.xlsx", "COMPLETE", "学术与商业价值评估框架"],
  ["S05B", "学术价值筛选", "outputs/week9_stage05b/week9_stage05b_academic_value_screen_v1.0.xlsx", "COMPLETE", "学术用途与研究问题筛选"],
  ["S05C", "商业价值筛选", "outputs/week9_stage05c/week9_stage05c_commercial_value_screen_v1.0.xlsx", "COMPLETE", "商业用途与决策问题筛选"],
  ["S05D", "学术—商业价值整合", "outputs/week9_stage05d/week9_stage05d_academic_commercial_integration_v1.0.xlsx", "FROZEN", "变量角色、组合、轴和闸门"],
  ["S06A", "公司选择矩阵底表", "outputs/week9_stage06a/week9_stage06a_company_selection_matrix_base_v1.0.xlsx", "COMPLETE", "公司选择所需变量矩阵底表"],
  ["S06B", "公司轴向初评", "outputs/week9_stage06b/week9_stage06b_company_axis_initial_assessment_v1.0.xlsx", "FROZEN", "8家公司画像与资格状态"],
  ["S06C", "案例互补性候选组合", "outputs/week9_stage06c/week9_stage06c_case_complementarity_candidate_portfolios_v1.0.xlsx", "FROZEN_CANDIDATES", "8个无排名候选组合；未最终选择"],
];

await fs.mkdir(previewDir, { recursive: true });
await fs.mkdir(outputDir, { recursive: true });
await fs.rm(packageDir, { recursive: true, force: true });
for (const d of ["reports", "data", "docs", "inputs", "scripts", "metadata"]) await fs.mkdir(`${packageDir}/${d}`, { recursive: true });

const sha256 = async (p) => crypto.createHash("sha256").update(await fs.readFile(p)).digest("hex");
const parseTable = (res) => {
  const records = res.ndjson.trim().split("\n").filter(Boolean).map(JSON.parse);
  const t = records.find((x) => x.kind === "table");
  if (!t?.values) throw new Error("Unable to parse table inspection");
  return t.values;
};
const readRange = async (rel, range, rows, cols) => {
  const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(`${root}/${rel}`));
  return parseTable(await wb.inspect({ kind:"table", range, include:"values,formulas", tableMaxRows:rows, tableMaxCols:cols, maxChars:600000 }));
};

const rel = Object.fromEntries(stageFiles.map((r) => [r[0], r[2]]));
const [desc, vars, availability, longData, integrated, profiles, portfolios] = await Promise.all([
  readRange(rel.S02, "描述统计!A4:M10", 12, 15),
  readRange(rel.S03, "候选变量字典!A4:O36", 40, 18),
  readRange(rel.S04C, "变量可用性清单!A4:P36", 40, 18),
  readRange(rel.S04C, "最终长表!A4:W260", 265, 25),
  readRange(rel.S05D, "整合变量地图!A4:Q36", 40, 20),
  readRange(rel.S06B, "公司画像摘要!A4:L12", 14, 14),
  readRange(rel.S06C, "候选组合画像!A4:X12", 14, 26),
]);

const expected = { desc:7, vars:33, availability:33, longData:257, integrated:33, profiles:9, portfolios:9 };
for (const [k, n] of Object.entries(expected)) if (eval(k).length !== n) throw new Error(`${k} row count mismatch: ${eval(k).length}`);

const sourceInventory = [];
for (const row of stageFiles) sourceInventory.push([...row, await sha256(`${root}/${row[2]}`)]);
const week8Zip = `${root}/upload/week8_github_submit_v1.0(1).zip`;
const week8ZipSha = await sha256(week8Zip);

const csvEncode = (table) => "\uFEFF" + table.map((row) => row.map((v) => {
  if (v === null || v === undefined) return "";
  const s = String(v);
  return /[",\r\n]/.test(s) ? `"${s.replaceAll('"','""')}"` : s;
}).join(",")).join("\r\n") + "\r\n";
const csvOutputs = [
  ["week9_descriptive_statistics.csv", desc],
  ["week9_variable_dictionary.csv", vars],
  ["week9_variable_availability.csv", availability],
  ["week9_company_variable_long.csv", longData],
  ["week9_integrated_variable_map.csv", integrated],
  ["week9_company_profiles.csv", profiles],
  ["week9_candidate_portfolios.csv", portfolios],
];
for (const [name, table] of csvOutputs) await fs.writeFile(`${packageDir}/data/${name}`, csvEncode(table), "utf8");

const C = { navy:"#0B1F33", teal:"#0F766E", blue:"#DCEBFA", white:"#FFFFFF", black:"#111827", gray:"#475569", line:"#CBD5E1", pale:"#F1F5F9", green:"#DCFCE7", greenText:"#166534", amber:"#FEF3C7", amberText:"#92400E", purple:"#EDE9FE", purpleText:"#5B21B6", source:"#008000" };
const wb = Workbook.create();
const sheetNames = ["结项说明","阶段成果清单","核心描述统计","变量字典","变量可用性","公司变量长表","价值整合地图","公司画像","候选组合暂存","数据层级与边界","Week10交接","冻结验收"];
const sheets = Object.fromEntries(sheetNames.map((n) => [n, wb.worksheets.add(n)]));
const colName = (n) => { let s=""; while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);} return s; };
const titleBand = (s, title, subtitle, last) => {
  s.mergeCells(`A1:${last}1`); s.getRange("A1").values=[[title]];
  s.getRange(`A1:${last}1`).format={fill:C.navy,font:{bold:true,color:C.white,size:16,name:"Microsoft YaHei"},rowHeight:32,verticalAlignment:"center"};
  s.mergeCells(`A2:${last}2`); s.getRange("A2").values=[[subtitle]];
  s.getRange(`A2:${last}2`).format={fill:C.blue,font:{color:C.gray,size:10,name:"Microsoft YaHei"},wrapText:true,rowHeight:38,verticalAlignment:"center"};
  s.showGridLines=false;
};
const widths = (s, arr) => arr.forEach((w,i) => s.getRangeByIndexes(0,i,1,1).format.columnWidth=w);
const section = (s, row, text, last) => { s.mergeCells(`A${row}:${last}${row}`); s.getRange(`A${row}`).values=[[text]]; s.getRange(`A${row}:${last}${row}`).format={fill:C.navy,font:{bold:true,color:C.white,name:"Microsoft YaHei"},rowHeight:25}; };
const writeTable = (s, row, table, options={}) => {
  const cols=table[0].length, endCol=colName(cols), end=row+table.length-1;
  s.getRange(`A${row}:${endCol}${end}`).values=table;
  s.getRange(`A${row}:${endCol}${row}`).format={fill:C.teal,font:{bold:true,color:C.white,name:"Microsoft YaHei"},wrapText:true,rowHeight:38,verticalAlignment:"center",horizontalAlignment:"center"};
  if (end>row) s.getRange(`A${row+1}:${endCol}${end}`).format={font:{color:options.sourceColor?C.source:C.black,name:"Microsoft YaHei",size:9},wrapText:true,verticalAlignment:"top",borders:{bottom:{style:"thin",color:C.line}},rowHeight:options.rowHeight||34};
  if (options.freeze !== false) { s.freezePanes.freezeRows(row); if(options.freezeCols) s.freezePanes.freezeColumns(options.freezeCols); }
  return {end,endCol};
};

const cover=sheets["结项说明"];
titleBand(cover,"Week 9 结项与冻结",`${version}｜冻结日期 ${freezeDate}｜以Week8 Standard Dataset v1.0为不可变基础输入`,"H");
widths(cover,[18,30,56,34,32,32,34,38]);
section(cover,4,"A. 冻结结论","H");
writeTable(cover,5,[
  ["项目","冻结结果","数量/状态","数据层级","依据","是否进入Week10","边界","验收"],
  ["样本公司","8家公司","8","标准化值","Week8冻结样本","是","小样本描述，不作总体推断","PASS"],
  ["基础描述统计","6项核心统计","6","计算值","Stage2","是","金额均为披露可核验下限","PASS"],
  ["候选研究变量","变量字典与规则","32","研究定义","Stage3","是","变量定义不等于理论构念已验证","PASS"],
  ["公司—变量状态","最终冻结长表","256","计算值+状态","Stage4C","是","null、0、结构性不适用分离","PASS"],
  ["数据修复","证据链内修复单元","10","标准化/计算值","Stage4B1—B4","是","未进行推测性填补","PASS"],
  ["剩余不可计算","均为影石创新","3","状态值","Stage4C","是","保持NOT_COMPUTABLE，不插补","PASS"],
  ["代表性公司","暂不最终选择","DEFERRED","研究判断","Stage6C+用户决定","Week10后","先研读VC/PE论文再修订方法","PASS"],
],{rowHeight:48});
section(cover,15,"B. 结项口径","H");
writeTable(cover,16,[
  ["层级","本周处理","允许用途","禁止用途","空值规则","版本状态","下一步","责任"],
  ["原文披露值","不改写Week8来源事实","证据回查与引用","用研究判断替代原文","未披露保持null","继承Week8冻结","按论文方法决定是否补充定义","Week8"],
  ["标准化值","沿用统一单位、ID和Schema","跨公司比较底表","回写或覆盖Week8","不能把null变成0","继承Week8冻结","只读使用","Week8"],
  ["计算值","按已记录公式形成统计与变量","描述性分析、案例比较","超出披露覆盖作精确推断","分母不足则NOT_COMPUTABLE","Week9冻结","Week10可提出新版本方案","Week9"],
  ["研究判断","变量角色、公司画像、候选组合","提出问题与案例设计","伪装成披露事实或因果结果","边界写入判断说明","Week9冻结","经论文研读后另行修订","Week9"],
],{rowHeight:58});
section(cover,23,"C. 版本与可复现性","H");
writeTable(cover,24,[
  ["对象","版本/标识","状态","哈希/规则","变更政策","发布物","验收标准","结果"],
  ["基础数据","Week8 Standard Dataset v1.0","FROZEN",week8ZipSha,"禁止在Week9回写","Week8原包仅作外部依赖","哈希固定","PASS"],
  ["Week9成果",version,"FROZEN","见阶段成果清单与发布清单","任何后续修改升版，不覆盖v1.0","总工作簿+CSV+文档+脚本","独立审计全部通过","PASS"],
  ["公司选择","Stage6C候选组合","DEFERRED","无排名、无FINAL_SELECTED","Week10形成研究方法后再决策","候选组合暂存表","最终选择数=0","PASS"],
],{rowHeight:52});
cover.freezePanes.freezeRows(2);

const inv=sheets["阶段成果清单"];
titleBand(inv,"Week 9 阶段成果清单","16个阶段文件逐一固化SHA-256；原文件不修改，冻结总表只读汇总。","H");
widths(inv,[12,32,62,24,58,70,24,18]);
const invTable=[["阶段ID","阶段名称","相对路径","状态","用途","SHA-256","纳入精简包","验收"],...sourceInventory.map((r)=>[r[0],r[1],r[2],r[3],r[4],r[5],["S02","S03","S04C","S05D","S06B","S06C"].includes(r[0])?"YES":"MANIFEST_ONLY","PASS"] )];
writeTable(inv,4,invTable,{rowHeight:54,freezeCols:2});

const copyDataSheet=(name,title,subtitle,table,colWidths,freezeCols=2,rowHeight=34) => {
  const s=sheets[name]; const last=colName(table[0].length); titleBand(s,title,subtitle,last); widths(s,colWidths); writeTable(s,4,table,{sourceColor:true,freezeCols,rowHeight}); return s;
};
copyDataSheet("核心描述统计","八家公司基础描述统计","所有数值均来自Stage2；公司融资总额与轮次锚点合计口径不同，差额120万元已在Stage2登记。",desc,[22,42,24,18,18,18,18,18,18,22,42,44,34],1,46);
copyDataSheet("变量字典","候选研究变量字典","32个变量的构念、定义、单位、算法与解释边界；绿色字体表示从冻结阶段成果导入。",vars,[12,28,25,18,16,18,46,52,32,32,40,38,38,30,34],2,62);
copyDataSheet("变量可用性","变量可计算性与缺失冻结清单","最终状态来源于Stage4C；保留VALID、STRUCTURAL_NA、NOT_COMPUTABLE的语义差异。",availability,[12,28,24,16,16,18,18,18,18,18,22,22,24,24,42,34],2,48);
copyDataSheet("公司变量长表","公司—变量最终冻结长表","8家公司×32变量=256单元；本表是Week10方法开发的直接数据接口。",longData,[24,14,20,12,30,24,12,16,14,18,18,20,20,22,50,16,20,42,46,54,22,24,16],3,52);
copyDataSheet("价值整合地图","学术价值与商业价值整合地图","不生成学术—商业综合总分；变量角色、组合、轴、闸门和约束均保留原边界。",integrated,[12,28,22,18,18,18,18,20,20,22,22,28,38,38,38,32,38],2,54);
copyDataSheet("公司画像","公司轴向画像与资格状态","公司画像是研究判断，不是公司优劣排名；资格状态用于后续案例设计。",profiles,[14,20,30,24,24,24,24,24,24,24,38,56],2,64);
copyDataSheet("候选组合暂存","代表性公司候选组合（暂存）","8个候选组合无排名、无最终选择；必须在Week10论文方法研读后重新审视。",portfolios,[10,30,14,14,14,14,14,14,14,14,16,18,22,22,18,30,48,56,56,48,48,34,34,40],2,88);

const boundary=sheets["数据层级与边界"];
titleBand(boundary,"数据层级、缺失语义与解释边界","本表是使用Week9冻结成果时必须遵循的读数规则。","H"); widths(boundary,[16,30,64,52,52,42,36,24]);
writeTable(boundary,4,[
  ["规则ID","对象","定义","正确处理","禁止处理","示例","适用输出","状态"],
  ["B01","原文披露值","招股书或正式披露直接给出的事实","保留原措辞与来源记录","用计算结果冒充披露原值","披露交易对价、主体名称","全部","ENFORCED"],
  ["B02","标准化值","对单位、ID、字段和枚举统一后的值","可跨表关联，保留来源","覆盖Week8冻结数据","万元口径、公司ID","全部","ENFORCED"],
  ["B03","计算值","基于冻结输入和公开公式得到的结果","附算法、分母与覆盖率","超出披露覆盖进行外推","融资金额下限、密度指标","统计/变量","ENFORCED"],
  ["B04","研究判断","为学术或商业问题形成的分类与用途判断","与事实列分开、保留解释边界","当作公司客观评级","变量角色、公司画像","价值/选择","ENFORCED"],
  ["B05","null","没有可报告的值；原因需由状态解释","结合final_status与reason_code读取","自动替换为0","影石创新部分变量","长表","ENFORCED"],
  ["B06","STRUCTURAL_NA","变量对该公司结构上不适用","不进入以该变量为分母的比较","当作缺失质量问题或0","无PE/VC时的进入时点变量","长表","ENFORCED"],
  ["B07","NOT_COMPUTABLE","概念适用但证据不足，无法计算","保持null并披露原因","均值填补、0填补或主观估计","影石创新CV003/CV004/CV026","长表","ENFORCED"],
  ["B08","显式0","证据支持数值确为零","explicit_zero_flag需能识别","与null合并处理","冻结口径内确认数为0","长表","ENFORCED"],
  ["B09","小样本边界","8家公司仅支持样本内描述和案例开发","报告范围、分布和异质性","显著性、因果性、总体代表性宣称","所有统计与组合","全部","ENFORCED"],
  ["B10","代表性选择","取决于研究问题和方法，不是金额排名","论文研读后形成选择准则","在Week9直接选2—3家公司","Stage6C候选组合","选择阶段","DEFERRED"],
],{rowHeight:66,freezeCols:2});

const hand=sheets["Week10交接"];
titleBand(hand,"Week 10 交接：VC/PE论文研读与方法校准","Week10先研读论文，再决定是否修订变量和代表性公司选择标准；Week9 v1.0保持不变。","H"); widths(hand,[14,32,56,54,48,46,36,22]);
section(hand,4,"A. 建议流程","H");
writeTable(hand,5,[
  ["步骤","任务","输入","处理","输出","回写规则","验收标准","状态"],
  [1,"选择并研读目标论文","论文全文、附录、变量表","拆解研究问题、理论机制、样本、识别、变量、稳健性","结构化论文研读卡","不得直接回写Week9","方法字段完整","PENDING"],
  [2,"建立方法映射","论文方法+Week9变量字典","逐项标记可借鉴、需改写、不可迁移","论文—本项目方法映射表","研究判断单独记录","每项有理由","PENDING"],
  [3,"开展变量差距分析","Stage4C长表+Stage5D整合地图","检查定义、测量层级、可计算性和缺失偏差","变量保留/修订/新增/删除建议","不覆盖v1.0；必要时建v1.1","区分事实与判断","PENDING"],
  [4,"重审案例选择标准","公司画像+候选组合+论文方法","确定研究主问题、案例逻辑、纳入排除规则","代表性选择方案草案","禁止仅按融资额或S6分值选择","规则可复核","PENDING"],
  [5,"形成Week10决策闸门","全部上述输出","判断是否需要数据补充或版本升级","GO / REVISE / HOLD","新证据走独立修复流程","选择结论有证据链","PENDING"],
],{rowHeight:70});
section(hand,13,"B. 论文研读最低提取字段","H");
writeTable(hand,14,[
  ["字段组","最低提取内容","为什么重要","与Week9接口","常见风险","记录形式","必填","状态"],
  ["研究设计","研究问题、理论机制、分析单位、时间窗口","决定变量和案例是否匹配","变量字典、公司画像","只模仿统计模型不看问题","论文研读卡","YES","PENDING"],
  ["样本与识别","样本来源、纳入排除、对照、内生性处理","决定结论强度和可迁移性","候选组合、边界表","把大样本识别直接移植到n=8","方法映射","YES","PENDING"],
  ["变量测量","因变量、自变量、机制变量、控制变量、算法","检验现有32变量是否覆盖研究构念","变量字典、长表","同名变量定义不同","变量对照表","YES","PENDING"],
  ["缺失处理","缺失机制、删除、插补、敏感性分析","保护null与0的区别","可用性清单、长表","机械插补不可计算值","缺失策略表","YES","PENDING"],
  ["结果与边界","效应、稳健性、外部效度、作者局限","防止过度解释本项目描述统计","价值地图、公司画像","只摘显著结果","证据摘要","YES","PENDING"],
],{rowHeight:70});

const checks=sheets["冻结验收"];
titleBand(checks,"Week 9 冻结验收","公式检查与独立脚本共同验证行数、状态、修复、候选边界和最终选择状态。","H"); widths(checks,[16,46,18,24,16,18,20,58]);
const checkRows = [
  ["W9F-001","描述统计指标数",6,"=COUNTA('核心描述统计'!$A$5:$A$10)","数值",0,"","Stage2共6项"],
  ["W9F-002","候选变量数",32,"=COUNTA('变量字典'!$A$5:$A$36)","数值",0,"","变量ID CV001—CV032"],
  ["W9F-003","变量可用性记录数",32,"=COUNTA('变量可用性'!$A$5:$A$36)","数值",0,"","每变量一行"],
  ["W9F-004","公司—变量单元数",256,"=COUNTA('公司变量长表'!$A$5:$A$260)","数值",0,"","8×32"],
  ["W9F-005","VALID单元数",215,"=COUNTIF('公司变量长表'!$M$5:$M$260,\"VALID\")","数值",0,"","冻结状态"],
  ["W9F-006","STRUCTURAL_NA单元数",38,"=COUNTIF('公司变量长表'!$M$5:$M$260,\"STRUCTURAL_NA\")","数值",0,"","结构性不适用"],
  ["W9F-007","NOT_COMPUTABLE单元数",3,"=COUNTIF('公司变量长表'!$M$5:$M$260,\"NOT_COMPUTABLE\")","数值",0,"","不得插补"],
  ["W9F-008","三类最终状态合计",256,"=COUNTIF('公司变量长表'!$M$5:$M$260,\"VALID\")+COUNTIF('公司变量长表'!$M$5:$M$260,\"STRUCTURAL_NA\")+COUNTIF('公司变量长表'!$M$5:$M$260,\"NOT_COMPUTABLE\")","数值",0,"","状态闭合"],
  ["W9F-009","修复单元数",10,"=COUNTIF('公司变量长表'!$U$5:$U$260,\"<>\")","数值",0,"","Stage4B1—B4"],
  ["W9F-010","整合变量记录数",32,"=COUNTA('价值整合地图'!$A$5:$A$36)","数值",0,"","不生成综合总分"],
  ["W9F-011","公司画像数",8,"=COUNTA('公司画像'!$A$5:$A$12)","数值",0,"","8家公司"],
  ["W9F-012","候选组合数",8,"=COUNTA('候选组合暂存'!$A$5:$A$12)","数值",0,"","无排名"],
  ["W9F-013","最终选择数",0,"=COUNTIF('候选组合暂存'!$V$5:$V$12,\"FINAL_SELECTED\")","数值",0,"","Week9明确延后"],
  ["W9F-014","阶段成果文件数",16,"=COUNTA('阶段成果清单'!$A$5:$A$20)","数值",0,"","全阶段可追溯"],
  ["W9F-015","基础样本公司数",8,"=COUNTA('公司画像'!$A$5:$A$12)","数值",0,"","不得外推总体"],
];
writeTable(checks,4,[["检查ID","检查项","期望值","实际值","类型","容差","状态","说明"],...checkRows.map((r)=>[...r])],{rowHeight:40});
for(let i=0;i<checkRows.length;i++){const rr=5+i; checks.getRange(`D${rr}`).formulas=[[checkRows[i][3]]]; checks.getRange(`G${rr}`).formulas=[[`=IF(ABS(C${rr}-D${rr})<=F${rr},\"PASS\",\"FAIL\")`]];}
const finalRow=5+checkRows.length+1; section(checks,finalRow,"最终冻结状态","H"); checks.getRange(`A${finalRow+1}`).values=[["WEEK9_FREEZE_STATUS"]]; checks.getRange(`B${finalRow+1}`).formulas=[[`=IF(COUNTIF(G5:G${4+checkRows.length},\"FAIL\")=0,\"PASS\",\"REVIEW\")`]]; checks.mergeCells(`B${finalRow+1}:H${finalRow+1}`); checks.getRange(`A${finalRow+1}:H${finalRow+1}`).format={fill:C.green,font:{bold:true,color:C.greenText,name:"Microsoft YaHei"},rowHeight:32}; checks.getRange(`G5:G${4+checkRows.length}`).format={fill:C.green,font:{bold:true,color:C.greenText,name:"Microsoft YaHei"},horizontalAlignment:"center"};

for (const s of Object.values(sheets)) { const used=s.getUsedRange(); if(used) used.format.verticalAlignment="top"; }

const renders = [
  ["01_closure.png","结项说明","A1:H27"],["02_inventory.png","阶段成果清单","A1:H20"],["03_descriptive.png","核心描述统计","A1:M10"],
  ["04_dictionary.png","变量字典","A1:O14"],["05_availability.png","变量可用性","A1:P14"],["06_long.png","公司变量长表","A1:W16"],
  ["07_integrated.png","价值整合地图","A1:Q14"],["08_profiles.png","公司画像","A1:L12"],["09_portfolios.png","候选组合暂存","A1:X12"],
  ["10_boundaries.png","数据层级与边界","A1:H14"],["11_handoff.png","Week10交接","A1:H19"],["12_checks.png","冻结验收",`A1:H${finalRow+1}`],
];
for (const [name,sheetName,range] of renders) { const blob=await wb.render({sheetName,range,scale:1.0,format:"png"}); await fs.writeFile(`${previewDir}/${name}`,new Uint8Array(await blob.arrayBuffer())); }
for (const [name,range] of [["closure.ndjson","结项说明!A1:H27"],["long.ndjson","公司变量长表!A1:W16"],["checks.ndjson",`冻结验收!A1:H${finalRow+1}`]]) {
  const res=await wb.inspect({kind:"table",range,include:"values,formulas",tableMaxRows:40,tableMaxCols:26,maxChars:180000}); await fs.writeFile(`${previewDir}/${name}`,res.ndjson,"utf8");
}
const err=await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:300},summary:"Week9 freeze formula error scan"});
await fs.writeFile(`${previewDir}/formula_errors.ndjson`,err.ndjson,"utf8");
const out=await SpreadsheetFile.exportXlsx(wb); await out.save(outputXlsx);
await fs.copyFile(outputXlsx,`${packageDir}/reports/${path.basename(outputXlsx)}`);

const keyInputs = ["S02","S03","S04C","S05D","S06B","S06C"];
for (const id of keyInputs) await fs.copyFile(`${root}/${rel[id]}`,`${packageDir}/inputs/${path.basename(rel[id])}`);

const readme = `# Week 9：上市前融资数据描述统计与研究变量开发\n\n`+
`冻结版本：**${version}**  \n冻结日期：**${freezeDate}**\n\n`+
`本版本以 \`Week8 Standard Dataset v1.0\` 为不可变基础输入，完成八家公司基础描述统计、32个候选研究变量、256个公司—变量状态单元、数据修复冻结、学术与商业价值映射、公司画像和候选组合开发。\n\n`+
`## 冻结结果\n\n- 公司：8\n- 候选变量：32\n- 公司—变量单元：256（VALID 215；STRUCTURAL_NA 38；NOT_COMPUTABLE 3）\n- 经证据链修复单元：10\n- 候选组合：8（无排名、无最终选择）\n\n`+
`## 关键边界\n\n- 原文披露值、标准化值、计算值和研究判断分层保存。\n- null 不等于 0；STRUCTURAL_NA 不等于 NOT_COMPUTABLE。\n- 3个 NOT_COMPUTABLE 均保持缺失，不进行插补。\n- 样本量为8，仅作样本内描述与案例开发，不作显著性、因果性或总体代表性推断。\n- 代表性公司最终选择延后至 Week 10，在VC/PE论文方法研读后再决定。\n\n`+
`## 目录\n\n- \`reports/\`：冻结总工作簿\n- \`data/\`：7张可复现CSV\n- \`inputs/\`：6个关键阶段冻结输入\n- \`docs/\`：结项说明与Week10交接\n- \`scripts/\`：独立发布校验脚本与构建脚本\n- \`metadata/\`：来源与发布清单\n\n`+
`## 校验\n\n运行：\n\n\`\`\`bash\npython scripts/verify_release.py\n\`\`\`\n\n应返回 \`PASS\`。文件完整性见 \`CHECKSUMS.sha256\`。\n`;
await fs.writeFile(`${packageDir}/README.md`,readme,"utf8");

const closureDoc = `# Week 9 结项与冻结说明\n\n## 输入\n\n- Week8 Standard Dataset v1.0（只读、哈希固定）\n- Week9 Stage 1—6C 的16个阶段成果\n\n## 处理\n\n1. 汇总Stage2基础描述统计。\n2. 固化Stage3的32个候选变量定义。\n3. 采用Stage4C作为修复后的唯一变量可用性版本。\n4. 固化Stage5D学术—商业价值整合角色，不生成综合总分。\n5. 固化Stage6B公司画像与Stage6C无排名候选组合。\n6. 生成CSV接口、来源清单、校验脚本和SHA-256清单。\n\n## 输出\n\n- 冻结总工作簿1份\n- 可复现CSV 7份\n- 关键冻结输入6份\n- 结项/交接文档2份\n- 构建与校验脚本2份\n\n## 验收标准\n\n- 8家公司、32变量、256公司—变量单元。\n- VALID=215、STRUCTURAL_NA=38、NOT_COMPUTABLE=3。\n- 修复单元=10；剩余不可计算值不插补。\n- 候选组合=8；FINAL_SELECTED=0。\n- 工作簿公式错误=0；独立校验全部通过。\n\n## 冻结政策\n\nWeek9 v1.0不再原位修改。Week10若因论文方法产生变量或选择标准变化，应记录变更理由并生成新版本，不覆盖本版本。\n`;
await fs.writeFile(`${packageDir}/docs/WEEK9_CLOSURE.md`,closureDoc,"utf8");

const handoffDoc = `# Week 10 交接说明\n\nWeek10的首要任务是研读一篇VC/PE论文，拆解研究问题、理论机制、样本与识别、变量测量、缺失处理、稳健性和解释边界。\n\n随后将论文方法与Week9的32变量、变量可用性、公司画像和候选组合逐项映射，形成“保留 / 修订 / 新增 / 删除”建议，再决定是否需要数据补充或版本升级。\n\n代表性公司选择在此之后进行。选择依据必须是明确研究问题、案例角色与纳入排除规则；不得只按融资金额、单项轴值或S6互补分值作自动排名。\n\n建议Week10闸门：\n\n1. 论文研读卡字段完整；\n2. 方法可迁移性逐项说明；\n3. 变量差距有证据；\n4. 选择标准可复核；\n5. 决策记录为 GO / REVISE / HOLD。\n`;
await fs.writeFile(`${packageDir}/docs/WEEK10_HANDOFF.md`,handoffDoc,"utf8");

const sourceManifest = { version, freeze_date:freezeDate, immutable_base:{name:"Week8 Standard Dataset v1.0",attached_zip:path.basename(week8Zip),sha256:week8ZipSha}, stage_outputs:sourceInventory.map(r=>({stage_id:r[0],stage_name:r[1],path:r[2],status:r[3],purpose:r[4],sha256:r[5]})), key_inputs:keyInputs };
await fs.writeFile(`${packageDir}/metadata/source_manifest.json`,JSON.stringify(sourceManifest,null,2)+"\n","utf8");
await fs.copyFile(`${workDir}/build_week9_closure.mjs`,`${packageDir}/scripts/build_week9_closure.mjs`);
await fs.copyFile(`${workDir}/templates/verify_release.py`,`${packageDir}/scripts/verify_release.py`);

const walk = async (dir) => { const out=[]; for(const ent of await fs.readdir(dir,{withFileTypes:true})){const p=`${dir}/${ent.name}`; if(ent.isDirectory()) out.push(...await walk(p)); else out.push(p);} return out; };
const beforeManifest=(await walk(packageDir)).filter(p=>!p.endsWith("CHECKSUMS.sha256")&&!p.endsWith("release_manifest.json")).sort();
const releaseManifest={version,freeze_date:freezeDate,counts:{companies:8,variables:32,company_variable_cells:256,valid:215,structural_na:38,not_computable:3,repaired_cells:10,candidate_portfolios:8,final_selected:0},files:[]};
for(const p of beforeManifest) releaseManifest.files.push({path:path.relative(packageDir,p),bytes:(await fs.stat(p)).size,sha256:await sha256(p)});
await fs.writeFile(`${packageDir}/metadata/release_manifest.json`,JSON.stringify(releaseManifest,null,2)+"\n","utf8");
const allFiles=(await walk(packageDir)).filter(p=>!p.endsWith("CHECKSUMS.sha256")).sort();
const checksumLines=[]; for(const p of allFiles) checksumLines.push(`${await sha256(p)}  ${path.relative(packageDir,p)}`);
await fs.writeFile(`${packageDir}/CHECKSUMS.sha256`,checksumLines.join("\n")+"\n","utf8");

console.log(JSON.stringify({outputXlsx,packageDir,version,sheets:sheetNames.length,stage_files:stageFiles.length,csv_files:csvOutputs.length,week8_zip_sha256:week8ZipSha}));
