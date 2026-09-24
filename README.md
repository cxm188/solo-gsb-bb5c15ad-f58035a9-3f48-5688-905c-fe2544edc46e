# 电子表格公式影响分析（Spreadsheet Formula Impact Analysis）

一个可自托管的 Web 应用：上传 Excel（`.xlsx`）文件后，系统解析工作表中的公式，
将**跨工作表引用**、**矩形区域引用**与**名称管理器定义名称**展开为单元格级别的
依赖图，自动检测**循环引用**（含名称循环），并可针对任意单元格列出其**所有直接 /
间接受影响单元格及完整影响路径**。

- 后端：FastAPI + openpyxl
- 前端：原生 HTML / CSS / JavaScript 单页（无需构建，无需 Node）
- 部署：Dockerfile + docker-compose 一键启动

## 功能一览

- 📤 拖拽 / 点选上传 `.xlsx`（上限 25MB，内存中保留最近 5 个文件）
- 🔗 解析单元格引用、矩形区域（`A1:B10`）、整列（`B:B`）、整行（`2:5`）、
  绝对引用（`$A$1`）、跨表引用（`参数!B2`）、带空格的工作表名（`'月报 2026'!B2`）
- 🏷 **名称管理器支持**：解析公式中的定义名称并展开其指向的单元格 / 区域为依赖边
  - 同时识别**工作簿级**与**工作表级**名称；无限定名称按「先本表本地、再工作簿」
    解析，跨表限定名称（`计算!单价`）按「先限定表本地、再工作簿」解析，正确处理同名遮蔽
  - 支持多区域（`A1:A3,C1:C2`）、整行 / 整列、嵌套名称链、公式表达式名称、常量名称
  - **未定义名称、限定到不存在工作表的名称、`#REF!` 失效名称、名称循环**生成可定位的
    审计问题；指向不存在工作表 / 越界的名称目标**不加入影响面**
  - 常量 / 外部工作簿名称不产生依赖边，仅保留解析状态与证据
- 🔁 Tarjan 强连通分量算法检测循环引用（含自环），并单独检测名称定义循环，
  页面顶部高亮展示
- 🔍 选定单元格后展示：
  - **下游影响**：哪些单元格直接 / 间接引用了它（改了它会影响谁）
  - **上游依赖**：该单元格直接 / 间接依赖谁（它的计算需要谁）
  - 每个受影响单元格的**层级（深度）**和一条**最短影响路径**
- 🔑 **关键节点分析**：指定目标单元格后，在其上游依赖子图中计算近似最小割点
  - 先将循环引用的**强连通分量（SCC）压缩为单节点**，再在 DAG 上统计每个节点出现在
    多少条「输入源 → 目标」依赖路径上（覆盖路径数 = 经该节点的路径数，精确计数）
  - 覆盖率 100% 的节点标记为**必经节点**（移除后目标无法计算）；返回每个节点的
    覆盖路径数与占比、**受影响下游**（含目标）与一条**经过该节点的证据路径**
  - 支持按覆盖路径数 / 受影响下游数 / 距目标层级**排序**、按工作表**筛选**、
    覆盖数下限过滤与**导出关键节点 JSON**（路径数极大时附带大整数字符串字段）
- 🗂 可按工作表筛选、按单元格 / 公式文本 / 名称名 / 名称定义搜索、切换是否只看公式单元格
- 📄 查看每个单元格的公式原文、Excel 缓存计算值，及其引用的定义名称与解析状态
- 🩺 **依赖健康审计**：扫描依赖图并按规则定级，支持按级别 / 工作表筛选
  - `error` —— 跨表公式或定义名称指向**不存在的工作表**、公式使用了**未定义名称**、
    定义名称目标**失效（`#REF!` / 空定义 / 名称循环）**
  - `warning` —— 被公式引用却**为空的单元格**、引用**超出目标工作表已用范围**、
    定义名称**指向外部工作簿**、**扇出（直接依赖数）超过阈值**的公式节点
  - `info` —— **未被任何公式引用的孤立公式**
  - 每条问题给出级别、单元格位置、问题说明、涉及名称的作用域 / 定义 / 解析链与
    **证据边**，并附带按级别 / 类型的汇总计数；页面提供审计面板与「导出审计 JSON」
- 🧭 **同列公式偏离检查**：按工作表、列把相邻公式格划为区段，公式按所在格归一为
  模板（引用转为相对位移，`$` 绝对行列、跨表限定、字符串与其余文本原样保留，
  如 `C2` 的 `=A2+$B$1` 与 `C3` 的 `=A3+$B$1` 属同一模板）；区段 ≥ 3 格且某模板
  **严格过半**时标记其余偏离公式，平票 / 无多数不报；结果列出异常坐标与原公式、
  主流公式示例与区段范围，页面列表与 JSON 接口均支持按工作表筛选
- 💾 一键导出完整分析 JSON（工作表统计、**名称清单与名称使用留痕**、公式清单、
  依赖边、循环环、依赖健康审计）
- 📋 **公式区域搬移预演**：选源工作表矩形范围、目标工作表与左上角坐标，按
  **Excel 复制 / 粘贴语义**把范围内公式映射到目标区域，逐格列出原公式、拟写入
  公式与目标已有内容冲突，可下载 TSV 粘贴；不修改已上传工作簿
  - 相对引用按行列位移平移，`$A$1` **绝对引用保持固定**，`$A1` / `A$1` 混合
    引用只平移未锁定维度；区域、整行 / 整列引用同样按锁定状态平移
  - 跨表引用保留表名前缀（同表引用搬到别的表自动补表名，跨表引用搬回该表
    去掉冗余限定）；定义名称按复制语义保留，源表的**工作表级名称**跨表时
    自动改写为「源表!名称」
  - 预演严格基于源 / 目标区域的**一次性快照**计算，源与目标重叠也不会边读边改
  - `#REF!`、指向不存在工作表的引用 / 名称、平移后越过工作表物理边界
    （XFD / 1048576）或未定义 / 失效名称属于**无法平移**：逐格报错且整次预演
    不生成可导出 TSV；引用落入目标表已用范围之外仅提示、不阻断

## 一键启动（Docker Compose）

```bash
docker compose up --build -d
```

然后浏览器打开：

> **http://localhost:8000**

停止：

```bash
docker compose down
```

如需修改端口，编辑 `docker-compose.yml` 中的 `"8000:8000"`（左为主机端口）。

容器启动后可直接验证健康状态与审计接口：

```bash
curl http://localhost:8000/api/health
# 首页下载示例文件上传后，用返回的 workbook_id 查看审计结果
curl -F "file=@samples/sample_impact.xlsx" http://localhost:8000/api/upload
curl "http://localhost:8000/api/workbooks/$WID/audit?level=warning"
```

> 上传的工作簿仅保存在容器内存中（最近 5 个），容器重启即清空；
> 镜像构建时会预生成示例文件 `samples/sample_impact.xlsx`。

## 本地开发（不使用 Docker）

需要 Python 3.11+：

```bash
pip install -r requirements.txt
python scripts/make_sample.py          # 生成示例文件（可选，仓库已附带）
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

运行测试：

```bash
pip install pytest httpx
python -m pytest tests/ -v
```

## 端口与 API

服务监听 **8000** 端口，全部接口：

| 方法 | 路径 | 说明 |
|---|---|---|
| `GET` | `/` | 前端单页 |
| `POST` | `/api/upload` | 上传 xlsx（multipart 字段名 `file`），返回工作簿 id、统计与循环信息 |
| `GET` | `/api/workbooks/{id}` | 工作簿统计摘要 |
| `GET` | `/api/workbooks/{id}/cells?sheet=&q=&formulas_only=` | 单元格列表（支持工作表筛选 / 搜索） |
| `GET` | `/api/workbooks/{id}/impact?cell=参数!B2&direction=downstream` | 影响面分析；`direction` 取 `downstream`（默认，受影响）或 `upstream`（依赖链） |
| `GET` | `/api/workbooks/{id}/key-nodes?cell=&sheet=&sort=&min_coverage=&max_paths=` | 关键节点（上游近似最小割点）；`sort` 取 `coverage`（默认）/`affected`/`depth` |
| `GET` | `/api/workbooks/{id}/key-nodes/export` | 导出关键节点分析 JSON（行数与受影响下游不截断） |
| `GET` | `/api/workbooks/{id}/audit?level=&sheet=&fanout_threshold=` | 依赖健康审计；支持按级别（`error`/`warning`/`info`，可重复）、工作表、扇出阈值筛选 |
| `GET` | `/api/workbooks/{id}/audit/export` | 导出依赖健康审计 JSON（证据边不截断） |
| `GET` | `/api/workbooks/{id}/deviation?sheet=` | 同列公式偏离检查；支持按工作表筛选 |
| `POST` | `/api/workbooks/{id}/relocate/preview` | 公式区域搬移预演（JSON 体：`source_sheet` / `source_range` / `target_sheet` / `target_anchor`） |
| `POST` | `/api/workbooks/{id}/relocate/tsv` | 下载预演 TSV（目标位置 + 拟写入公式；不可平移时返回 400） |
| `GET` | `/api/workbooks/{id}/export` | 导出完整分析 JSON（含 `audit` 段） |
| `GET` | `/samples/sample_impact.xlsx` | 下载示例文件 |
| `GET` | `/api/health` | 健康检查 |

### 接口示例

```bash
# 上传并取得 workbook_id
curl -F "file=@samples/sample_impact.xlsx" http://localhost:8000/api/upload

# 全量审计（默认扇出阈值 20）
curl "http://localhost:8000/api/workbooks/$WID/audit"

# 只看错误与警告、限定「审计示例」工作表、扇出阈值改为 5
curl "http://localhost:8000/api/workbooks/$WID/audit?level=error&level=warning&sheet=审计示例&fanout_threshold=5"

# 下载审计 JSON（证据边完整，不做 100 条截断）
curl -OJ "http://localhost:8000/api/workbooks/$WID/audit/export"

# 同列公式偏离检查（可加 ?sheet=偏离示例 限定工作表）
curl "http://localhost:8000/api/workbooks/$WID/deviation"

# 公式区域搬移预演：源 搬移演示!B2:B6 -> 同表 D2
curl -s -X POST "http://localhost:8000/api/workbooks/$WID/relocate/preview" \
  -H "Content-Type: application/json" \
  -d '{"source_sheet":"搬移演示","source_range":"B2:B6","target_sheet":"搬移演示","target_anchor":"D2"}'

# 下载预演 TSV（目标位置 + 拟写入公式，可直接粘贴进 Excel；存在无法平移的引用时返回 400）
curl -OJ -X POST "http://localhost:8000/api/workbooks/$WID/relocate/tsv" \
  -H "Content-Type: application/json" \
  -d '{"source_sheet":"搬移演示","source_range":"B2:B6","target_sheet":"搬移演示","target_anchor":"D2"}'
```

错误响应均为 4xx：工作簿不存在 / 已被淘汰返回 `404`；
`level` 非法、`fanout_threshold` 不是 1–10000 的整数返回 `400`
（响应体形如 `{"detail": "fanout_threshold 必须是 1-10000 的整数"}`）。
搬移预演的源 / 目标工作表不存在同样为 `404`；区域 / 坐标格式错误、区域过大、
目标区域越过工作表边界为 `400`。引用越界 / `#REF!` / 未定义名称等**单元格级**
错误不使用 4xx：预览接口仍返回 `200`，但响应中 `exportable=false`、`errors`
逐格列出问题，TSV 接口此时返回 `400`（整次预演不生成可导出清单）。

### 关键节点 JSON 结构

`GET /api/workbooks/{id}/key-nodes?cell=汇总!B2`（导出在 `/key-nodes/export`）：

```json
{
  "workbook_id": "ab12cd34ef56",
  "root": "汇总!B2",
  "root_display": "汇总!B2",
  "upstream_node_count": 10,
  "scc_count": 0,
  "source_count": 5,
  "total_paths": 8,
  "total_paths_text": "8",
  "sampled_paths": 8,
  "path_truncated": false,
  "upstream_truncated": false,
  "root_in_cycle": false,
  "total": 10,
  "returned": 10,
  "filters": { "sheet": null, "sort": "coverage", "min_coverage": 1 },
  "key_nodes": [
    {
      "cell": "计算!B6",
      "type": "cell",
      "node_id": "scc:9",
      "sheet": "计算", "coord": "B6",
      "display": "计算!B6",
      "members": ["计算!B6"],
      "member_displays": ["计算!B6"],
      "in_cycle": false,
      "formula": "=B5*参数!B6", "value": null,
      "depth": 1,
      "covered_paths": 8,
      "covered_paths_text": "8",
      "coverage_ratio": 1.0,
      "cuts_all_paths": true,
      "affected_downstream_count": 1,
      "affected_downstream": ["汇总!B2"],
      "affected_downstream_display": ["汇总!B2"],
      "affected_downstream_truncated": false,
      "evidence_path": ["参数!B6", "计算!B6", "汇总!B2"],
      "evidence_path_raw": ["参数!B6", "计算!B6", "汇总!B2"],
      "evidence_components": [
        { "node_id": "scc:2", "type": "cell", "displays": ["参数!B6"] },
        { "node_id": "scc:9", "type": "cell", "displays": ["计算!B6"] },
        { "node_id": "scc:10", "type": "cell", "displays": ["汇总!B2"] }
      ]
    }
  ]
}
```

字段说明：

- `type=scc` 的行表示一个被压缩的**循环分量**：`cell` 为 `null`，`members`
  列出分量内全部单元格，`display` 以 🔁 前缀展示；普通节点 `type=cell`。
- `covered_paths` 是经过该节点的「输入源 → 目标」**路径总数（精确值，DAG 路径
  计数 DP）**，`coverage_ratio` 为其占全部路径的比例；等于 1 时
  `cuts_all_paths=true`，即移除该节点会切断目标的**全部**独立依赖路径（必经节点）。
- `affected_downstream` 是上游子图内该节点可达的下游单元格（**含目标自身**，
  接口最多返回 100 个样本，总数见 `affected_downstream_count`；导出不截断）。
- `evidence_path` 是一条从输入源经该节点到目标的依赖链；路径数量爆炸时证据最多
  枚举前 `max_paths`（默认 200）条，未被枚举节点的证据由 DP 父链拼接，此时
  `path_truncated=true`（覆盖数仍为精确值）。极端工作簿路径数可能超出 JS 安全
  整数范围，请同时参考 `covered_paths_text` / `total_paths_text` 字符串字段。
- 目标自身处于循环中时 `root_in_cycle=true`；上游节点超过 5000 时
  `upstream_truncated=true`。非法 `sort` 返回 422，`cell` 格式错误返回
  400，工作表 / 工作簿不存在返回 404。

### 审计 JSON 结构

```json
{
  "workbook_id": "ab12cd34ef56",
  "issues": [
    {
      "code": "undefined_name",
      "level": "error",
      "sheet": "名称",
      "cell": "名称!B9",
      "cell_display": "名称!B9",
      "message": "定义名称 不存在的名称：工作簿名称管理器中均未定义名称 '不存在的名称'",
      "formula": "=不存在的名称+1",
      "name": "不存在的名称",
      "name_qualifier": null,
      "name_display": "不存在的名称",
      "name_scope": null,
      "name_definition": null,
      "name_status": "undefined",
      "name_chain": [],
      "target_sheet": null,
      "target_display": "不存在的名称",
      "evidence": [
        { "from": "名称!B9", "from_display": "名称!B9",
          "to_display": "不存在的名称", "ref": "不存在的名称" }
      ]
    },
    {
      "code": "missing_sheet",
      "level": "error",
      "sheet": "名称",
      "cell": "名称!B11",
      "message": "引用了不存在的工作表 '不存在的表'（不存在的表!$A$1）（经定义名称 幽灵表名称 展开）",
      "via_name": "幽灵表名称",
      "target_sheet": "不存在的表",
      "target_display": "不存在的表!$A$1",
      "evidence": [
        { "from": "名称!B11", "to_display": "不存在的表!$A$1",
          "ref": "不存在的表!$A$1" }
      ]
    },
    {
      "code": "empty_ref",
      "level": "warning",
      "sheet": "参数",
      "cell": "参数!B8",
      "cell_display": "参数!B8",
      "message": "单元格 参数!B8 为空，但被 2 个公式引用",
      "formula": null,
      "referenced_by": ["汇总!B7", "审计示例!B3"],
      "evidence": [
        { "from": "汇总!B7", "from_display": "汇总!B7",
          "to": "参数!B8", "to_display": "参数!B8" }
      ]
    },
    {
      "code": "high_fanout",
      "level": "warning",
      "cell": "审计示例!B4", "fanout": 200, "threshold": 20,
      "message": "直接引用 200 个单元格，超过扇出阈值 20",
      "evidence": [ { "from": "审计示例!B4", "to": "数据!B2" } ]
    }
  ],
  "total": 58,
  "counts_by_level": { "error": 6, "warning": 7, "info": 45 },
  "counts_by_code": { "missing_sheet": 2, "undefined_name": 2, "invalid_name": 2,
                      "external_name": 1, "empty_ref": 2, "out_of_range": 2,
                      "high_fanout": 2, "orphan_formula": 45 },
  "fanout_threshold": 20,
  "filters": { "levels": ["error", "warning", "info"], "sheet": null }
}
```

问题类型（`code`）与级别：

| code | 级别 | 含义 |
|---|---|---|
| `missing_sheet` | error | 跨表公式或定义名称指向不存在的工作表（证据边只含 `to_display` 文本；经名称展开时带 `via_name`） |
| `undefined_name` | error | 公式使用了名称管理器中不存在的名称，或跨表限定的工作表不存在；问题带 `name*` 字段 |
| `invalid_name` | error | 定义名称存在但目标无效：`#REF!`、空定义、名称间循环引用；目标不加入依赖图 |
| `external_name` | warning | 定义名称指向外部工作簿（`[1]预算.xlsx!...`），不参与本工作簿影响面 |
| `empty_ref` | warning | 目标单元格无公式也无值，按目标单元格聚合所有引用它的公式（含经名称展开）为证据边 |
| `out_of_range` | warning | 显式单元格 / 区域引用（含经名称展开）超出目标工作表已用范围（整行 / 整列引用不判越界） |
| `high_fanout` | warning | 公式直接依赖的去重单元格数 **> 阈值**（默认 20，可用 `fanout_threshold` 调整；名称展开目标同样计入） |
| `orphan_formula` | info | 公式结果未被任何其他公式引用 |

名称解析状态（`name_status`，出现在名称使用留痕与名称类问题中）：

| status | 含义 |
|---|---|
| `resolved` | 已展开为一个或多个具体单元格 / 区域并加入依赖图 |
| `formula` | 公式表达式名称（其中显式引用仍展开为边） |
| `constant` | 数字 / 字符串等常量名称，不产生依赖边 |
| `undefined` | 名称在对应作用域未定义（对应 `undefined_name` 问题） |
| `invalid` | 目标失效 / 名称循环（对应 `invalid_name` 问题），无效目标不入图 |
| `external` | 指向外部工作簿（对应 `external_name` 问题），不入图 |

> 接口返回的单条问题最多附带 100 条证据边（超出时带 `evidence_truncated` /
> `evidence_total`）；`/audit/export` 与完整导出不截断。

### 公式偏离检查 JSON 结构

`GET /api/workbooks/{id}/deviation`（可用 `?sheet=` 限定工作表）：

```json
{
  "workbook_id": "ab12cd34ef56",
  "segments_checked": 8,
  "segments_with_anomalies": 1,
  "anomaly_count": 1,
  "min_segment_size": 3,
  "filters": { "sheet": null },
  "segments": [
    {
      "sheet": "偏离示例",
      "column": "C",
      "range": "C2:C8",
      "display": "偏离示例!C2:C8",
      "cell_count": 7,
      "template_count": 2,
      "majority": {
        "template": "=R[0]C[-2]+R1C2",
        "count": 6,
        "example_cell": "偏离示例!C2",
        "example_display": "偏离示例!C2",
        "example_formula": "=A2+$B$1"
      },
      "anomalies": [
        { "cell": "偏离示例!C5", "coord": "C5", "display": "偏离示例!C5",
          "formula": "=A4+$B$1", "template": "=R[-1]C[-2]+R1C2" }
      ]
    }
  ]
}
```

字段说明：

- 模板串为 R1C1 风格归一结果：相对维写作位移（`R[0]C[-2]` 表示同行、左数
  第 2 列），`$` 绝对维保留原值（`R1C2` 即 `$B$1`）；跨表限定、字符串、
  定义名称与其余文本原样保留，因此 `C2` 的 `=A2+$B$1` 与 `C3` 的 `=A3+$B$1`
  归一结果相同、属同一模板。
- 仅当区段（同工作表、同列、行号相邻）公式格数 ≥ 3 且某模板数量
  **严格超过一半**时，其余模板的公式格进入 `anomalies`；平票或无多数模板的
  区段不报。`segments` 只列出存在偏离的区段，`segments_checked` 为参与检查
  的区段总数。
- `majority` 给出主流模板、命中格数与一个主流公式示例；`range` / `display`
  为区段范围。`sheet` 参数指定不存在的工作表返回 404。

### 导出 JSON 结构

完整导出 `/api/workbooks/{id}/export`：

```json
{
  "filename": "sample_impact.xlsx",
  "summary": { "sheets": [...], "formula_count": 59, "edge_count": 487,
               "defined_name_count": 16, "workbook_scoped_count": 14,
               "sheet_scoped_count": 2, "name_usage_count": 19,
               "cycle_count": 2, "cycles": [["循环!A1","循环!B1"],["循环!C1"]] },
  "defined_names": [
    { "name": "单价", "scope": "workbook", "scope_sheet": null,
      "definition": "参数!$B$2", "hidden": false, "comment": null,
      "usage_status_counts": { "resolved": 2 }, "used_in_formulas": 2 }
  ],
  "name_usages": [
    { "cell": "名称!B1", "name": "单价", "qualifier": null,
      "name_display": "单价", "ref_text": "单价", "status": "resolved",
      "scope": "工作簿级", "scope_id": "workbook",
      "definition": "参数!$B$2", "chain": ["单价"],
      "target_count": 1, "issues": [] }
  ],
  "formulas": {
    "计算!B2": { "formula": "=B2*参数!B4", "value": null, "in_cycle": false,
                 "names": [ ...同 name_usages 单项... ] }
  },
  "edges": [ { "from": "参数!B2", "to": "计算!B2" } ],
  "cycles": [ ["循环!A1", "循环!B1"], ["循环!C1"] ],
  "audit": { "issues": [ ... ], "total": 41, "counts_by_level": { ... } }
}
```

审计导出 `/api/workbooks/{id}/audit/export` 在审计结果外另含
`defined_names`（名称清单）与 `name_usages`（完整名称使用留痕，含解析问题证据）。

> 说明：`value` 取自 Excel 保存时的缓存计算结果。openpyxl 不执行公式计算，
> 由程序新建、从未在 Excel 中打开保存的文件其缓存值为 `null`，属正常现象。

## 示例文件

`samples/sample_impact.xlsx`（由 `scripts/make_sample.py` 生成，也可在首页直接下载）
包含 9 个工作表，覆盖所有分析场景：

| 工作表 | 内容 |
|---|---|
| `参数` | 纯输入项（单价、数量、折扣率、税率、月数） |
| `计算` | 跨表引用链：单价×数量 → 折扣 → 税额 → 月合计 → 年合计；含工作表级名称 `系数` |
| `汇总` | 跨表矩形区域 `SUM(计算!B2:B5)`、整列引用 `SUM(参数!B:B)` |
| `月报 2026` | 带空格的工作表名（公式里写作 `='月报 2026'!...`） |
| `循环` | 故意制造的循环引用：`A1↔B1` 互引、`C1=C1*2` 自环、`D1` 为环的下游 |
| `数据` | 200 行数据表，供审计中的高扇出公式 `SUM(数据!B2:B201)` 引用 |
| `审计示例` | 覆盖直接引用类审计问题：缺失工作表、越界引用、空单元格引用、高扇出、孤立公式与一条正常公式 |
| `名称` | 覆盖名称管理器：工作簿级 / 工作表级同名遮蔽、跨表限定、区域 / 嵌套 / 公式 / 常量名称，
  以及未定义名称、`#REF!`、指向不存在工作表、外部工作簿、名称循环等审计问题 |
| `搬移演示` | 公式搬移预演示例：相对 / 绝对 / 混合引用、区域、跨表引用、定义名称，
  以及一个目标区已有内容（演示覆盖冲突）；把「名称」表 `B9:B15` 作为源则可看到
  无法平移时整次预演不可导出 |
| `偏离示例` | 同列公式偏离检查示例：`C2:C8` 为 `=A{行}+$B$1` 向下填充（`C5` 误引
  上一行，被多数模板标记）；`E` 列两种模板各半（平票不报）、`G` 列仅 2 格
  （不足 3 格区段，不检查） |

推荐体验路径：

1. 首页下载并上传 `sample_impact.xlsx`，顶部可看到红色循环引用提示、定义名称数与审计问题计数；
2. 左侧选择 `参数!B2`（单价），右侧显示受影响单元格（含经名称展开的 `名称!B2` 等）与逐层路径；
3. 切换到「上游依赖」，选择 `'月报 2026'!B3`，可看到完整回溯链直到 `参数!B2`；
4. 用工作表筛选定位 `循环!A1`，观察循环标记与影响面的截断行为；
5. 在「名称」表查看同名作用域遮蔽（`名称!B3` 用本地 `折扣`、`名称!B4` 跨表引用
   `计算!系数`），以及未定义 / 失效 / 外部名称的状态标记；
6. 点击右上角「🩺 依赖健康审计」，在审计面板按级别 / 工作表筛选（试试限定到「名称」表），
   或将扇出阈值调小查看高扇出节点，点击证据边中的单元格可跳转到其影响面；
7. 点击「导出分析 JSON」或「导出审计 JSON」获得机器可读的完整结果（含名称清单与使用留痕）；
8. 选中一个汇总类单元格（如 `汇总!B2`）后点击「🔑 关键节点分析」，查看上游依赖中
   覆盖率最高的必经节点（如 `计算!B6` 覆盖 100% 路径），按受影响下游数排序或限定
   工作表观察，并可展开每条证据路径、导出关键节点 JSON；
9. 点击「📋 公式搬移预演」，源工作表选 `搬移演示`、源范围填 `B2:B6`、目标左上角填
   `D2` 执行：可看到 `B2` 的相对引用平移为 `=C2*2`、`B3` 中 `$A$2` 固定而 `B2`
   平移、目标 `D2` 已有内容被标为覆盖冲突，并可「下载目标坐标与公式 TSV」；
   把目标工作表换成 `月报 2026` 可观察同表引用自动补表名、跨表引用 `计算!B6`
   保持限定；把源换成 `名称` 表 `B9:B15` 则整次预演因未定义 / 失效名称变为不可导出。
10. 点击「🧭 公式偏离检查」，可看到「偏离示例」表 `C2:C8` 区段中 `C5` 偏离主流模板
    （主流 6/7 格为 `=A{行}+$B$1`，`C5` 误写为 `=A4+$B$1`）；`E` 列平票、`G` 列
    不足 3 格均不报。用工作表下拉可筛选检查范围，或访问
    `/api/workbooks/{id}/deviation` 获取同样结构的 JSON。

## 实现要点

- `app/formula_parser.py`：词法式正则扫描公式，区分字符串字面量 / `#REF!` 错误 /
  函数名 / 结构化引用（`表[列]`），避免把 `LOG10`、`"A1"`、`1E2` 误判为单元格；
  输出未展开的 `RawRef` 与定义名称 `RawName`（含跨表限定）。
- `app/names.py`：名称管理器解析。
  - 从 `wb.defined_names`（工作簿级）与各 `ws.defined_names`（工作表级）收集名称，
    忽略 `_xlnm.*` 内置保留名，同名不同作用域并存；
  - 按 Excel 语义解析作用域（本表本地 → 工作簿；限定表本地 → 工作簿），
    顶层逗号切分多区域，区分纯引用 / 嵌套名称 / 公式表达式 / 常量 / `#REF!` / 外部引用；
  - 名称体内未限定引用：工作表级名称相对其所属表，工作簿级名称相对使用公式所在表；
  - 迭代工作清单展开嵌套名称链（上限 `MAX_NAME_DEPTH=50`），用不可变路径集合检测
    名称循环且不误判钻石型引用；返回解析状态、目标 `RawRef`、名称链与可定位问题。
- `app/graph.py`：
  - 加载两次工作簿（`data_only=False` 取公式、`data_only=True` 取缓存值）；
  - 区域展开为具体单元格，整行 / 整列按工作表已用范围截断，单个区域上限 5000 格
    （`MAX_RANGE_CELLS`，防 `SUM(A:XFD)` 之类打爆内存）；
  - 名称目标与直接引用走同一套展开 / 越界 / 缺失工作表逻辑，无效目标一律不入图；
  - 迭代版 Tarjan SCC 检测单元格循环引用（经名称建立的边同样参与）；
  - BFS 求影响面，记录最短路径、层级与环标记，默认最大深度 50 / 5000 节点；
  - 关键节点：在上游诱导子图上跑 Tarjan SCC 并将循环分量压缩为单节点得到 DAG，
    路径计数 DP（`in_paths × out_paths`，Python 大整数）精确计算每个节点覆盖的
    「源 → 目标」依赖路径数与必经节点，位集 DP 求受影响下游，广度枚举前 200 条
    路径作证据（超出部分由 DP 父链拼接）；
  - 构图时记录不入图的异常引用与名称使用留痕，`audit()` 汇总八类问题并按级别 /
    工作表筛选，证据边超过 100 条时截断；导出保留名称、解析状态与完整证据。
- `app/relocate.py`：公式区域搬移预演。
  - 复用 `scan_formula()` 的**带字符区间 token**（含每个端的 `$` 绝对标记），
    从右向左替换片段，字符串 / 函数名 / 结构化引用原样保留；
  - 区域内每个公式格按其自身的源格 -> 目标格行列位移平移：绝对维度固定、相对
    维度平移；跨表引用按目标表决定是否保留 / 去除表限定，同表引用跨表时补源表名；
  - 名称不随复制改定义：仅当无限定名称命中源表的**工作表级名称**且跨表搬移时，
    改写为「源表!名称」限定；未定义 / 失效名称与外部名称经 `NameResolver` 判定；
  - 计算前对源、目标两个矩形各做**一次快照**，几何重叠位置的粘贴前内容取源快照
    （不回读活动单元格），全程纯函数、绝不写工作簿；
  - 平移后越过 XFD / 1048576、`#REF!`、缺失工作表等逐格记入 `errors` 且
    `exportable=false`，TSV 接口此时返回 400；落入已用范围之外只记 `notices`。
- `app/deviation.py`：同列公式偏离检查。
  - 复用 `scan_formula()` 的**带字符区间 token**，把引用按公式所在格归一为
    R1C1 风格模板串（相对维为 `[位移]`、`$` 绝对维保留原值），跨表限定、
    字符串字面量、定义名称与其余非引用文本逐字保留；
  - 按（工作表、列）收集公式格，行号相邻划为区段，仅检查 ≥ 3 格的区段；
    某模板计数**严格过半**时标记区段内其余公式为偏离，平票 / 无多数不报。
- `app/main.py`：FastAPI 路由与内存 LRU 式工作簿缓存（不落盘，重启即清空）；
  上传时同时保留公式形态工作簿对象供搬移预演读取快照。

## 目录结构

```
.
├── app/
│   ├── main.py              # FastAPI 入口与 API
│   ├── formula_parser.py    # 公式引用 / 定义名称词法解析（带字符区间与 $ 标记）
│   ├── names.py             # 名称管理器：作用域解析、展开、名称环检测
│   ├── graph.py             # 依赖图 / 环检测 / 影响面 BFS / 审计
│   ├── relocate.py          # 公式区域搬移预演（快照平移 / 冲突 / TSV）
│   ├── deviation.py         # 同列公式偏离检查（模板归一 / 区段 / 多数规则）
│   └── static/index.html    # 前端单页
├── samples/sample_impact.xlsx
├── scripts/make_sample.py   # 示例文件生成脚本
├── tests/                   # pytest 用例（127 个）
├── Dockerfile
├── docker-compose.yml
└── requirements.txt
```
