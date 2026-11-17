# attendance-fill

为田组固定版式考勤表生成个人出勤记录的 Codex skill。首次保存使用者姓名，每期确认周末、出差和其他例外后，按个人常规规则填写一份新的 Excel 表。

## 下载与安装

[下载最新技能包](https://github.com/210554097-netizen/attendance-fill/releases/latest/download/attendance-fill.zip)

将压缩包中的 `attendance-fill` 文件夹放入本机 Codex 的 `skills` 目录。默认位置如下。

- Windows 使用 `%USERPROFILE%\.codex\skills\attendance-fill`。
- macOS 和 Linux 使用 `~/.codex/skills/attendance-fill`。
- 设置了 `CODEX_HOME` 时，使用 `$CODEX_HOME/skills/attendance-fill`。

从仓库源码安装时，复制仓库内的 `attendance-fill` 文件夹到上述位置。

## 运行环境

完整 Excel 填写流程使用支持本地 skill、`load_workspace_dependencies` 和 `@oai/artifact-tool` 的 Codex 环境。Python 与 Node.js 的可执行文件及工作簿运行库由该环境的依赖加载器提供。

Python 脚本使用标准库完成模板识别、个人配置、历史学习、填写计划和保存后校验。Excel 写入和预览由 `@oai/artifact-tool` 完成。运行目录的依赖链接在 Windows 使用 junction，在 macOS/Linux 使用符号链接，具体步骤见技能说明。

## 使用

上传同版式的 `.xlsx` 模板，向 Codex 发送以下请求。

```text
用 $attendance-fill 填写我的本期出勤表。
```

首次提供姓名后，后续会沿用该姓名。每期集中确认以下信息。

- 各周末是否出勤，可批量回复或指定日期。
- 出差日期及是否出勤。提供本期地点时采用该地点。
- 请假、调休、临时作息等其他例外。

例如可以回复以下内容。

```text
周末全部出勤，12日杭州出差且出勤，无其他例外。
```

如需更新个人规则，上传本人的历史考勤表并发送以下请求。

```text
用这份历史记录更新 $attendance-fill 的常规出勤规则。
```

## 默认规则

| 日期类型 | 上午 | 下午 | 晚上 |
| --- | --- | --- | --- |
| 工作日 | 9:00–11:00 | 12:00–17:00 | 19:00–21:00 |
| 周六出勤 | 9:00–11:00 | 14:00–17:00 | 19:00–21:00 |
| 周日出勤 | 9:00–11:00 | 14:00–17:00 | 19:00–22:00 |

同类日期约 70% 复用基准，约 30% 调整一至两段。变化时，每个起止时刻在基准前后 1.5 小时内按 30 分钟步长调整，结束时间晚于开始时间，各段之间保留至少 30 分钟休息间隔。每类至少有三个自动出勤日时，至少两天重复基准、至少一天发生变化。

日工时按生成时段的分钟数计算，期间合计从总分钟数汇总。出差日的当日总工时采用“地点＋工时”，例如 `杭州9h`。默认出差地点为初始样例中的“嘉兴”，本期地点说明优先。

同一期生成的时段会保存。调整例外时优先复用已有时段，必要时以最少调整恢复重复与变化的分布。明确要求重新生成时可重抽时段。

## 模板与配置

模板每天三行，依次填写上午、下午、晚上。每人使用相邻两列，分别为时段和当日总工时，底部设“合计总工时”。技能根据姓名和日期定位栏目，支持人员顺序与名单变化。

个人配置默认保存在 `$CODEX_HOME/skill-data/attendance-fill/`。`CODEX_HOME` 未设置时，使用 `~/.codex/skill-data/attendance-fill/`。姓名、历史更新和本期状态由本机分别保存。

`attendance-fill/SKILL.md` 为技能入口，`assets/default_rules.json` 为初始规则，`references/rules.md` 说明确认数据与学习口径。

保存后会校验写入值、日工时、期间合计、其他单元格内容、合并和可见版式。当前导出引擎会归一字体字符集元数据，该项会在验证报告中单独列出。
