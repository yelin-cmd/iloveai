---
name: meeting-weekly-brief
description: Meeting and weekly-report organization for Chinese work contexts. Use when Codex needs to整理会议逐字稿、会议纪要、项目同步内容、行动项、风险、阶段工作汇报、周报或月报，尤其需要按项目提炼“项目进展、核心结论、截至汇总时最新状态、需关注事项、下一步计划”。
---

# Meeting Weekly Brief

## Core Workflow

1. Read the provided meeting transcript, notes, existing project pages, action list, and relevant weekly/monthly summaries.
2. Separate content by project or workstream, not by raw meeting order.
3. For each project, extract:
   - 项目进展
   - 核心结论
   - 截至汇总时最新状态
   - 需关注事项
   - 下一步计划
4. If latest status is unknown or may have changed after the meeting, ask the user for a status supplement before finalizing the weekly report.
5. Keep the final report concise and suitable for direct sharing upward or cross-team.

## Meeting Notes

For single-meeting整理, use these sections unless the user asks otherwise:

- 会议主题
- 会议结论
- 项目进展
- 风险与阻塞
- 行动项
- 待确认问题

When the meeting has mixed topics, split first by topic, such as:

- 工作内容汇报
- AI探索研究内容
- 课程交付风险
- 组织与协同事项

## Weekly Report Style

Use this fixed structure for each project:

```markdown
**项目名称**

项目进展：

核心结论：

截至汇总时最新状态：

需关注事项：

下一步计划：
```

Prefer tight paragraphs or short bullets. Do not dump long transcript details. Mention names, dates, and owners only when they matter for execution.

## Latest Status Collection

Before finalizing a weekly report, identify open or time-sensitive items and ask the user for updates in a compact checklist.

Use statuses:

- 已完成
- 进行中
- 无进展
- 已有结论
- 风险解除
- 风险升级
- 需要继续推动

Ask only about high-impact unknowns that could change the report.

## Action Item Rules

When extracting action items, include:

- 事项
- 负责人
- 截止时间
- 关联项目
- 来源会议
- 状态

If owner or deadline is unclear, mark as `待确认` instead of inventing.

## Tone And Output

Write in concise professional Chinese. The output should feel like a practical work report, not a transcript summary.

Avoid:

- Overly long background
- Unverified conclusions
- Repeating every speaker’s words
- Mixing AI exploration with normal project progress unless they are the same workstream

Use `references/report-template.md` when a stable report template is needed.

