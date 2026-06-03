# DESIGN.md

## Design Direction

This project is a local workspace for market news collection, review, and client-ready reporting. The interface should feel calm, precise, and helpful: closer to a light productivity tool than a marketing page.

Use **Cal.com + Notion** as the main visual reference:

- Cal.com influence: clean structure, neutral surfaces, restrained controls, clear workflows.
- Notion influence: soft reading rhythm, warm off-white backgrounds, approachable content blocks, low-pressure information density.

The product should look trustworthy and quiet. It should help a non-technical user understand what to do next, while still giving developers enough density and diagnostics in developer mode.

## Product Personality

- Calm, lightweight, and professional.
- Friendly without being cute.
- Structured without feeling like a heavy enterprise backend.
- Designed for repeated use: scanning, filtering, selecting, exporting.
- Content-first: news cards, form labels, and generated outputs are more important than decoration.

Avoid:

- Dark-mode-first design.
- Purple-heavy SaaS gradients.
- Oversized hero sections.
- Decorative cards inside cards.
- Marketing-page layout patterns.
- Dense technical labels in ordinary user mode.

## Modes

### Ordinary User Mode

Ordinary user mode is a guided workflow. It should feel like a small operations desk:

1. Configure API if needed.
2. Choose country, time range, and brands.
3. Start crawling and watch progress.
4. Review news.
5. Add manual news if needed.
6. Export client materials.

Design priorities:

- Make the next action obvious.
- Use plain Chinese labels.
- Hide diagnostics unless they help the current task.
- Prefer friendly batch names over internal run IDs.
- Keep navigation sticky and simple.
- Cards should be readable at a glance.

### Developer Mode

Developer mode is a configuration and diagnosis workspace. It can be denser, but still should stay orderly.

Design priorities:

- Preserve full control over aliases, keywords, prompts, countries, and sources.
- Keep diagnostic text scannable.
- Use compact grids and tables where useful.
- Make saved/default/custom states visible.
- Avoid making the developer page visually louder than the user page.

## Color System

Use a warm, light neutral base with muted teal as the primary accent.

Recommended CSS variables:

```css
:root {
  --bg: #f7f5f0;
  --surface: #ffffff;
  --surface-soft: #fbfaf7;
  --surface-muted: #f0ede6;
  --border: #ded8cc;
  --border-strong: #cfc6b8;
  --text: #22201c;
  --text-soft: #5f5a52;
  --muted: #827b70;
  --primary: #357f78;
  --primary-strong: #276760;
  --primary-soft: #e1f0ed;
  --warning: #b7791f;
  --warning-soft: #fff4d8;
  --danger: #b54747;
  --danger-soft: #fbe7e7;
  --success: #2f7d52;
  --success-soft: #e4f3ea;
  --neutral-chip: #e8e4dc;
}
```

Usage:

- Background: `--bg`
- Page sections: transparent or full-width bands, not floating nested cards.
- Repeated articles, modals, and contained tools: `--surface`
- Primary buttons: `--primary`
- Secondary buttons: white surface with border.
- Status chips should use soft backgrounds and clear text.

Sentiment colors:

- Positive: green text on soft green.
- Neutral: warm gray text on neutral chip.
- Negative: red text on soft red.

## Typography

Use a clean Chinese-friendly stack, avoiding overly generic default UI feeling:

```css
font-family:
  "LXGW WenKai Screen",
  "Noto Sans SC",
  "Microsoft YaHei",
  "PingFang SC",
  sans-serif;
```

If `LXGW WenKai Screen` is unavailable, the rest of the stack should still render cleanly.

Type scale:

- Page title: 24-28px, weight 700.
- Section title: 18-20px, weight 700.
- Card title and article body: 14-15px, consistent line height.
- Form labels: 12-13px, weight 600, muted.
- Helper text: 12-13px, muted.
- Chips: 12px, weight 700.

Rules:

- Do not scale font size with viewport width.
- Keep letter spacing at 0.
- Avoid very large type inside tool surfaces.
- In news cards, article title, translated title, metrics, and related explanation should use the same font size; hierarchy comes from weight, color, and spacing.

## Layout

Use a calm workspace layout:

- Max content width: 1180-1280px.
- Page padding: 20-28px desktop, 14-16px mobile.
- Section spacing: 20-28px.
- Card radius: 8px or less.
- Border: 1px solid `--border`.
- Shadow: very subtle, only where separation is needed.

Recommended shadow:

```css
box-shadow: 0 8px 24px rgba(34, 32, 28, 0.06);
```

Avoid:

- Cards inside cards.
- Oversized panels for simple controls.
- Three-column form grids when fields become cramped.
- Dense walls of text without section rhythm.

Responsive behavior:

- Desktop: two-column forms where useful.
- Desktop: article cards can remain single-column for readability.
- Mobile: everything becomes single-column.
- Long text fields always span the full row.

## Navigation

Ordinary user mode should have a sticky navigation bar:

- 总览
- 新闻抓取
- 新闻查看
- 添加新闻
- 导出材料

Developer mode can keep its fuller navigation:

- 总览
- 新闻抓取
- 新闻查看
- 国家管理
- 来源管理

Sticky navigation style:

- Soft surface with slight transparency.
- Thin border.
- High enough z-index to stay above cards.
- Horizontal scroll or wrapping on mobile.
- Scroll targets need `scroll-margin-top`.

## Components

### Buttons

Primary:

- Used for `开始抓取`, `生成资讯表`, `生成新闻总结`, `保存`.
- Solid teal background.
- White text.
- Weight 700.

Secondary:

- Used for refresh, test API, expand/collapse, restore previous.
- White background.
- Neutral border.

Danger:

- Used only for destructive or disabling actions.
- Soft red styling unless action is final.

Button shape:

- Border radius: 8px.
- Height: 36-42px.
- Keep labels short.

### Forms

Labels should be concise and human:

- `国家`
- `时间范围`
- `品牌`
- `数据批次`
- `关键词`
- `显示数量`

Developer-only technical labels can remain technical, but should still be grouped clearly.

Inputs:

- Use white surfaces.
- Keep borders visible but soft.
- Focus state uses teal border and a subtle ring.
- Password visibility toggles should sit near the key input.

### Chips

Use chips for:

- Platform name.
- Sentiment.
- Publish date.
- Industry trend.
- Manual added.
- AI confidence.

Chips should be compact and bold. They are metadata, not paragraphs.

### News Cards

News cards are central to the product.

Structure:

1. First row: platform, sentiment, industry trend/manual tags, date.
2. Original title as link.
3. Translated title if available.
4. NPS dimensions.
5. Related explanation.
6. Optional selection checkbox in a stable position.

Rules:

- Platform, sentiment, and date stay on the first row.
- First-row chips are bold.
- Metrics and related explanation are normal weight.
- Title, translated title, metrics, and explanation use the same font size.
- Avoid truncating important article text too aggressively.
- Manual-added news should be clearly marked.

### Progress

Progress should be reassuring and simple.

Ordinary user mode:

- Show one clear elapsed/remaining time line.
- Progress bar below.
- Current stage text below the bar.
- Avoid duplicate timing text.

Developer mode:

- Can show richer timing and diagnostic stage details.

### Modals And Expandable Areas

Use expandable sections for lower-frequency tasks:

- Add manual news.
- Advanced API configuration.
- Country management.
- Source management.

Expandable areas should have:

- Clear title.
- One-line summary.
- Chevron or explicit expand button.
- Smooth but minimal transition.

## Copywriting

Ordinary user mode should use task-oriented Chinese.

Good:

- `选择一次抓取结果，后续查看、添加新闻和导出材料都会基于该批次。`
- `当前筛选命中 28 条，展示 28 条。`
- `资讯表已生成。`
- `请先配置 AI API，或联系管理员在开发者模式配置。`

Avoid:

- Raw run IDs as primary labels.
- CLI-style wording.
- Prompt engineering terms in ordinary user mode.
- Long diagnostic paragraphs near primary actions.

Developer mode may expose technical terms when they are useful:

- `metadata`
- `AI 筛选提示词`
- `关键词块`
- `来源适配器`
- `SQLite 批次`

## Data Batch Labels

Ordinary user mode should never show long internal directory names as the main label.

Preferred format:

```text
最近一次抓取｜意大利｜2026-05-03 至 2026-06-02｜26条
2026-06-02 05:12 抓取｜意大利｜2026-05-03 至 2026-06-02｜21条
历史结果｜意大利｜时间范围未知｜CSV
```

Internal values can remain unchanged in hidden form values and option values.

## Tables And Dense Data

Use tables only when comparison matters.

For ordinary user mode:

- Prefer cards and simple lists.
- Hide before/after technical distinction unless needed.

For developer mode:

- Tables are acceptable for source management, diagnostics, and configuration lists.
- Keep row height moderate.
- Use muted helper text for secondary metadata.

## Accessibility

- Maintain clear focus states.
- Do not rely on color alone for sentiment; keep text labels.
- Form controls need visible labels.
- Buttons must have obvious disabled states.
- Contrast should remain readable on warm backgrounds.
- Sticky nav must not cover section headings after navigation.

## Implementation Guidance

When updating templates:

- Prefer improving existing classes before adding many new ones.
- Keep ordinary user and developer styles separable when behavior differs.
- Preserve existing `id`, `name`, and `value` attributes unless backend logic is being updated.
- User-facing brand labels can be full names, while backend values remain stable internal IDs.
- Do not change output file schemas for visual-only work.

When changing UI logic:

- Verify ordinary user mode first.
- Then check developer mode did not regress.
- Run at least:

```powershell
python -m py_compile market_news_crawler/web_app.py
python -m unittest discover -s market_news_crawler/tests -v
```

## Design Acceptance Checklist

Before considering a UI change complete:

- Ordinary user mode has no visible mojibake or raw technical labels.
- Main workflow is visible without reading documentation.
- News cards are readable and metadata is stable.
- Forms do not squeeze into awkward three-column layouts.
- Sentiment chips have color and text labels.
- Sticky navigation works on desktop and mobile.
- API configuration is visible when needed but not overwhelming.
- Developer mode still exposes advanced controls.
- No output files, API keys, DB files, or local settings are committed.
