---
name: powerpoint-charts
description: PptxGenJS chart patterns — bar, line, pie, donut, scatter, area charts with styling and layout guidance.
version: 1.0.0
license: MIT
hosts: [powerpoint]
---

# PowerPoint Charts Skill

Use `slide.addChart()` inside `add_slide_from_code` to create data-driven charts. PptxGenJS supports all major chart types with full styling control.

## Supported Chart Types

| PptxGenJS Constant | Use Case |
| --- | --- |
| `pptx.charts.BAR` | Comparing categories side by side |
| `pptx.charts.BAR3D` | 3D bar charts for visual impact |
| `pptx.charts.LINE` | Trends over time |
| `pptx.charts.PIE` | Part-of-whole (≤6 slices) |
| `pptx.charts.DOUGHNUT` | Modern part-of-whole with center label |
| `pptx.charts.SCATTER` | Correlation between two variables |
| `pptx.charts.AREA` | Volume/magnitude over time |
| `pptx.charts.RADAR` | Multi-axis comparison (e.g. skill ratings) |

## Data Format

All chart types use an array of series objects:

```js
const chartData = [
  {
    name: 'Series 1',
    labels: ['Q1', 'Q2', 'Q3', 'Q4'],
    values: [120, 180, 150, 210]
  },
  {
    name: 'Series 2',
    labels: ['Q1', 'Q2', 'Q3', 'Q4'],
    values: [90, 140, 170, 190]
  }
];
```

- **Bar, Line, Area, Radar**: support multiple series (stacked or grouped).
- **Pie, Doughnut**: use a single series only — multiple series are ignored.
- **Scatter**: each series uses `values` as Y and `labels` as X (both numeric arrays).

## Essential Options

```js
slide.addChart(pptx.charts.BAR, chartData, {
  x: 0.5,               // left edge (inches)
  y: 1.2,               // top edge (inches)
  w: W - 1,             // width (use W variable)
  h: H - 2,             // height (use H variable)
  showTitle: true,
  title: 'Quarterly Revenue',
  titleFontSize: 18,
  showValue: true,       // data labels on bars/points
  showPercent: false,     // percentage labels (pie/donut)
  showLegend: true,
  legendPos: 'b',        // 'b' bottom, 'r' right, 't' top, 'l' left
  valGridLine: { style: 'none' },  // clean look — no gridlines
});
```

### Pie / Doughnut Specific Options

```js
{
  showPercent: true,      // show % on each slice
  showLegend: true,
  legendPos: 'b',
  dataLabelPosition: 'outEnd',  // 'outEnd', 'inEnd', 'ctr', 'bestFit'
  holeSize: 70,           // donut hole size (0–90, 70 is modern)
}
```

### Axis Options (Bar, Line, Area, Scatter)

```js
{
  catAxisLabelFontSize: 12,
  valAxisLabelFontSize: 12,
  catAxisOrientation: 'minMax',
  valAxisOrientation: 'minMax',
  valGridLine: { style: 'none' },
  catGridLine: { style: 'none' },
  showCatAxisTitle: false,
  showValAxisTitle: false,
}
```

## Color Palettes

Use hex values without `#` prefix. Apply via `chartColors` array.

### Corporate Blue

```js
const CORPORATE_BLUE = ['0B5394', '3D85C6', '6FA8DC', 'A4C2F4', 'CFE2F3', '073763'];
```

### Warm

```js
const WARM = ['BF4B28', 'D4782F', 'E8A838', 'F0C75E', 'F5DEB3', '8B2500'];
```

### Nature

```js
const NATURE = ['2E7D32', '43A047', '66BB6A', 'A5D6A7', 'C8E6C9', '1B5E20'];
```

Apply to any chart:

```js
slide.addChart(pptx.charts.BAR, chartData, {
  chartColors: CORPORATE_BLUE,
  // ...other options
});
```

## Chart + Callout Combo Pattern

Place the chart on the left (60% width) and a key insight text box on the right (35% width):

```js
// Chart — left 60%
const chartW = (W - 1) * 0.6;
slide.addChart(pptx.charts.BAR, chartData, {
  x: 0.5, y: 1.2, w: chartW, h: H - 2,
  showTitle: true, title: 'Revenue by Quarter',
  titleFontSize: 18,
  showValue: true,
  showLegend: true, legendPos: 'b',
  chartColors: CORPORATE_BLUE,
  valGridLine: { style: 'none' },
});

// Callout — right 35%
const calloutX = 0.5 + chartW + (W - 1) * 0.05;
const calloutW = (W - 1) * 0.35;
slide.addText([
  { text: 'Key Insight', options: { fontSize: 20, bold: true, color: '0B5394' } },
  { text: '\n\nQ4 revenue grew 40% year-over-year, driven by enterprise adoption.', options: { fontSize: 14, color: '333333' } },
], {
  x: calloutX, y: 1.2, w: calloutW, h: H - 2,
  valign: 'top',
  fill: { color: 'F2F7FC' },
  rectRadius: 0.1,
});
```

## Best Practices

- **Always use `W` and `H` variables** for positioning — never hardcode slide dimensions.
- **Keep charts in the safe area**: 0.5" margins on all sides (x ≥ 0.5, y ≥ 1.0, right edge ≤ W − 0.5, bottom ≤ H − 0.5).
- **Max 6–8 data points** per series for readability. Aggregate if you have more.
- **Title font ≥ 18pt**, axis labels ≥ 12pt for legibility on screen and projector.
- **Use `valGridLine: { style: 'none' }`** for a clean, modern look.
- **Donut charts**: set `holeSize: 70` for a modern appearance. Add a centered text box inside the hole for a key metric.
- **Consistent colors**: pick one palette per presentation and reuse it across all charts.
- **Legend placement**: use `legendPos: 'b'` (bottom) for wide charts, `'r'` (right) for tall charts.

## Complete Examples

### Example 1: Bar Chart with Data Labels

```js
// add_slide_from_code: receives slide, W, H

const COLORS = ['0B5394', '3D85C6', '6FA8DC', 'A4C2F4'];

// Title
slide.addText('Quarterly Revenue by Region', {
  x: 0.5, y: 0.3, w: W - 1, h: 0.6,
  fontSize: 24, bold: true, color: '1A1A1A',
});

// Chart data
const chartData = [
  { name: 'North', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [120, 180, 150, 210] },
  { name: 'South', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [90, 140, 170, 190] },
  { name: 'East',  labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [110, 160, 130, 200] },
  { name: 'West',  labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [80, 120, 160, 175] },
];

// Bar chart
slide.addChart(pptx.charts.BAR, chartData, {
  x: 0.5, y: 1.0, w: W - 1, h: H - 1.8,
  showTitle: false,
  showValue: true,
  valueFontSize: 9,
  showLegend: true,
  legendPos: 'b',
  legendFontSize: 11,
  chartColors: COLORS,
  catAxisLabelFontSize: 12,
  valAxisLabelFontSize: 12,
  valGridLine: { style: 'none' },
});
```

### Example 2: Donut Chart with Percentages

```js
// add_slide_from_code: receives slide, W, H

const COLORS = ['0B5394', '3D85C6', '6FA8DC', 'A4C2F4', 'CFE2F3'];

// Title
slide.addText('Market Share by Segment', {
  x: 0.5, y: 0.3, w: W - 1, h: 0.6,
  fontSize: 24, bold: true, color: '1A1A1A',
});

// Chart data (single series for donut)
const chartData = [
  {
    name: 'Market Share',
    labels: ['Enterprise', 'SMB', 'Consumer', 'Government', 'Education'],
    values: [42, 25, 18, 10, 5],
  },
];

// Donut chart — left side
const chartW = (W - 1) * 0.6;
slide.addChart(pptx.charts.DOUGHNUT, chartData, {
  x: 0.5, y: 1.0, w: chartW, h: H - 1.8,
  showTitle: false,
  holeSize: 70,
  showPercent: true,
  dataLabelPosition: 'outEnd',
  dataLabelFontSize: 12,
  showLegend: true,
  legendPos: 'b',
  legendFontSize: 11,
  chartColors: COLORS,
});

// Center label inside donut hole
slide.addText([
  { text: '42%', options: { fontSize: 32, bold: true, color: '0B5394' } },
  { text: '\nEnterprise', options: { fontSize: 13, color: '555555' } },
], {
  x: 0.5 + chartW * 0.28, y: H * 0.35, w: chartW * 0.44, h: 1.2,
  align: 'center', valign: 'middle',
});

// Callout — right side
const calloutX = 0.5 + chartW + (W - 1) * 0.05;
const calloutW = (W - 1) * 0.35;
slide.addText([
  { text: 'Key Takeaway', options: { fontSize: 18, bold: true, color: '0B5394' } },
  { text: '\n\nEnterprise segment leads at 42%, up 5 points from last year. SMB shows strongest growth trajectory.', options: { fontSize: 13, color: '333333' } },
], {
  x: calloutX, y: 1.0, w: calloutW, h: H - 1.8,
  valign: 'top',
  fill: { color: 'F2F7FC' },
  rectRadius: 0.1,
});
```

### Example 3: Multi-Series Line Chart

```js
// add_slide_from_code: receives slide, W, H

const COLORS = ['0B5394', 'BF4B28', '2E7D32'];

// Title
slide.addText('Monthly Active Users — 2024', {
  x: 0.5, y: 0.3, w: W - 1, h: 0.6,
  fontSize: 24, bold: true, color: '1A1A1A',
});

// Chart data
const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
const chartData = [
  { name: 'Mobile',  labels: months, values: [320, 380, 410, 450, 520, 580] },
  { name: 'Desktop', labels: months, values: [500, 490, 510, 530, 525, 540] },
  { name: 'Tablet',  labels: months, values: [140, 155, 160, 170, 185, 195] },
];

// Line chart
slide.addChart(pptx.charts.LINE, chartData, {
  x: 0.5, y: 1.0, w: W - 1, h: H - 1.8,
  showTitle: false,
  showValue: false,
  showMarker: true,
  markerSize: 6,
  lineSize: 2,
  showLegend: true,
  legendPos: 'b',
  legendFontSize: 11,
  chartColors: COLORS,
  catAxisLabelFontSize: 12,
  valAxisLabelFontSize: 12,
  valGridLine: { style: 'none' },
  catGridLine: { style: 'none' },
});
```
