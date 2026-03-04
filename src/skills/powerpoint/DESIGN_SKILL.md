---
name: powerpoint-design
description: Modern slide design patterns — gradients, shadows, card layouts, dashboards, and professional color guidance.
version: 1.0.0
license: MIT
hosts: [powerpoint]
---

# Design Skill

Activate this skill when creating visually polished slides with `add_slide_from_code`. Covers advanced PptxGenJS techniques for professional presentations.

## Gradient Fills

PptxGenJS does not reliably support `fill: { type: 'gradient', ... }` across all renderers. Use **overlapping shapes with varying opacity** to simulate gradients:

```js
// Simulated gradient: stack 4 semi-transparent strips left to right
const baseColor = '4472C4';
const strips = 4;
const stripW = W / strips;
for (let i = 0; i < strips; i++) {
  slide.addShape('rect', {
    x: i * stripW, y: 0, w: stripW + 0.01, h: H,
    fill: { color: baseColor },
    line: { color: baseColor, width: 0 },
    opacity: 0.3 + (i * 0.2),
  });
}
```

**Key points:**
- Overlap strips by 0.01" to avoid hairline gaps
- Opacity range: 0.2 (lightest) to 1.0 (darkest)
- Use `line: { color: baseColor, width: 0 }` to hide shape borders
- For vertical gradients, stack strips top to bottom instead

## Shadow Effects

Add depth to shapes and text boxes with the `shadow` property:

```js
slide.addShape('roundRect', {
  x: 1, y: 1, w: 4, h: 2.5,
  fill: { color: 'FFFFFF' },
  rectRadius: 0.1,
  shadow: { type: 'outer', blur: 3, offset: 2, color: '000000', opacity: 0.3 },
});
```

**Shadow presets:**
| Style        | blur | offset | opacity | Use for               |
|------------- |------|--------|---------|---------------------- |
| Subtle       | 2    | 1      | 0.15    | Cards, text boxes     |
| Medium       | 3    | 2      | 0.3     | Elevated panels       |
| Dramatic     | 5    | 3      | 0.4     | Hero elements, CTAs   |
| Floating     | 8    | 4      | 0.25    | Floating cards        |

Always use `color: '000000'` for shadows. Adjust `opacity` for intensity.

## Card Layout Pattern

Three-column card grid with rounded rect backgrounds, icon area, title, and description.

**Structure per card:**
1. Rounded rect background (`rectRadius: 0.1`, subtle shadow)
2. Colored icon strip or circle at top
3. Title text (bold, 16-18pt)
4. Description text (regular, 12-14pt)

```js
const cards = [
  { icon: '🚀', title: 'Performance', desc: 'Blazing fast load times under 2s', color: '4472C4' },
  { icon: '🔒', title: 'Security', desc: 'Enterprise-grade encryption at rest', color: '548235' },
  { icon: '📊', title: 'Analytics', desc: 'Real-time dashboards and insights', color: 'BF8F00' },
];
const margin = 0.5;
const gap = 0.25;
const cardW = (W - 2 * margin - (cards.length - 1) * gap) / cards.length;
const cardH = 3.5;
const cardY = (H - cardH) / 2;

cards.forEach((c, i) => {
  const x = margin + i * (cardW + gap);
  // Card background
  slide.addShape('roundRect', {
    x, y: cardY, w: cardW, h: cardH,
    fill: { color: 'FFFFFF' },
    rectRadius: 0.1,
    shadow: { type: 'outer', blur: 3, offset: 2, color: '000000', opacity: 0.2 },
  });
  // Colored accent strip at top
  slide.addShape('roundRect', {
    x: x + 0.15, y: cardY + 0.2, w: cardW - 0.3, h: 0.06,
    fill: { color: c.color },
    rectRadius: 0.03,
  });
  // Icon
  slide.addText(c.icon, {
    x, y: cardY + 0.45, w: cardW, h: 0.6,
    align: 'center', fontSize: 28, shrinkText: true,
  });
  // Title
  slide.addText(c.title, {
    x: x + 0.2, y: cardY + 1.15, w: cardW - 0.4, h: 0.5,
    fontSize: 18, bold: true, color: '1F1F1F', align: 'center', shrinkText: true,
  });
  // Description
  slide.addText(c.desc, {
    x: x + 0.2, y: cardY + 1.7, w: cardW - 0.4, h: 1.4,
    fontSize: 13, color: '666666', align: 'center', valign: 'top', shrinkText: true,
  });
});
```

## Dashboard Layout

Four-quadrant layout with KPI numbers, sparkline-style accent shapes, and labels.

**Structure per quadrant:**
1. Background rect with subtle fill
2. Large KPI number (32-36pt, bold)
3. KPI label (12-14pt, muted color)
4. Trend indicator shape (small colored bar or arrow)

```js
const kpis = [
  { value: '$2.4M', label: 'Revenue', trend: '+12%', color: '548235' },
  { value: '1,847', label: 'New Users', trend: '+8%', color: '4472C4' },
  { value: '94.2%', label: 'Uptime', trend: '+0.3%', color: '7B57A0' },
  { value: '4.8★', label: 'Rating', trend: '+0.2', color: 'BF8F00' },
];
const margin = 0.5;
const gap = 0.25;
const quadW = (W - 2 * margin - gap) / 2;
const quadH = (H - 2 * margin - gap - 0.8) / 2;
const startY = 1.1;

// Title
slide.addText('Dashboard Overview', {
  x: margin, y: 0.3, w: W - 2 * margin, h: 0.6,
  fontSize: 24, bold: true, color: '1F1F1F', shrinkText: true,
});

kpis.forEach((kpi, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = margin + col * (quadW + gap);
  const y = startY + row * (quadH + gap);

  // Quadrant background
  slide.addShape('roundRect', {
    x, y, w: quadW, h: quadH,
    fill: { color: 'F5F5F5' },
    rectRadius: 0.08,
    shadow: { type: 'outer', blur: 2, offset: 1, color: '000000', opacity: 0.1 },
  });
  // Accent bar at top
  slide.addShape('rect', {
    x: x + 0.2, y: y + 0.15, w: 0.5, h: 0.06,
    fill: { color: kpi.color },
  });
  // KPI value
  slide.addText(kpi.value, {
    x: x + 0.2, y: y + 0.35, w: quadW - 0.4, h: 0.8,
    fontSize: 34, bold: true, color: '1F1F1F', shrinkText: true,
  });
  // KPI label
  slide.addText(kpi.label, {
    x: x + 0.2, y: y + 1.15, w: quadW - 0.4, h: 0.4,
    fontSize: 13, color: '888888', shrinkText: true,
  });
  // Trend indicator
  slide.addText(kpi.trend, {
    x: x + 0.2, y: y + quadH - 0.55, w: 1.0, h: 0.35,
    fontSize: 12, bold: true, color: kpi.color, shrinkText: true,
  });
  // Sparkline-style bar
  slide.addShape('rect', {
    x: x + quadW - 1.5, y: y + quadH - 0.4, w: 1.2, h: 0.15,
    fill: { color: kpi.color }, opacity: 0.2,
  });
  slide.addShape('rect', {
    x: x + quadW - 1.5, y: y + quadH - 0.4, w: 0.7 + i * 0.1, h: 0.15,
    fill: { color: kpi.color },
  });
});
```

## Hero Slide

Full-bleed colored background with large centered title and subtitle.

```js
// Full-bleed background
slide.addShape('rect', {
  x: 0, y: 0, w: W, h: H,
  fill: { color: '1B2A4A' },
});
// Decorative accent bar
slide.addShape('rect', {
  x: W / 2 - 1, y: H / 2 - 1.2, w: 2, h: 0.06,
  fill: { color: '4A9BD9' },
});
// Title
slide.addText('Transform Your Business', {
  x: 1, y: H / 2 - 1.0, w: W - 2, h: 1.0,
  fontSize: 36, bold: true, color: 'FFFFFF',
  align: 'center', valign: 'middle', shrinkText: true,
});
// Subtitle
slide.addText('A data-driven approach to sustainable growth', {
  x: 1.5, y: H / 2 + 0.2, w: W - 3, h: 0.7,
  fontSize: 20, color: 'B0C4DE',
  align: 'center', valign: 'top', shrinkText: true,
});
```

## Section Divider

Half-color, half-white split with section number and title.

```js
// Left half — colored
slide.addShape('rect', {
  x: 0, y: 0, w: W / 2, h: H,
  fill: { color: '2D5F8A' },
});
// Right half — white
slide.addShape('rect', {
  x: W / 2, y: 0, w: W / 2, h: H,
  fill: { color: 'FFFFFF' },
});
// Section number on left
slide.addText('02', {
  x: 0.8, y: H / 2 - 1.0, w: W / 2 - 1.6, h: 1.2,
  fontSize: 64, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle',
  opacity: 0.3, shrinkText: true,
});
// Section title on right
slide.addText('Market Analysis', {
  x: W / 2 + 0.8, y: H / 2 - 0.8, w: W / 2 - 1.6, h: 0.8,
  fontSize: 32, bold: true, color: '1F1F1F', shrinkText: true,
});
// Subtitle on right
slide.addText('Competitive landscape and opportunities', {
  x: W / 2 + 0.8, y: H / 2 + 0.2, w: W / 2 - 1.6, h: 0.6,
  fontSize: 16, color: '888888', shrinkText: true,
});
// Accent line
slide.addShape('rect', {
  x: W / 2 + 0.8, y: H / 2 + 0.95, w: 1.5, h: 0.05,
  fill: { color: '2D5F8A' },
});
```

## Professional Color Palettes

Use these hex values (without `#` prefix) for cohesive, professional presentations.

### Corporate Blue
| Role       | Hex      | Use for                        |
|----------- |--------- |------------------------------- |
| Primary    | 1B3A5C   | Headers, dark backgrounds      |
| Secondary  | 2D6DA4   | Accents, buttons, links        |
| Tertiary   | 4A9BD9   | Charts, highlights             |
| Light      | B8D4E8   | Card backgrounds, light fills  |
| Neutral    | 5A6978   | Body text, muted labels        |
| Background | F0F4F8   | Slide backgrounds, panels      |

### Warm
| Role       | Hex      | Use for                        |
|----------- |--------- |------------------------------- |
| Primary    | 8B4513   | Headers, anchors               |
| Secondary  | D2691E   | Accents, highlights            |
| Tertiary   | E8A84C   | Charts, callouts               |
| Light      | F5DEB3   | Card fills, soft backgrounds   |
| Neutral    | 6B5B4F   | Body text, muted labels        |
| Background | FDF6EC   | Slide backgrounds              |

### Nature
| Role       | Hex      | Use for                        |
|----------- |--------- |------------------------------- |
| Primary    | 2D5F2D   | Headers, dark backgrounds      |
| Secondary  | 548235   | Accents, key shapes            |
| Tertiary   | 8FBC8F   | Charts, highlights             |
| Light      | C8E6C9   | Card fills, panels             |
| Neutral    | 5D6B5D   | Body text, muted labels        |
| Background | F1F8E9   | Slide backgrounds              |

### Modern Dark
| Role       | Hex      | Use for                        |
|----------- |--------- |------------------------------- |
| Primary    | 1A1A2E   | Slide backgrounds              |
| Secondary  | 16213E   | Panel backgrounds              |
| Accent 1   | E94560   | Highlights, KPI values         |
| Accent 2   | 0F3460   | Charts, secondary accents      |
| Text       | E8E8E8   | Body text on dark              |
| Muted      | 7A7A8E   | Labels, captions on dark       |

### Minimal
| Role       | Hex      | Use for                        |
|----------- |--------- |------------------------------- |
| Primary    | 1F1F1F   | Headings                       |
| Secondary  | 4A4A4A   | Subheadings                    |
| Body       | 6B6B6B   | Body text                      |
| Light      | D4D4D4   | Borders, dividers              |
| Accent     | 2D7DD2   | Single pop of color, links     |
| Background | FAFAFA   | Slide backgrounds              |

## Typography Hierarchy

| Element    | Size      | Weight | Color (Minimal palette)        |
|----------- |---------- |------- |------------------------------- |
| Title      | 32–36pt   | Bold   | 1F1F1F (or white on dark bg)   |
| Subtitle   | 20–24pt   | Normal | 4A4A4A                         |
| Body       | 16–18pt   | Normal | 4A4A4A                         |
| Caption    | 12–14pt   | Normal | 6B6B6B                         |

**Rules:**
- Never use more than 3 font sizes on one slide
- Titles are always bold; body text is never bold (except inline labels)
- Use `shrinkText: true` on every `addText()` call

## Spacing Rules

| Element          | Value    | Notes                          |
|----------------- |--------- |------------------------------- |
| Slide margins    | 0.5"     | All four edges                 |
| Content gutters  | 0.2–0.3" | Between columns/cards          |
| Shape spacing    | 0.15–0.25" | Between vertically stacked shapes |
| Bottom buffer    | 0.3"     | Never fill to y+h = H exactly |
| Card padding     | 0.15–0.2" | Inside card backgrounds        |

**Key rules:**
- Keep all content within `x ≥ 0.5`, `y ≥ 0.5`, `x+w ≤ W-0.5`, `y+h ≤ H-0.5`
- Use `(W - 2*margin - (n-1)*gap) / n` formula for equal-width columns
- Center content vertically: `y = (H - contentHeight) / 2`

## Complete Example: 3-Column Card Layout

```js
// --- Full add_slide_from_code function body ---
// Background
slide.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: 'F0F4F8' } });

// Title
slide.addText('Our Core Services', {
  x: 0.5, y: 0.4, w: W - 1, h: 0.7,
  fontSize: 32, bold: true, color: '1B3A5C', align: 'center', shrinkText: true,
});
// Subtitle
slide.addText('Delivering excellence across three key areas', {
  x: 1, y: 1.1, w: W - 2, h: 0.5,
  fontSize: 16, color: '5A6978', align: 'center', shrinkText: true,
});

// Card data
const cards = [
  { icon: '💡', title: 'Strategy', desc: 'Data-driven planning and market positioning to accelerate growth.', accent: '1B3A5C' },
  { icon: '⚙️', title: 'Engineering', desc: 'Scalable architecture and cloud-native development for reliability.', accent: '2D6DA4' },
  { icon: '📈', title: 'Growth', desc: 'Performance marketing and analytics that deliver measurable ROI.', accent: '4A9BD9' },
];

const margin = 0.5;
const gap = 0.3;
const cardW = (W - 2 * margin - (cards.length - 1) * gap) / cards.length;
const cardH = 3.8;
const cardY = 1.9;

cards.forEach((c, i) => {
  const x = margin + i * (cardW + gap);

  // Card background with shadow
  slide.addShape('roundRect', {
    x, y: cardY, w: cardW, h: cardH,
    fill: { color: 'FFFFFF' },
    rectRadius: 0.1,
    shadow: { type: 'outer', blur: 3, offset: 2, color: '000000', opacity: 0.15 },
  });

  // Top accent bar
  slide.addShape('roundRect', {
    x: x + 0.2, y: cardY + 0.2, w: cardW - 0.4, h: 0.06,
    fill: { color: c.accent },
    rectRadius: 0.03,
  });

  // Icon circle background
  slide.addShape('ellipse', {
    x: x + cardW / 2 - 0.4, y: cardY + 0.5, w: 0.8, h: 0.8,
    fill: { color: c.accent }, opacity: 0.1,
  });

  // Icon
  slide.addText(c.icon, {
    x: x + cardW / 2 - 0.4, y: cardY + 0.5, w: 0.8, h: 0.8,
    align: 'center', valign: 'middle', fontSize: 26, shrinkText: true,
  });

  // Title
  slide.addText(c.title, {
    x: x + 0.2, y: cardY + 1.5, w: cardW - 0.4, h: 0.5,
    fontSize: 18, bold: true, color: '1F1F1F', align: 'center', shrinkText: true,
  });

  // Description
  slide.addText(c.desc, {
    x: x + 0.2, y: cardY + 2.1, w: cardW - 0.4, h: 1.3,
    fontSize: 13, color: '5A6978', align: 'center', valign: 'top', shrinkText: true,
  });
});
```

## Complete Example: KPI Dashboard

```js
// --- Full add_slide_from_code function body ---
// Background
slide.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: '1A1A2E' } });

// Header
slide.addText('Q4 Performance Dashboard', {
  x: 0.5, y: 0.3, w: W - 1, h: 0.6,
  fontSize: 24, bold: true, color: 'E8E8E8', shrinkText: true,
});
slide.addText('Updated December 2024', {
  x: 0.5, y: 0.85, w: W - 1, h: 0.35,
  fontSize: 12, color: '7A7A8E', shrinkText: true,
});

// KPI data
const metrics = [
  { value: '$12.8M', label: 'Annual Revenue', trend: '▲ 24%', trendColor: '00C781', bar: 0.85 },
  { value: '3,241', label: 'Active Customers', trend: '▲ 18%', trendColor: '00C781', bar: 0.72 },
  { value: '99.7%', label: 'Service Uptime', trend: '▲ 0.4%', trendColor: '00C781', bar: 0.95 },
  { value: '47 min', label: 'Avg Response Time', trend: '▼ 12%', trendColor: 'E94560', bar: 0.55 },
];

const margin = 0.5;
const gap = 0.3;
const panelW = (W - 2 * margin - (metrics.length - 1) * gap) / metrics.length;
const panelH = 4.2;
const panelY = 1.5;

metrics.forEach((m, i) => {
  const x = margin + i * (panelW + gap);

  // Panel background
  slide.addShape('roundRect', {
    x, y: panelY, w: panelW, h: panelH,
    fill: { color: '16213E' },
    rectRadius: 0.1,
  });

  // KPI value
  slide.addText(m.value, {
    x: x + 0.15, y: panelY + 0.4, w: panelW - 0.3, h: 0.9,
    fontSize: 34, bold: true, color: 'FFFFFF', align: 'center', shrinkText: true,
  });

  // KPI label
  slide.addText(m.label, {
    x: x + 0.15, y: panelY + 1.3, w: panelW - 0.3, h: 0.4,
    fontSize: 12, color: '7A7A8E', align: 'center', shrinkText: true,
  });

  // Divider line
  slide.addShape('rect', {
    x: x + 0.25, y: panelY + 1.85, w: panelW - 0.5, h: 0.02,
    fill: { color: '7A7A8E' }, opacity: 0.3,
  });

  // Trend indicator
  slide.addText(m.trend, {
    x: x + 0.15, y: panelY + 2.0, w: panelW - 0.3, h: 0.4,
    fontSize: 14, bold: true, color: m.trendColor, align: 'center', shrinkText: true,
  });

  // Sparkline bar background
  slide.addShape('roundRect', {
    x: x + 0.25, y: panelY + 2.7, w: panelW - 0.5, h: 0.2,
    fill: { color: 'FFFFFF' }, opacity: 0.08,
    rectRadius: 0.05,
  });

  // Sparkline bar filled portion
  slide.addShape('roundRect', {
    x: x + 0.25, y: panelY + 2.7, w: (panelW - 0.5) * m.bar, h: 0.2,
    fill: { color: m.trendColor },
    rectRadius: 0.05,
  });

  // Mini bar chart (3 bars for visual interest)
  const bars = [0.5, 0.7, m.bar];
  const barGap = 0.08;
  const barAreaW = panelW - 0.5;
  const barW = (barAreaW - (bars.length - 1) * barGap) / bars.length;
  const maxBarH = 0.8;
  bars.forEach((pct, bi) => {
    const bx = x + 0.25 + bi * (barW + barGap);
    const bh = maxBarH * pct;
    const by = panelY + panelH - 0.25 - bh;
    slide.addShape('roundRect', {
      x: bx, y: by, w: barW, h: bh,
      fill: { color: m.trendColor }, opacity: 0.3 + bi * 0.3,
      rectRadius: 0.04,
    });
  });
});
```
