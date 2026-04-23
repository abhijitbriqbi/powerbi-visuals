# Briqlab Power BI Visuals — AppSource Submission Guide
> Complete reference for all 23 visuals. Fill this in Partner Center for each offer.
> Last updated: April 2026

---

## ⚡ Quick Reference — Common Fields (Same for All 23 Visuals)

| Field | Value |
|---|---|
| Publisher | Briqlab |
| Publisher contact email | abhijit@briqlab.io |
| Support URL | https://briqlab.io/support |
| Privacy Policy URL | https://briqlab.io/privacy-policy |
| Website URL | https://briqlab.io |
| Category | Analytics |
| Industry | All industries |
| Legal terms | https://briqlab.io/terms |
| API version | 5.3.0 |

---

## 📋 Part 1 — Privacy Policy Page (100.6.1)

### What to create at `https://briqlab.io/privacy-policy`

The page must explicitly mention the Power BI visuals. Use this text:

---

**Briqlab Privacy Policy**

*Last updated: April 2026*

**1. Data We Collect**
Briqlab Power BI custom visuals ("Visuals") run entirely inside Microsoft Power BI. The Visuals process only the data you bind to them within your Power BI report. We do not collect, store, or transmit any of your report data to our servers.

**2. Pro Key Validation**
When you enter a Briqlab Pro licence key in the visual format pane, the Visual makes a single HTTPS request to `https://api.briqlab.io/custom-visuals/validate/licence?key=YOUR_KEY` to verify the key. No personal data or report data is included in this request. The key itself is stored locally in your browser's localStorage (`briqlab_*_prokey`) so you do not need to re-enter it after a page refresh.

**3. Trial Period**
A 4-day free trial start date is stored locally in your browser's localStorage (`briqlab_trial_*_start`). This data never leaves your device.

**4. Visuals Covered by This Policy**
- Briqlab Animated Counter
- Briqlab Bar Chart
- Briqlab Bullet Pro
- Briqlab Calendar Heatmap
- Briqlab Donut Chart
- Briqlab Dot Matrix
- Briqlab Drill Down Bubble
- Briqlab Drill Down Gauge
- Briqlab Drill Down Pie
- Briqlab Sankey Chart
- Briqlab Gauge
- Briqlab KPI Card
- Briqlab KPI Sparkline
- Briqlab Mekko Chart
- Briqlab Pie Chart
- Briqlab Progress Ring
- Briqlab Pulse KPI
- Briqlab Radar Pro
- Briqlab Text Scroller
- Briqlab Search & Filter
- Briqlab Slope Chart
- Briqlab Violin Plot
- Briqlab Word Cloud

**5. Third-Party Services**
Visuals do not use Google Analytics, advertising networks, or any third-party tracking services.

**6. Children's Privacy**
Our Visuals are not directed to children under 13.

**7. Contact**
If you have questions about this policy, contact us at: privacy@briqlab.io

---

### Action needed:
1. Create this page at `https://briqlab.io/privacy-policy`
2. In Partner Center → each offer → Properties → **Privacy policy URL** → enter `https://briqlab.io/privacy-policy`

---

## 📸 Part 2 — Screenshots & Thumbnails (Per Visual)

### Specifications
| Asset | Size | Format | Where used |
|---|---|---|---|
| `assets/screenshot1.png` | 620 × 350 px | PNG | AppSource listing main image |
| `assets/thumbnail.png` | 220 × 176 px | PNG | Power BI Desktop visual picker |

### How to capture
1. Open Power BI Desktop
2. Add each visual with sample data
3. Resize the visual to approximately 620×350 in the report canvas
4. Use Windows Snipping Tool (Win+Shift+S) → exact selection
5. Save as PNG to `C:\Users\Abhijit\BriqlabVisuals\<VISUAL>\assets\screenshot1.png`
6. Resize to 220×176 and save as `thumbnail.png`
7. After saving both files, run: `cd C:\Users\Abhijit\BriqlabVisuals\<VISUAL> && npx pbiviz package`

### Screenshot content guide
| Visual | What to show |
|---|---|
| Animated Counter | Count-up animation frame mid-animation with large number, progress bar |
| Bar Chart | Horizontal grouped bars, comparison values, data labels visible |
| Bullet Pro | 3–5 bullet bars with target markers, colour zones visible |
| Calendar Heatmap | Full year view with colour gradient cells, at least 6 months with data |
| Donut Chart | 5–6 segments, centre label showing total, legend visible |
| Dot Matrix | 80% filled grid, 2 categories with different colours |
| Drill Down Bubble | Packed bubbles with 2+ colour groups, bubble size variation |
| Drill Down Gauge | Gauge at ~75% with target needle, teal fill zone |
| Drill Down Pie | Pie with 4 segments, one partially transparent (showing drill capability) |
| Sankey Chart | 3-node flow with proportional link widths |
| Gauge | Speedometer style at ~65%, coloured threshold zones visible |
| KPI Card | Large metric value, comparison delta (▲ green), label below |
| KPI Sparkline | KPI value top, sparkline trend below, target progress bar |
| Mekko Chart | 3–4 variable-width stacked columns, segment labels |
| Pie Chart | 5–6 segments, percentage labels, legend on right |
| Progress Ring | 3 concentric rings at different completion percentages |
| Pulse KPI | KPI value with pulsing green dot, trend line below |
| Text Scroller | Dark background with scrolling white text/metrics |
| Search & Filter | Search box with results list, one item highlighted |
| Slope Chart | 5–6 entities with slope lines, rank changes visible |
| Violin Plot | 2–3 violin shapes with box overlay, outlier dots |
| Word Cloud | 30+ words in varying sizes and colours |

---

## 📝 Part 3 — Partner Center Listing Details (All 23 Visuals)

### HOW TO FILL IN PARTNER CENTER
For each offer: `Marketplace offers → Power BI visual → <offer name> → Store listing`

---

### 1. Briqlab Animated Counter

| Field | Content |
|---|---|
| **Offer name** | Briqlab Animated Counter |
| **Short description** (100 chars) | Animated KPI counter with smooth count-up, progress bars, and customisable styles for Power BI. |
| **Long description** | Briqlab Animated Counter brings your KPIs to life with smooth count-up animations. Connect any numeric measure and watch it count from zero to its final value on every report refresh — instantly grabbing attention in executive dashboards, digital signage, and live operations screens. Format the value with currency prefixes, K/M suffixes, and decimal precision. Pair it with a configurable progress bar to show performance against target. Choose fonts, colors, and background to match your brand. A 4-day free trial is included — activate Briqlab Pro for unlimited use. Visit briqlab.io for details. |
| **Keywords** | animated counter, KPI, count-up, ticker, dashboard, metric, number animation |
| **Help link** | https://briqlab.io/docs/animated-counter |
| **Privacy policy** | https://briqlab.io/privacy-policy |
| **Screenshots** | screenshot1.png (620×350) |
| **Thumbnail** | thumbnail.png (220×176) |

---

### 2. Briqlab Bar Chart

| Field | Content |
|---|---|
| **Offer name** | Briqlab Bar Chart |
| **Short description** | Professional interactive Bar Chart for Power BI with comparison values, labels, and custom colors. |
| **Long description** | Briqlab Bar Chart is a polished, production-ready bar chart for Power BI that goes beyond the built-in visuals. Display single or grouped bars, toggle between vertical and horizontal orientation in one click, and add a comparison series to instantly see actuals vs. targets side by side. Data labels, axis titles, gridlines, corner radius, bar padding, and font settings are all configurable from the format pane — no DAX tricks needed. Cross-filtering and cross-highlighting work out of the box. Right-click any bar for the full Power BI context menu. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | bar chart, column chart, grouped bars, comparison, data labels, interactive |
| **Help link** | https://briqlab.io/docs/bar-chart |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 3. Briqlab Bullet Pro

| Field | Content |
|---|---|
| **Offer name** | Briqlab Bullet Pro |
| **Short description** | Premium bullet chart for Power BI — compare actuals vs. targets with colour-coded performance zones. |
| **Long description** | Briqlab Bullet Pro is a Stephen Few-inspired bullet chart that makes performance tracking intuitive. Plot actual values against targets with a clear marker line, and surround them with three colour-coded qualitative ranges (e.g. unsatisfactory / satisfactory / good). The compact horizontal layout lets you stack dozens of KPIs without scrolling. Configure zone colours, comparative marker style, animated bar fill, and axis scale from the format pane. Ideal for balanced scorecards, HR dashboards, financial reporting, and any scenario where you need to show "how did we do against the goal?" A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | bullet chart, KPI, target, performance, scorecard, progress, benchmark |
| **Help link** | https://briqlab.io/docs/bullet-pro |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 4. Briqlab Calendar Heatmap

| Field | Content |
|---|---|
| **Offer name** | Briqlab Calendar Heatmap |
| **Short description** | Daily activity heatmap calendar for Power BI — visualise patterns across weeks, months, and years. |
| **Long description** | Briqlab Calendar Heatmap renders any daily numeric measure as a GitHub-style activity calendar. Instantly spot seasonal patterns, weekly cycles, and outlier days that vanish in line or bar charts. Map a date column and a numeric measure — the visual automatically arranges cells by week and month, applies a colour gradient from low to high, and shows month/weekday labels. Configure the low and high gradient colours to match your brand or report theme. Hover any cell for an exact date and value tooltip. Perfect for sales velocity, web traffic, customer contacts, energy usage, and social media engagement analysis. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | calendar heatmap, daily activity, date analysis, seasonal patterns, GitHub style |
| **Help link** | https://briqlab.io/docs/calendar-heatmap |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 5. Briqlab Donut Chart

| Field | Content |
|---|---|
| **Offer name** | Briqlab Donut Chart |
| **Short description** | Professional animated Donut Chart for Power BI — centre label, legend, custom colours, data labels. |
| **Long description** | Briqlab Donut Chart combines elegant design with full Power BI interactivity. The adjustable inner radius lets you control whether it reads as a thick donut or a slim ring. The centre label automatically shows total or a custom metric. Up to 10 segment colours are fully configurable. Data labels display value, percentage, or category. The built-in legend supports top, bottom, left, right, and none positions. Click any segment to cross-filter the rest of the report; right-click for the Power BI context menu. Smooth entry animation draws the eye on first load. Ideal for part-to-whole analysis, budget allocation, market share, and portfolio breakdowns. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | donut chart, pie chart, ring chart, part-to-whole, animated, legend, centre label |
| **Help link** | https://briqlab.io/docs/donut-chart |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 6. Briqlab Dot Matrix

| Field | Content |
|---|---|
| **Offer name** | Briqlab Dot Matrix |
| **Short description** | Isotype / dot matrix chart for Power BI — show achieved vs. total with configurable dot shapes. |
| **Long description** | Briqlab Dot Matrix (also called an isotype or unit chart) makes progress intuitive by using discrete dots instead of abstract bar lengths. Each dot represents one unit of the whole — filled dots show what's been achieved, empty dots show what remains. Configure dot shape (circle, square, diamond, star), achieved colour, empty colour, dot size, and grid columns from the format pane. Optionally break down the filled portion by category to show composition within progress. The visual handles data reduction automatically for large counts. Great for HR headcount vs. vacancies, manufacturing units produced vs. planned, and campaign goal tracking. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | dot matrix, isotype, unit chart, progress, waffle chart, goal tracking |
| **Help link** | https://briqlab.io/docs/dot-matrix |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 7. Briqlab Drill Down Bubble

| Field | Content |
|---|---|
| **Offer name** | Briqlab Drill Down Bubble |
| **Short description** | Interactive drill-down bubble chart for Power BI — 3-dimensional packed layout with cross-filtering. |
| **Long description** | Briqlab Drill Down Bubble brings three-dimensional data analysis to Power BI without the clutter of traditional scatter plots. Map a category, X measure, Y measure, and bubble size to explore the relationship between three variables simultaneously. The packed layout automatically prevents overlaps. Colour groups let you segment bubbles by a fourth dimension. Click any bubble to drill into its sub-categories, and double-click to drill back up. Cross-filtering updates all other visuals on the page. Custom tooltips display all four dimensions plus any additional tooltip fields you drag in. Animated transitions make hierarchical navigation smooth and intuitive. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | bubble chart, scatter plot, drill down, 3D analysis, packed circles, hierarchy |
| **Help link** | https://briqlab.io/docs/drill-bubble |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 8. Briqlab Drill Down Gauge

| Field | Content |
|---|---|
| **Offer name** | Briqlab Drill Down Gauge |
| **Short description** | Premium gauge and speedometer for Power BI — auto colour thresholds, target marker, smooth animations. |
| **Long description** | Briqlab Drill Down Gauge is a production-quality speedometer visual for Power BI that makes performance status unmistakable at a glance. Map a value, optional minimum, maximum, and target to get an animated arc gauge with automatic colour thresholds (red / amber / teal / green) or a single manual colour. The target needle precisely marks the goal line. The comparison delta below the value shows how far above or below target you are in absolute and percentage terms. Configure arc thickness, value font size, and whether to display min/max labels. Smooth 600ms arc animation plays on every data refresh. Ideal for revenue attainment, OEE, NPS, and any single-metric performance view. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | gauge, speedometer, KPI, target, threshold, performance, meter |
| **Help link** | https://briqlab.io/docs/drill-gauge |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 9. Briqlab Drill Down Pie

| Field | Content |
|---|---|
| **Offer name** | Briqlab Drill Down Pie |
| **Short description** | Drill-down Pie & Donut for Power BI — navigate hierarchies with animated segments and breadcrumb. |
| **Long description** | Briqlab Drill Down Pie turns static pie charts into interactive data exploration tools. Drag in a hierarchy of categories — country, region, city for example — and click any segment to instantly drill into the next level. Animated segment transitions and a multi-level breadcrumb always show where you are in the hierarchy. A single toggle switches between pie and donut mode. Up to 10 segment colours are fully configurable. Data labels show value, percentage, or category name. Cross-filtering keeps the rest of your report in sync as you drill. Right-click any segment for the Power BI context menu. Custom tooltips support extra fields. Used by sales, marketing, finance, and operations teams for hierarchical part-to-whole analysis. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | drill down pie, hierarchical pie, donut chart, drill-through, animated, breadcrumb |
| **Help link** | https://briqlab.io/docs/drill-pie |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 10. Briqlab Sankey Chart

| Field | Content |
|---|---|
| **Offer name** | Briqlab Sankey Chart |
| **Short description** | Premium Sankey flow diagram for Power BI — proportional flows between nodes with animated links. |
| **Long description** | Briqlab Sankey Chart visualises value flows between nodes with width-proportional links that make the magnitude of each flow immediately apparent. Connect any two category columns as Source and Destination and a numeric measure as the Flow Value. The auto-layout engine positions nodes and routes links without overlaps. Flow opacity, node width, node gap, label threshold, and font are all configurable. Hover any link for a tooltip showing exact source, destination, and value. Ideal for budget allocation (department → cost centre), customer journey analysis (channel → channel), energy flows, supply chain mapping, and process analysis. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | Sankey chart, flow diagram, chord diagram, value flow, budget allocation, supply chain |
| **Help link** | https://briqlab.io/docs/sankey |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 11. Briqlab Gauge

| Field | Content |
|---|---|
| **Offer name** | Briqlab Gauge |
| **Short description** | Professional Gauge & Speedometer for Power BI — colour zones, target needle, configurable thresholds. |
| **Long description** | Briqlab Gauge is a clean, configurable gauge visual for Power BI that displays a single measure against configurable minimum, maximum, and target values. Three colour zones (red / amber / green) are individually configurable with custom thresholds so the gauge immediately communicates whether performance is poor, acceptable, or excellent. The target needle marks the goal exactly. Choose between automatic zone colouring and a single manual colour. Configure value font size, arc track colour, and whether to show min/max labels. The gauge updates smoothly on every data refresh. Works for revenue attainment, satisfaction scores, inventory levels, utilisation rates, and any single KPI requiring a visual "how are we doing?" answer. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | gauge, speedometer, KPI gauge, threshold, colour zones, performance indicator |
| **Help link** | https://briqlab.io/docs/gauge |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 12. Briqlab KPI Card

| Field | Content |
|---|---|
| **Offer name** | Briqlab KPI Card |
| **Short description** | Professional KPI Card for Power BI — primary metric, comparison delta, trend arrow, and custom styling. |
| **Long description** | Briqlab KPI Card is a beautifully designed metric card visual for Power BI that displays a primary KPI value, a comparison metric, a trend indicator arrow (▲/▼) with colour-coded performance, and a text label — all in a single, compact tile. The value auto-formats to K or M for large numbers. Font size, primary colour, background colour, border, and corner radius are fully configurable so the card matches any report theme. Animated pop effect highlights when the value changes. Ideal for executive summary pages, financial dashboards, sales scorecards, and any report where a single headline number must stand out. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | KPI card, metric tile, scorecard, comparison, trend, dashboard tile |
| **Help link** | https://briqlab.io/docs/kpi-card |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 13. Briqlab KPI Sparkline

| Field | Content |
|---|---|
| **Offer name** | Briqlab KPI Sparkline |
| **Short description** | Advanced KPI card with embedded sparkline trend, target progress bar, and animated count-up for Power BI. |
| **Long description** | Briqlab KPI Sparkline combines a headline KPI metric, a historical sparkline trend line, a target progress bar, and a comparison delta into one compact card — giving you four layers of context in the space of a single visual. The sparkline auto-colours based on trend direction. The target progress bar fills from left to right with a colour that changes based on attainment (above/below target). The count-up animation draws attention on report load. Configure sparkline type (line or area), height, colour, accent colour, background, corner radius, and shadow from the format pane. Built for finance, sales, and operations dashboards where trend context is as important as the current value. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | KPI sparkline, trend card, sparkline, target bar, count-up, metric card |
| **Help link** | https://briqlab.io/docs/kpi-sparkline |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 14. Briqlab Mekko Chart

| Field | Content |
|---|---|
| **Offer name** | Briqlab Mekko Chart |
| **Short description** | Premium Mekko / Marimekko chart for Power BI — show market size and market share simultaneously. |
| **Long description** | Briqlab Mekko Chart (also called a Marimekko or mosaic chart) encodes two dimensions in a single chart: column width represents one measure (e.g. total market size) and column height represents another (e.g. market share). This makes it the go-to chart for competitive analysis, market landscape reviews, and portfolio assessments where both the size and the share of each category matter. Segment labels display category names and values inside each cell. Configure segment colours, label font size, and axis formatting from the format pane. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | Mekko chart, Marimekko, mosaic chart, market share, market size, competitive analysis |
| **Help link** | https://briqlab.io/docs/mekko-chart |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 15. Briqlab Pie Chart

| Field | Content |
|---|---|
| **Offer name** | Briqlab Pie Chart |
| **Short description** | Professional animated Pie Chart for Power BI — data labels, legend, custom colours, border styling. |
| **Long description** | Briqlab Pie Chart is a polished pie chart that adds professional data labels, a configurable legend, and custom segment colours to Power BI. Choose to show values, percentages, or category names as labels. Position the legend at top, bottom, left, or right. Configure up to 10 individual segment colours to match your brand. A white segment border with configurable width separates the slices cleanly. The total is optionally displayed in the chart title area. Click any segment to cross-filter the rest of the report; right-click for the Power BI context menu. Smooth entry animation plays on first data load. Use it for part-to-whole analysis, survey results, category breakdowns, and budget allocation. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | pie chart, part-to-whole, segment, data labels, legend, percentage |
| **Help link** | https://briqlab.io/docs/pie-chart |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 16. Briqlab Progress Ring

| Field | Content |
|---|---|
| **Offer name** | Briqlab Progress Ring |
| **Short description** | Multi-ring progress chart for Power BI — up to 6 concurrent goals as concentric animated rings. |
| **Long description** | Briqlab Progress Ring displays up to six goals simultaneously as concentric rings, making it easy to compare performance across multiple metrics at once. Each ring fills proportionally to its percentage complete with a smooth animated arc. Rings auto-colour using the Briqlab brand palette, or set individual colours manually. A centre summary optionally shows the average completion percentage. Milestone markers at configurable thresholds (e.g. 50%, 80%, 100%) provide visual reference points. Labels outside each ring show category name and value. Ideal for OKR tracking, multi-metric scorecards, departmental goal dashboards, and health-and-safety compliance monitoring. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | progress ring, donut gauge, circular progress, OKR, goal tracking, multi-metric |
| **Help link** | https://briqlab.io/docs/progress-ring |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 17. Briqlab Pulse KPI

| Field | Content |
|---|---|
| **Offer name** | Briqlab Pulse KPI |
| **Short description** | Animated KPI card with pulsing status indicator, trend line, and comparison delta for Power BI. |
| **Long description** | Briqlab Pulse KPI adds a living heartbeat to your dashboards. The pulsing status dot (green when above target, red when below) immediately communicates whether a KPI is healthy or needs attention — even from across the room. The card shows the current value, comparison delta with directional arrow, a mini trend indicator, and a configurable label. The pulse animation speed and colour automatically respond to performance status. Configure primary colour, background colour, and font from the format pane. Great for real-time operations dashboards, sales floor displays, IT monitoring, and any scenario where status at a glance matters more than detail. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | pulse KPI, animated KPI, status indicator, heartbeat, real-time dashboard, alert |
| **Help link** | https://briqlab.io/docs/pulse-kpi |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 18. Briqlab Radar Pro

| Field | Content |
|---|---|
| **Offer name** | Briqlab Radar Pro |
| **Short description** | Premium Radar / Spider chart for Power BI — compare multiple entities across shared dimensions. |
| **Long description** | Briqlab Radar Pro is a spider / radar chart that makes multi-dimensional comparison intuitive. Plot two or more entities (products, teams, time periods, competitors) across a shared set of dimensions (criteria) and the filled polygons immediately reveal which entity is strongest across which dimensions. A benchmark overlay ring highlights a target level across all axes. A scorecard panel alongside the chart ranks entities by weighted score. Configure polygon fill opacity, axis labels, gridline count, and font. Ideal for balanced scorecard analysis, competitive benchmarking, skills assessment, product feature comparison, and multi-criteria evaluation. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | radar chart, spider chart, multi-dimensional, benchmark, comparison, scorecard |
| **Help link** | https://briqlab.io/docs/radar-pro |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 19. Briqlab Text Scroller

| Field | Content |
|---|---|
| **Offer name** | Briqlab Text Scroller |
| **Short description** | Animated scrolling news-ticker for Power BI — display live metrics and text as a smooth horizontal scroll. |
| **Long description** | Briqlab Text Scroller turns any text or numeric column into a smooth horizontally-scrolling ticker tape — perfect for live metric feeds, announcements, and digital signage. Connect a text or numeric measure and watch it scroll continuously across the visual at your configured speed. Support for positive/negative colouring automatically colours up/down metrics in green/red. Configure scroll speed, font size, font family, text colour, background colour, and separator character. The scroller loops seamlessly with no jumps. Use it on dashboards displayed on large screens in retail floors, operations centres, trading desks, event venues, and reception areas. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | text scroller, news ticker, scrolling text, live dashboard, digital signage, ticker tape |
| **Help link** | https://briqlab.io/docs/scroller |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 20. Briqlab Search & Filter

| Field | Content |
|---|---|
| **Offer name** | Briqlab Search & Filter |
| **Short description** | Type-to-filter cross-filter visual for Power BI — instantly search any text column and update all visuals. |
| **Long description** | Briqlab Search & Filter adds a live search box to any Power BI report. As you type, the visual instantly filters a scrollable list of matching category values and cross-filters all other visuals on the page — no slicers, no dropdowns. The search is case-insensitive and matches anywhere in the value (contains search). Click any result to pin a single selection; clear the box to reset. Configure accent colour, background, border, and text colour from the format pane. Ideal for reports with hundreds or thousands of product names, customer names, employee IDs, or any other text dimension where a slicer list would be too long to scroll. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | search filter, type-ahead, cross-filter, slicer, search box, text filter |
| **Help link** | https://briqlab.io/docs/search-filter |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 21. Briqlab Slope Chart

| Field | Content |
|---|---|
| **Offer name** | Briqlab Slope Chart |
| **Short description** | Premium Slope Chart for Power BI — compare two-period rankings with animated lines and rank changes. |
| **Long description** | Briqlab Slope Chart is the clearest way to show how rankings or values changed between two time periods. Each entity is a line from the left period to the right period — rising lines are instantly distinguishable from falling lines, and the slope angle communicates the magnitude of change. Rank-change badges (▲3, ▼2) appear next to each label. Colour by direction (green/red) or by category (brand palette). A summary bar at the bottom shows what percentage of entities improved. Label collision detection prevents overlapping text even with many entities. Used in sales territory analysis, product ranking shifts, employee performance review, and before/after comparisons. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | slope chart, bump chart, ranking change, before-after, two-period comparison |
| **Help link** | https://briqlab.io/docs/slope-chart |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 22. Briqlab Violin Plot

| Field | Content |
|---|---|
| **Offer name** | Briqlab Violin Plot |
| **Short description** | Statistical distribution visual for Power BI — violin curves, box plot overlay, outliers, reference lines. |
| **Long description** | Briqlab Violin Plot reveals the full distribution of your data — not just the average. The mirrored KDE (kernel density estimate) curves show where data points are concentrated. The optional box plot overlay adds median, quartiles, and whiskers. Individual outlier dots mark extreme values. A configurable reference line marks a benchmark or target. Compare distributions across multiple categories side by side. This is the visual to use when you suspect your data is bimodal, skewed, or has outliers that averages hide. Ideal for quality control, salary band analysis, customer satisfaction score distribution, delivery time analysis, and any statistical review. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | violin plot, distribution, KDE, box plot, outliers, statistics, data analysis |
| **Help link** | https://briqlab.io/docs/violin-plot |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

### 23. Briqlab Word Cloud

| Field | Content |
|---|---|
| **Offer name** | Briqlab Word Cloud |
| **Short description** | Premium Word Cloud for Power BI — weighted keywords with sentiment colouring and stop-word filtering. |
| **Long description** | Briqlab Word Cloud renders text data as a weighted keyword cloud where word size corresponds to frequency or a numeric weight measure. Choose from spiral, rectangular, or random placement layouts. Colour mode options include uniform brand colour, random from the palette, or sentiment (positive/negative). Built-in stop-word filtering removes common words (the, and, is) automatically, with the option to add custom stop words. Configurable rotation angle, min/max font size, and font family give you full layout control. Click any word to cross-filter the rest of the report. Ideal for social media analysis, customer feedback themes, survey text analysis, and product review mining. A 4-day free trial is included; activate Briqlab Pro for unlimited use. |
| **Keywords** | word cloud, text analysis, sentiment, keywords, frequency, NLP, text mining |
| **Help link** | https://briqlab.io/docs/word-cloud |
| **Privacy policy** | https://briqlab.io/privacy-policy |

---

## 📦 Part 4 — Sample .pbix File Hints (1180.2.3.1)

### What reviewers expect
Each visual's sample `.pbix` file should have a **text box** on the report page explaining how to use the visual. This is a "soft" requirement — it won't block certification but improves reviewer experience.

### Text to add in each sample .pbix (as a report text box):

**Generic template:**
```
How to use this visual:
1. Drag a [CATEGORY/DATE] field into the [FIELD NAME] bucket
2. Drag a numeric [MEASURE] into the [FIELD NAME] bucket
3. Use the Format pane to customise colours, fonts, and labels
4. Enter your Briqlab Pro key in Format › Briqlab Pro to activate
5. Visit briqlab.io/docs for full documentation
```

**Per-visual field guidance:**

| Visual | Field 1 | Field 2 | Field 3 |
|---|---|---|---|
| Animated Counter | — | Value (measure) | — |
| Bar Chart | Category (text) | Values (measure) | Comparison Values (optional) |
| Bullet Pro | Category (text) | Actual (measure) | Target (measure) |
| Calendar Heatmap | Date (date) | Value (measure) | — |
| Donut Chart | Category (text) | Values (measure) | — |
| Dot Matrix | — | Total (measure) | Achieved (measure) |
| Drill Down Bubble | Category (text) | X Axis (measure) | Y Axis + Bubble Size |
| Drill Down Gauge | — | Value (measure) | Target (optional) |
| Drill Down Pie | Category hierarchy | Values (measure) | — |
| Sankey Chart | Source (text) | Destination (text) | Flow Value (measure) |
| Gauge | — | Value (measure) | Min / Max / Target (optional) |
| KPI Card | Category (text, optional) | Measure (measure) | Comparison (optional) |
| KPI Sparkline | Trend Date (date) | KPI Value (measure) | Target + Trend Values |
| Mekko Chart | Row Category (text) | Column Category (text) | Values (measure) |
| Pie Chart | Category (text) | Values (measure) | — |
| Progress Ring | Category (text) | Actual (measure) | Target (measure) |
| Pulse KPI | Label (text, optional) | Value (measure) | Target (optional) |
| Radar Pro | Category (axis) | Entity (series) | Values (measure) |
| Text Scroller | — | Text/Value (measure) | — |
| Search & Filter | Category to search (text) | — | — |
| Slope Chart | Entity (text) | Period 1 (measure) | Period 2 (measure) |
| Violin Plot | Category (text) | Values (measure) | — |
| Word Cloud | Words (text) | Weight (measure, optional) | — |

---

## 🎨 Part 5 — Brand Colour Reference

All 23 visuals now use the unified Briqlab brand palette:

| Name | Hex | Use |
|---|---|---|
| **Primary Teal** | `#0D9488` | Primary bars, arcs, gauges, accent elements |
| **Orange** | `#F97316` | Comparison series, second data colour |
| **Blue** | `#3B82F6` | Third data colour |
| **Purple** | `#8B5CF6` | Fourth data colour |
| **Green** | `#10B981` | Positive / above target |
| **Red** | `#EF4444` | Negative / below target / alerts |
| **Amber** | `#F59E0B` | Warning / medium performance |
| **Pink** | `#EC4899` | Eighth data colour |
| **Cyan** | `#06B6D4` | Ninth data colour |
| **Lime** | `#84CC16` | Tenth data colour |
| **Text Dark** | `#374151` | Labels, axis text, annotations |
| **Border/Track** | `#E5E7EB` | Grid lines, track backgrounds |
| **Card Border** | `#E2E8F0` | Card borders, empty/neutral |
| **White** | `#FFFFFF` | Card backgrounds |

---

## ✅ Part 6 — Submission Checklist

### Code / Package (completed)
- [x] All 23 .pbiviz packages built without errors
- [x] Context menu working (right-click shows Power BI menu)
- [x] Tooltip service implemented (host.tooltipService)
- [x] Per-element contextmenu on BarChart, DrillPie, PieChart, DonutChart, WordCloud, SlopeChart
- [x] capabilities.json: `tooltips`, `supportsLandingPage`, `supportsEmptyDataView` added to all 23
- [x] capabilities.json: `privileges` (WebAccess) for api.briqlab.io on all 23
- [x] pbiviz.json: `supportUrl`, `websiteUrl`, `author.email` set on all 23
- [x] Brand colour palette unified across all 23 visuals
- [x] All icons unique and present (assets/icon.png)
- [x] DrillGauge blank output fixed (CSS fill override removed)
- [x] DonutChart pro key erasure bug fixed

### Partner Center (action needed)
- [ ] Create `https://briqlab.io/privacy-policy` page (use text from Part 1)
- [ ] Create `https://briqlab.io/terms` page
- [ ] Create `https://briqlab.io/support` page or redirect
- [ ] Create `https://briqlab.io/docs` pages (one per visual, optional but recommended)
- [ ] Upload screenshots (620×350px) for each offer — see Part 2
- [ ] Upload thumbnails (220×176px) for each offer — see Part 2
- [ ] Fill in offer listing for each of 23 visuals — use descriptions from Part 3
- [ ] Set Privacy Policy URL to `https://briqlab.io/privacy-policy` on all 23 offers
- [ ] Upload sample .pbix file for each visual with How-to text box — see Part 4
- [ ] Submit all 23 for review

---

*Generated by Briqlab development tools — April 2026*
