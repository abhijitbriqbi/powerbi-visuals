"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import * as d3 from "d3";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions      = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual                  = powerbi.extensibility.visual.IVisual;
import IVisualHost              = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager        = powerbi.extensibility.ISelectionManager;
import ISelectionId             = powerbi.visuals.ISelectionId;

import { VisualFormattingSettingsModel } from "./settings";
import { checkMicrosoftLicence, resetLicenceCache } from "./licenceManager";
import { getTrialDaysRemaining, isTrialExpired, getPurchaseUrl, getButtonText } from "./trialManager";

// ── Types ──────────────────────────────────────────────────────────────────────
interface SparkPoint { date: string; value: number; }

interface KpiData {
    kpiValue:        number | null;
    targetValue:     number | null;
    comparisonValue: number | null;
    sparkPoints:     SparkPoint[];
    kpiLabel:        string;
    selectionId:     ISelectionId | null;
}

// ── Constants ──────────────────────────────────────────────────────────────────
const TRIAL_KEY     = "briqlab_trial_kpisparkline_start";
const PRO_STORE_KEY = "briqlab_key_kpisparkline";
const TRIAL_MS      = 4 * 24 * 60 * 60 * 1000;

// ── Helpers ────────────────────────────────────────────────────────────────────
function fmtVal(v: number): string {
    const a = Math.abs(v);
    if (a >= 1e9) { const r = (v / 1e9).toFixed(2); return (r.endsWith(".00") ? r.slice(0, -3) : r.endsWith("0") ? r.slice(0, -1) : r) + "B"; }
    if (a >= 1e6) { const r = (v / 1e6).toFixed(2); return (r.endsWith(".00") ? r.slice(0, -3) : r.endsWith("0") ? r.slice(0, -1) : r) + "M"; }
    if (a >= 1e3) { const r = (v / 1e3).toFixed(1); return (r.endsWith(".0") ? r.slice(0, -2) : r) + "K"; }
    return v.toLocaleString("en-US");
}

function dur(base: number): number {
    try { return window.matchMedia("(prefers-reduced-motion: reduce)").matches ? 0 : base; }
    catch { return base; }
}

function clamp(v: number, lo: number, hi: number): number {
    return Math.max(lo, Math.min(hi, v));
}

// ── Visual ─────────────────────────────────────────────────────────────────────
export class Visual implements IVisual {

    private readonly host:   IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private readonly selMgr: ISelectionManager;
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private _handlersAttached = false;
    private readonly fmtSvc: FormattingSettingsService;
    private readonly root:   HTMLElement;

    // DOM structure
    private readonly contentEl:      HTMLDivElement;
    private readonly cardShell:      HTMLDivElement;
    private readonly accentStripe:   HTMLDivElement;
    private readonly cardBody:       HTMLDivElement;
    private readonly headerRow:      HTMLDivElement;
    private readonly labelEl:        HTMLDivElement;
    private readonly trendBadge:     HTMLDivElement;
    private readonly trendArrow:     HTMLSpanElement;
    private readonly trendText:      HTMLSpanElement;
    private readonly valueArea:      HTMLDivElement;
    private readonly valueEl:        HTMLDivElement;
    private readonly targetSection:  HTMLDivElement;
    private readonly targetLabelRow: HTMLDivElement;
    private readonly targetLabelTxt: HTMLSpanElement;
    private readonly targetPct:      HTMLSpanElement;
    private readonly targetTrack:    HTMLDivElement;
    private readonly targetFill:     HTMLDivElement;
    private readonly sparkArea:      HTMLDivElement;
    private readonly sparkTooltip:   HTMLDivElement;
    private readonly trialBadge:     HTMLDivElement;
    private readonly proBadge:       HTMLDivElement;
    private readonly keyErrorEl:     HTMLDivElement;
    private readonly overlayEl:      HTMLDivElement;

    // State
    private settings!:    VisualFormattingSettingsModel;
    private viewport:     powerbi.IViewport = { width: 300, height: 200 };
    private data:         KpiData           = { kpiValue: null, targetValue: null, comparisonValue: null, sparkPoints: [], kpiLabel: "", selectionId: null };
    private prevKpiValue: number | null     = null;

    // Trial / Pro
    private isPro    = false;
    private lastKey  = "";
    private readonly keyCache: Map<string, boolean> = new Map();

    // ── Constructor ────────────────────────────────────────────────────────────
    constructor(options: VisualConstructorOptions) {
        this.host   = options.host;
        this.renderingManager = options.host.eventService;
        this.selMgr = options.host.createSelectionManager();
        this.tooltipSvc = this.host.tooltipService;
        this.fmtSvc = new FormattingSettingsService();
        this.root   = options.element;
        this.root.classList.add("briqlab-kpi-sparkline");

        // Content wrapper
        this.contentEl = this.mkDiv("briqlab-visual-content");
        this.root.appendChild(this.contentEl);

        // Card shell
        this.cardShell = this.mkDiv("kpi-card-shell");
        this.contentEl.appendChild(this.cardShell);

        // Accent stripe
        this.accentStripe = this.mkDiv("kpi-accent-stripe");
        this.cardShell.appendChild(this.accentStripe);

        // Card body
        this.cardBody = this.mkDiv("kpi-card-body");
        this.cardShell.appendChild(this.cardBody);

        // Header row
        this.headerRow = this.mkDiv("kpi-header-row");
        this.cardBody.appendChild(this.headerRow);

        this.labelEl = this.mkDiv("kpi-label");
        this.headerRow.appendChild(this.labelEl);

        this.trendBadge = this.mkDiv("kpi-trend-badge");
        this.trendArrow = document.createElement("span");
        this.trendArrow.className = "trend-arrow";
        this.trendText = document.createElement("span");
        this.trendBadge.appendChild(this.trendArrow);
        this.trendBadge.appendChild(this.trendText);
        this.headerRow.appendChild(this.trendBadge);

        // Value area
        this.valueArea = this.mkDiv("kpi-value-area");
        this.cardBody.appendChild(this.valueArea);
        this.valueEl = this.mkDiv("kpi-value");
        this.valueArea.appendChild(this.valueEl);

        // Target section
        this.targetSection  = this.mkDiv("kpi-target-section");
        this.cardBody.appendChild(this.targetSection);
        this.targetLabelRow = this.mkDiv("kpi-target-label-row");
        this.targetSection.appendChild(this.targetLabelRow);
        this.targetLabelTxt = document.createElement("span");
        this.targetLabelTxt.className = "kpi-target-label-text";
        this.targetLabelRow.appendChild(this.targetLabelTxt);
        this.targetPct = document.createElement("span");
        this.targetPct.className = "kpi-target-pct";
        this.targetLabelRow.appendChild(this.targetPct);
        this.targetTrack = this.mkDiv("kpi-target-bar-track");
        this.targetSection.appendChild(this.targetTrack);
        this.targetFill = this.mkDiv("kpi-target-bar-fill");
        this.targetTrack.appendChild(this.targetFill);

        // Sparkline area
        this.sparkArea = this.mkDiv("kpi-sparkline-area");
        this.cardBody.appendChild(this.sparkArea);

        // Sparkline tooltip (outside card so it's never clipped)
        this.sparkTooltip = this.mkDiv("kpi-spark-tooltip");
        this.root.appendChild(this.sparkTooltip);

        // Badges
        this.trialBadge = this.mkDiv("briqlab-trial-badge hidden");
        this.root.appendChild(this.trialBadge);
        this.proBadge = this.mkDiv("briqlab-pro-badge hidden");
        this.root.appendChild(this.proBadge);
        this.keyErrorEl = this.mkDiv("briqlab-key-error hidden");
        this.keyErrorEl.textContent = "\u2717 Invalid key";
        this.root.appendChild(this.keyErrorEl);

        // Trial overlay
        this.overlayEl = this.mkDiv("briqlab-trial-overlay hidden");
        this.buildTrialOverlay();
        this.root.appendChild(this.overlayEl);

        // Cross-filter on card click
        this.cardShell.addEventListener("click", () => this.onCardClick());

        this.restoreProKey();
    }

    // ── Trial overlay (no innerHTML) ───────────────────────────────────────────
    private buildTrialOverlay(): void {
        const card = this.mkDiv("briqlab-trial-card");

        const title = document.createElement("p");
        title.className = "trial-title";
        title.textContent = "Free trial ended";
        card.appendChild(title);

        const body = document.createElement("p");
        body.className = "trial-body";
        body.textContent = "Activate Briqlab Pro to continue using this visual and unlock all features.";
        card.appendChild(body);

        const btn = document.createElement("button");
        btn.className = "trial-btn";
        btn.textContent = getButtonText();
        btn.addEventListener("click", () => this.host.launchUrl(getPurchaseUrl()));
        card.appendChild(btn);

        const sub = document.createElement("p");
        sub.className = "trial-subtext";
        sub.textContent = "Purchase on Microsoft AppSource to unlock all features instantly.";
        card.appendChild(sub);

        this.overlayEl.appendChild(card);
    }

    // ── Lifecycle: update() ────────────────────────────────────────────────────
    public update(options: VisualUpdateOptions): void {
        this.renderingManager.renderingStarted(options);
        try {
            // AppSource: attach context-menu + tooltip once
            if (!this._handlersAttached) {
                this._handlersAttached = true;
                this.root.addEventListener("contextmenu", (e: MouseEvent) => {
                    e.preventDefault();
                    this.selMgr.showContextMenu(
                        null as unknown as powerbi.visuals.ISelectionId,
                        { x: e.clientX, y: e.clientY }
                    );
                });
                this.root.addEventListener("mousemove", (e: MouseEvent) => {
                    this.tooltipSvc.show({
                        dataItems: [{ displayName: "Briqlab KPI Sparkline", value: "" }],
                        identities: [],
                        coordinates: [e.clientX, e.clientY],
                        isTouchEvent: false
                    });
                });
                this.root.addEventListener("mouseleave", () => {
                    this.tooltipSvc.hide({ isTouchEvent: false, immediately: false });
                });
            }
            try {
                this.viewport = options.viewport;
                this.settings = this.fmtSvc.populateFormattingSettingsModel(
                    VisualFormattingSettingsModel,
                    options.dataViews?.[0]
                );
    
                this.handleProKey();
                this.updateLicenseUI();
    
                const dv      = options.dataViews?.[0];
                const parsed  = this.parseData(dv);
                this.prevKpiValue = this.data.kpiValue;
                this.data         = parsed;
    
                this.render();
    
            } catch (err) {
                console.error("[BriqlabKPISparkline] update error:", err);
            }
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    // ── Data parsing ───────────────────────────────────────────────────────────
    private parseData(dv: powerbi.DataView | undefined): KpiData {
        const empty: KpiData = { kpiValue: null, targetValue: null, comparisonValue: null, sparkPoints: [], kpiLabel: "", selectionId: null };
        if (!dv?.categorical?.values) { return empty; }

        const cat  = dv.categorical;
        const cats = cat.categories;
        const vals = cat.values;

        let kpiCol:        powerbi.DataViewValueColumn | null = null;
        let sparkCol:      powerbi.DataViewValueColumn | null = null;
        let targetCol:     powerbi.DataViewValueColumn | null = null;
        let comparisonCol: powerbi.DataViewValueColumn | null = null;

        for (let i = 0; i < vals.length; i++) {
            const col   = vals[i];
            const roles = col.source.roles ?? {};
            if (roles["kpiValue"]        && !kpiCol)        { kpiCol        = col; continue; }
            if (roles["sparklineValues"] && !sparkCol)      { sparkCol      = col; continue; }
            if (roles["targetValue"]     && !targetCol)     { targetCol     = col; continue; }
            if (roles["comparisonValue"] && !comparisonCol) { comparisonCol = col; continue; }
        }

        const dateCol  = cats && cats.length > 0 ? cats[0] : null;
        const rowCount = dateCol ? dateCol.values.length
            : (kpiCol ? kpiCol.values.length : (sparkCol ? sparkCol.values.length : 0));

        // Sparkline points
        const sparkPoints: SparkPoint[] = [];
        if (sparkCol && rowCount > 0) {
            for (let i = 0; i < rowCount; i++) {
                const sv = sparkCol.values[i] as number;
                if (sv === null || sv === undefined) { continue; }
                const date = dateCol ? String(dateCol.values[i] ?? "") : String(i + 1);
                sparkPoints.push({ date, value: sv });
            }
        }

        // KPI value: last non-null row of kpiCol, or last spark point
        let kpiValue: number | null = null;
        if (kpiCol && rowCount > 0) {
            for (let i = rowCount - 1; i >= 0; i--) {
                const v = kpiCol.values[i] as number;
                if (v !== null && v !== undefined) { kpiValue = v; break; }
            }
        } else if (sparkPoints.length > 0) {
            kpiValue = sparkPoints[sparkPoints.length - 1].value;
        }

        // Target: first non-null
        let targetValue: number | null = null;
        if (targetCol && rowCount > 0) {
            for (let i = 0; i < rowCount; i++) {
                const v = targetCol.values[i] as number;
                if (v !== null && v !== undefined) { targetValue = v; break; }
            }
        }

        // Comparison: first non-null
        let comparisonValue: number | null = null;
        if (comparisonCol && rowCount > 0) {
            for (let i = 0; i < rowCount; i++) {
                const v = comparisonCol.values[i] as number;
                if (v !== null && v !== undefined) { comparisonValue = v; break; }
            }
        }

        // KPI label from measure column name
        const kpiLabel = kpiCol
            ? (kpiCol.source.displayName || "")
            : (sparkCol ? (sparkCol.source.displayName || "") : "");

        // Selection ID (first date row)
        let selectionId: ISelectionId | null = null;
        if (dateCol && rowCount > 0) {
            try {
                selectionId = this.host.createSelectionIdBuilder()
                    .withCategory(dateCol, 0)
                    .createSelectionId();
            } catch { /* ignore */ }
        }

        return { kpiValue, targetValue, comparisonValue, sparkPoints, kpiLabel, selectionId };
    }

    // ── Render ─────────────────────────────────────────────────────────────────
    private render(): void {
        const { width, height } = this.viewport;

        // Card style
        const bgColor      = this.settings.cardStyle.backgroundColor.value?.value ?? "#FFFFFF";
        const borderRadius = clamp(this.settings.cardStyle.borderRadius.value, 4, 20);
        const showShadow   = this.settings.cardStyle.showShadow.value;

        this.cardShell.style.backgroundColor = bgColor;
        this.cardShell.style.borderRadius     = `${borderRadius}px`;
        this.cardShell.style.boxShadow        = showShadow ? "0 2px 8px rgba(0,0,0,0.06)" : "none";

        // Status colour
        const statusColor = this.resolveStatusColor();

        // Accent stripe
        const accentMode  = this.settings.cardStyle.accentColor.value.value as string;
        const accentColor = accentMode === "Manual"
            ? (this.settings.cardStyle.manualAccentColor.value?.value ?? "#0D9488")
            : statusColor;
        this.accentStripe.style.backgroundColor = accentColor;
        this.accentStripe.style.borderRadius     = `${borderRadius}px 0 0 ${borderRadius}px`;

        const valueSize = clamp(this.settings.valueSettings.valueSize.value, 16, 96);
        const prefixRaw = this.settings.valueSettings.currencyPrefix.value.value as string;
        const prefix    = prefixRaw === "None" ? "" : prefixRaw + "\u202F";

        const isTiny = width < 140 || height < 90;

        // Label
        this.labelEl.textContent = this.data.kpiLabel;

        // Trend
        this.renderTrend();

        // Value
        this.renderValue(valueSize, prefix, statusColor);

        // Target bar
        const showTarget = this.settings.targetSettings.showTargetBar.value
            && this.data.kpiValue !== null
            && this.data.targetValue !== null
            && !isTiny;
        this.renderTargetBar(statusColor, showTarget);

        // Sparkline
        const showSpark = this.settings.sparklineSettings.showSparkline.value
            && this.data.sparkPoints.length >= 2
            && !isTiny;
        if (showSpark) {
            this.renderSparkline(width, height, statusColor);
        } else {
            this.clearEl(this.sparkArea);
            this.sparkArea.style.display = "none";
        }
    }

    // ── Status color ──────────────────────────────────────────────────────────
    private resolveStatusColor(): string {
        const kpi    = this.data.kpiValue;
        const target = this.data.targetValue;
        if (kpi !== null && target !== null) {
            return kpi >= target
                ? (this.settings.targetSettings.aboveTargetColor.value?.value ?? "#10B981")
                : (this.settings.targetSettings.belowTargetColor.value?.value ?? "#EF4444");
        }
        return "#0D9488";
    }

    // ── Trend ─────────────────────────────────────────────────────────────────
    private renderTrend(): void {
        const kpi  = this.data.kpiValue;
        const comp = this.data.comparisonValue;

        if (!this.settings.valueSettings.showTrend.value || kpi === null || comp === null || comp === 0) {
            this.trendBadge.style.display = "none";
            return;
        }
        this.trendBadge.style.display = "";

        const pct      = ((kpi - comp) / Math.abs(comp)) * 100;
        const positive = pct >= 0;

        this.trendArrow.textContent = positive ? "\u25b2" : "\u25bc";
        this.trendText.textContent  = `${positive ? "+" : ""}${pct.toFixed(1)}%`;
        this.trendBadge.className   = `kpi-trend-badge ${positive ? "positive" : "negative"}`;

        // Bounce once on data change
        if (kpi !== this.prevKpiValue) {
            this.trendBadge.classList.remove("bounce");
            void this.trendBadge.offsetWidth;
            this.trendBadge.classList.add("bounce");
        }
    }

    // ── Value (count-up) ──────────────────────────────────────────────────────
    private renderValue(valueSize: number, prefix: string, statusColor: string): void {
        const kpi = this.data.kpiValue;

        this.valueEl.style.fontSize = `${valueSize}px`;
        this.valueEl.style.color    = statusColor;

        if (kpi === null) { this.valueEl.textContent = "\u2014"; return; }

        const doAnimate = this.settings.valueSettings.countUpAnimation.value
            && this.prevKpiValue !== kpi
            && dur(1) > 0;

        if (doAnimate) {
            const from   = this.prevKpiValue ?? 0;
            const interp = d3.interpolateNumber(from, kpi);

            d3.select(this.valueEl)
                .interrupt()
                .transition()
                .duration(dur(800))
                .ease(d3.easeCubicOut)
                .tween("text", () => (t: number) => {
                    this.valueEl.textContent = prefix + fmtVal(interp(t));
                });
        } else {
            d3.select(this.valueEl).interrupt();
            this.valueEl.textContent = prefix + fmtVal(kpi);
        }
    }

    // ── Target bar ────────────────────────────────────────────────────────────
    private renderTargetBar(statusColor: string, show: boolean): void {
        if (!show) { this.targetSection.style.display = "none"; return; }
        this.targetSection.style.display = "";

        const kpi    = this.data.kpiValue!;
        const target = this.data.targetValue!;
        const pct    = (kpi / target) * 100;
        const filled = clamp(pct, 0, 100);
        const label  = (this.settings.targetSettings.targetLabel.value || "vs Target").trim();

        this.targetLabelTxt.textContent       = `${label}: ${fmtVal(target)}`;
        this.targetPct.textContent            = `\u2713 ${pct.toFixed(0)}%`;
        this.targetPct.style.color            = statusColor;
        this.targetFill.style.width           = `${filled}%`;
        this.targetFill.style.backgroundColor = statusColor;
    }

    // ── Sparkline ─────────────────────────────────────────────────────────────
    private renderSparkline(chartW: number, chartH: number, statusColor: string): void {
        this.clearEl(this.sparkArea);
        this.sparkArea.style.display = "";

        const s         = this.settings.sparklineSettings;
        const heightPct = clamp(s.sparklineHeight.value, 15, 60) / 100;
        const spH       = Math.max(20, chartH * heightPct);
        const spW       = Math.max(10, chartW - 28); // subtract padding

        this.sparkArea.style.height = `${spH}px`;

        const colorMode = s.sparklineColor.value.value as string;
        const color     = colorMode === "Manual"
            ? (s.manualSparklineColor.value?.value ?? statusColor)
            : statusColor;

        const pts = this.data.sparkPoints;

        const svg = d3.select(this.sparkArea)
            .append<SVGSVGElement>("svg")
            .attr("width",  spW)
            .attr("height", spH);

        const xScale = d3.scaleLinear()
            .domain([0, pts.length - 1])
            .range([2, spW - 2]);

        const yMin = d3.min(pts, p => p.value) ?? 0;
        const yMax = d3.max(pts, p => p.value) ?? 1;
        const yPad = (yMax - yMin) * 0.1 || Math.abs(yMax) * 0.05 || 1;

        const yScale = d3.scaleLinear()
            .domain([yMin - yPad, yMax + yPad])
            .range([spH - 2, 2]);

        const type = s.sparklineType.value.value as string;

        if (type === "Bar") {
            const barW = Math.max(1, (spW - 4) / pts.length - 1);
            svg.selectAll<SVGRectElement, SparkPoint>(".sparkline-bar-rect")
                .data(pts)
                .enter().append("rect")
                .attr("class",  "sparkline-bar-rect")
                .attr("x",      (_, i) => xScale(i) - barW / 2)
                .attr("y",      p => yScale(p.value))
                .attr("width",  barW)
                .attr("height", p => Math.max(0, spH - 2 - yScale(p.value)))
                .attr("fill",   color)
                .attr("rx",     1);
        } else {
            const curveFn = d3.curveCatmullRom.alpha(0.5);

            const line = d3.line<SparkPoint>()
                .x((_, i) => xScale(i))
                .y(p => yScale(p.value))
                .curve(curveFn);

            if (type === "Area") {
                const area = d3.area<SparkPoint>()
                    .x((_, i) => xScale(i))
                    .y0(spH - 2)
                    .y1(p => yScale(p.value))
                    .curve(curveFn);

                svg.append("path")
                    .datum(pts)
                    .attr("class", "sparkline-area-path")
                    .attr("d",    area)
                    .attr("fill", color);
            }

            const linePath = svg.append("path")
                .datum(pts)
                .attr("class",  "sparkline-line-path")
                .attr("d",      line)
                .attr("stroke", color);

            // Draw-on animation
            const node    = linePath.node() as SVGPathElement;
            const pathLen = node ? node.getTotalLength() : 0;
            if (pathLen > 0) {
                linePath
                    .attr("stroke-dasharray",  pathLen)
                    .attr("stroke-dashoffset", pathLen)
                    .transition()
                    .duration(dur(700))
                    .ease(d3.easeCubicOut)
                    .attr("stroke-dashoffset", 0);
            }

            // Last-point highlight dot
            const last = pts[pts.length - 1];
            svg.append("circle")
                .attr("class", "sparkline-dot")
                .attr("cx",    xScale(pts.length - 1))
                .attr("cy",    yScale(last.value))
                .attr("r",     0)
                .attr("fill",  color)
                .transition().delay(dur(700)).duration(dur(200))
                .attr("r", 3);
        }

        // Hover tooltip overlay
        svg.append("rect")
            .attr("width",  spW)
            .attr("height", spH)
            .attr("fill",   "transparent")
            .on("mousemove", (event: MouseEvent) => {
                const [mx]  = d3.pointer(event);
                const idx   = Math.round(xScale.invert(mx));
                const i     = clamp(idx, 0, pts.length - 1);
                const p     = pts[i];

                this.clearEl(this.sparkTooltip);

                const dateEl = document.createElement("div");
                dateEl.className   = "spark-tt-date";
                dateEl.textContent = p.date;
                this.sparkTooltip.appendChild(dateEl);

                const valEl = document.createElement("div");
                valEl.className   = "spark-tt-value";
                valEl.textContent = fmtVal(p.value);
                this.sparkTooltip.appendChild(valEl);

                const rootR = this.root.getBoundingClientRect();
                let tx = event.clientX - rootR.left + 10;
                let ty = event.clientY - rootR.top  - 40;
                if (tx + 110 > rootR.width) { tx = event.clientX - rootR.left - 120; }
                if (ty < 0)                 { ty = event.clientY - rootR.top  + 14; }
                this.sparkTooltip.style.left = `${tx}px`;
                this.sparkTooltip.style.top  = `${ty}px`;
                this.sparkTooltip.classList.add("visible");
            })
            .on("mouseleave", () => {
                this.sparkTooltip.classList.remove("visible");
            });
    }

    // ── Cross-filtering ────────────────────────────────────────────────────────
    private onCardClick(): void {
        try {
            if (this.data.selectionId) {
                this.selMgr.select([this.data.selectionId], false);
            } else {
                this.selMgr.clear();
            }
        } catch { /* ignore */ }
    }

    // ── Trial / Pro ────────────────────────────────────────────────────────────
    private handleProKey(): void {
        checkMicrosoftLicence(this.host).then(p => this._msUpdateLicenceUI(p)).catch(() => this._msUpdateLicenceUI(false));
    }

    private async validateKey(key: string): Promise<boolean> {
        return checkMicrosoftLicence(this.host);
    }

    private restoreProKey(): void {
        try {
            const stored = localStorage.getItem(PRO_STORE_KEY);
            if (!stored) { return; }
            const parsed = JSON.parse(stored) as { key?: string };
            const key    = parsed?.key;
            if (!key) { return; }
            this.lastKey = key;
            this.validateKey(key).then(valid => {
                this.isPro = valid;
                if (!valid) {
                    this.lastKey = "";
                    try { localStorage.removeItem(PRO_STORE_KEY); } catch { /* ignore */ }
                }
                this.updateLicenseUI();
            });
        } catch { /* ignore */ }
    }

    private getTrialStatus(): { active: boolean; daysLeft: number } {
        try {
            let raw = localStorage.getItem(TRIAL_KEY);
            if (!raw) { raw = Date.now().toString(); localStorage.setItem(TRIAL_KEY, raw); }
            const elapsed  = Date.now() - parseInt(raw, 10);
            const daysLeft = Math.max(0, Math.ceil((TRIAL_MS - elapsed) / 86400000));
            return { active: elapsed <= TRIAL_MS, daysLeft };
        } catch {
            return { active: true, daysLeft: 4 };
        }
    }

    private updateLicenseUI(): void {
        checkMicrosoftLicence(this.host).then(p => this._msUpdateLicenceUI(p)).catch(() => this._msUpdateLicenceUI(false));
    }

    // ── Format pane ────────────────────────────────────────────────────────────

    // ── MS AppSource licence UI ──────────────────────────────────────────────
    private _msUpdateLicenceUI(isPro: boolean): void {
        const daysLeft = getTrialDaysRemaining();
        const self = this as any;

        // MS certification: never block the visual entirely
        if (self.overlay)   (self.overlay   as HTMLElement).classList.add("hidden");
        if (self.overlayEl) (self.overlayEl as HTMLElement).classList.add("hidden");

        if (isPro) {
            if (self.trialBadge) (self.trialBadge as HTMLElement).classList.add("hidden");
            if (self.proBadge)   (self.proBadge  as HTMLElement).classList.remove("hidden");
            this.removeWatermark();
            this.removeUpgradeBanner();
        } else if (daysLeft > 0) {
            const label = `Trial: ${daysLeft} day${daysLeft !== 1 ? "s" : ""} remaining`;
            if (self.trialBadge) {
                (self.trialBadge as HTMLElement).textContent = label;
                (self.trialBadge as HTMLElement).classList.remove("hidden");
            }
            if (self.proBadge) (self.proBadge as HTMLElement).classList.add("hidden");
            this.removeWatermark();
            this.removeUpgradeBanner();
        } else {
            if (self.trialBadge) (self.trialBadge as HTMLElement).classList.add("hidden");
            if (self.proBadge)   (self.proBadge  as HTMLElement).classList.add("hidden");
            this.showWatermark();
            this.showUpgradeBanner();
        }
    }

    private _msRoot(): HTMLElement {
        const s = this as any;
        return (s.root || s.target || s.el || s.container || s.svg?.node()?.parentElement) as HTMLElement;
    }

    private showWatermark(): void {
        const id  = "briqlab-ms-wm";
        const root = this._msRoot();
        if (!root || root.querySelector("#" + id)) return;
        const el = document.createElement("div");
        el.id = id;
        el.style.cssText = "position:absolute;inset:0;pointer-events:none;z-index:10;" +
            "display:flex;align-items:center;justify-content:center;" +
            "transform:rotate(-25deg);font-size:28px;font-weight:700;" +
            "color:#0D9488;opacity:0.08;font-family:Segoe UI,sans-serif;" +
            "white-space:nowrap;user-select:none";
        el.textContent = "BRIQLAB PRO";
        root.appendChild(el);
    }

    private removeWatermark(): void {
        const root = this._msRoot();
        if (!root) return;
        const el = root.querySelector("#briqlab-ms-wm");
        if (el) el.remove();
    }

    private showUpgradeBanner(): void {
        const id   = "briqlab-ms-banner";
        const root = this._msRoot();
        if (!root || root.querySelector("#" + id)) return;
        const banner = document.createElement("div");
        banner.id = id;
        banner.style.cssText = "position:absolute;bottom:0;left:0;right:0;height:30px;" +
            "background:rgba(13,148,136,0.95);display:flex;align-items:center;" +
            "justify-content:center;gap:8px;z-index:20;font-size:11px;" +
            "color:#fff;font-family:Segoe UI,sans-serif;padding:0 8px";
        const msg = document.createElement("span");
        msg.textContent = "\uD83D\uDD13 Unlock all features \u00B7 Briqlab Pro on AppSource";
        banner.appendChild(msg);
        const btn = document.createElement("button");
        btn.textContent = getButtonText();
        btn.style.cssText = "background:#F97316;border:none;color:#fff;border-radius:4px;" +
            "padding:2px 8px;font-size:10px;cursor:pointer;font-weight:600";
        btn.addEventListener("click", () => this.host.launchUrl(getPurchaseUrl()));
        banner.appendChild(btn);
        root.appendChild(banner);
    }

    private removeUpgradeBanner(): void {
        const root = this._msRoot();
        if (!root) return;
        const el = root.querySelector("#briqlab-ms-banner");
        if (el) el.remove();
    }

    private sanitise(input: string): string {
        if (!input) return "";
        return String(input)
            .replace(/&/g, "&amp;").replace(/</g, "&lt;")
            .replace(/>/g, "&gt;").replace(/"/g, "&quot;")
            .replace(/'/g, "&#x27;");
    }


    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.fmtSvc.buildFormattingModel(this.settings);
    }

    // ── DOM utilities ──────────────────────────────────────────────────────────
    private mkDiv(cls: string): HTMLDivElement {
        const el = document.createElement("div");
        el.className = cls;
        return el;
    }

    private clearEl(el: HTMLElement): void {
        while (el.firstChild) { el.removeChild(el.firstChild); }
    }
}
