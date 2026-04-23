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
import DataView                 = powerbi.DataView;

import { VisualFormattingSettingsModel } from "./settings";
import { checkMicrosoftLicence, resetLicenceCache } from "./licenceManager";
import { getTrialDaysRemaining, isTrialExpired, getPurchaseUrl, getButtonText } from "./trialManager";

// ── Types ──────────────────────────────────────────────────────────────────────
interface BubblePoint {
    name:         string;
    xVal:         number;
    yVal:         number;
    sizeVal:      number;
    radius:       number;
    color:        string;
    colorGroup:   string;
    selectionId:  ISelectionId;
    tooltipExtra: { displayName: string; value: string }[];
}

interface DrillFrame {
    label: string;
    data:  BubblePoint[];
}

// ── Constants ──────────────────────────────────────────────────────────────────
const TRIAL_KEY     = "briqlab_trial_drillbubble_start";
const PRO_STORE_KEY = "briqlab_key_drillbubble";
const TRIAL_MS      = 4 * 24 * 60 * 60 * 1000;

// ── Helpers ────────────────────────────────────────────────────────────────────
function fmtVal(v: number): string {
    const a = Math.abs(v);
    if (a >= 1e9) { const r = (v / 1e9).toFixed(1); return (r.endsWith(".0") ? r.slice(0, -2) : r) + "B"; }
    if (a >= 1e6) { const r = (v / 1e6).toFixed(1); return (r.endsWith(".0") ? r.slice(0, -2) : r) + "M"; }
    if (a >= 1e3) { const r = (v / 1e3).toFixed(1); return (r.endsWith(".0") ? r.slice(0, -2) : r) + "K"; }
    return v.toLocaleString("en-US");
}

function truncate(s: string, max: number): string {
    return s.length > max ? s.slice(0, max - 1) + "\u2026" : s;
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

    // Services
    private readonly host:   IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private readonly selMgr: ISelectionManager;
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private _handlersAttached = false;
    private readonly fmtSvc: FormattingSettingsService;

    // DOM roots
    private readonly root:         HTMLElement;
    private readonly contentEl:    HTMLDivElement;
    private readonly chartAreaEl:  HTMLDivElement;
    private readonly legendEl:     HTMLDivElement;
    private readonly breadcrumbEl: HTMLDivElement;
    private readonly drillUpBtn:   HTMLButtonElement;
    private readonly tooltipEl:    HTMLDivElement;
    private readonly trialBadge:   HTMLDivElement;
    private readonly proBadge:     HTMLDivElement;
    private readonly keyErrorEl:   HTMLDivElement;
    private readonly overlayEl:    HTMLDivElement;

    // SVG groups
    private readonly svg:      d3.Selection<SVGSVGElement, unknown, null, undefined>;
    private readonly gGrid:    d3.Selection<SVGGElement, unknown, null, undefined>;
    private readonly gXAxis:   d3.Selection<SVGGElement, unknown, null, undefined>;
    private readonly gYAxis:   d3.Selection<SVGGElement, unknown, null, undefined>;
    private readonly gBubbles: d3.Selection<SVGGElement, unknown, null, undefined>;
    private readonly gLabels:  d3.Selection<SVGGElement, unknown, null, undefined>;
    private readonly gParent:  d3.Selection<SVGGElement, unknown, null, undefined>;

    // State
    private settings!:    VisualFormattingSettingsModel;
    private viewport:     powerbi.IViewport = { width: 300, height: 300 };
    private currentData:  BubblePoint[]     = [];
    private selectedIds:  Set<string>       = new Set();
    private drillStack:   DrillFrame[]      = [];

    // Trial / Pro
    private isPro   = false;
    private lastKey = "";
    private readonly keyCache: Map<string, boolean> = new Map();

    // Tooltip debounce
    private tooltipShowTimer: ReturnType<typeof setTimeout> | null = null;
    private tooltipHideTimer: ReturnType<typeof setTimeout> | null = null;

    // ── Constructor ────────────────────────────────────────────────────────────
    constructor(options: VisualConstructorOptions) {
        this.host   = options.host;
        this.renderingManager = options.host.eventService;
        this.selMgr = options.host.createSelectionManager();
        this.tooltipSvc = this.host.tooltipService;
        this.fmtSvc = new FormattingSettingsService();

        this.root = options.element;
        this.root.classList.add("briqlab-drill-bubble");

        this.contentEl = this.mkDiv("briqlab-visual-content");
        this.root.appendChild(this.contentEl);

        this.breadcrumbEl = this.mkDiv("briqlab-breadcrumb hidden");
        this.contentEl.appendChild(this.breadcrumbEl);

        this.chartAreaEl = this.mkDiv("briqlab-chart-area");
        this.contentEl.appendChild(this.chartAreaEl);

        this.drillUpBtn = document.createElement("button");
        this.drillUpBtn.className = "briqlab-drill-up hidden";
        this.drillUpBtn.textContent = "\u2190 Drill Up";
        this.drillUpBtn.addEventListener("click", () => this.drillUp());
        this.chartAreaEl.appendChild(this.drillUpBtn);

        this.svg = d3.select(this.chartAreaEl)
            .append<SVGSVGElement>("svg")
            .attr("class", "briqlab-svg");

        this.gGrid    = this.svg.append("g").attr("class", "g-grid");
        this.gXAxis   = this.svg.append("g").attr("class", "g-x-axis");
        this.gYAxis   = this.svg.append("g").attr("class", "g-y-axis");
        this.gParent  = this.svg.append("g").attr("class", "g-parent");
        this.gBubbles = this.svg.append("g").attr("class", "g-bubbles");
        this.gLabels  = this.svg.append("g").attr("class", "g-labels");

        this.svg.on("click", (event: MouseEvent) => {
            const t = event.target as SVGElement;
            if (t === this.svg.node() || t.tagName === "svg") {
                this.selectedIds.clear();
                this.selMgr.clear();
                this.applySelectionState();
            }
        });

        this.legendEl = this.mkDiv("briqlab-legend legend-bottom");
        this.contentEl.appendChild(this.legendEl);

        this.tooltipEl = this.mkDiv("briqlab-tooltip");
        this.root.appendChild(this.tooltipEl);

        this.trialBadge = this.mkDiv("briqlab-trial-badge hidden");
        this.root.appendChild(this.trialBadge);

        this.proBadge = this.mkDiv("briqlab-pro-badge hidden");
        this.root.appendChild(this.proBadge);

        this.keyErrorEl = this.mkDiv("briqlab-key-error hidden");
        this.keyErrorEl.textContent = "\u2717 Invalid key";
        this.root.appendChild(this.keyErrorEl);

        this.overlayEl = this.mkDiv("briqlab-trial-overlay hidden");
        this.buildTrialOverlay();
        this.root.appendChild(this.overlayEl);

        this.restoreProKey();
    }

    // ── Trial overlay (no innerHTML) ───────────────────────────────────────────
    private buildTrialOverlay(): void {
        const card = this.mkDiv("briqlab-trial-card");

        const iconSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        iconSvg.setAttribute("viewBox", "0 0 56 56");
        iconSvg.setAttribute("fill", "none");
        iconSvg.classList.add("trial-icon");
        this.buildBubbleIconSvg(iconSvg, "#0D9488", 0.5);
        card.appendChild(iconSvg);

        const title = document.createElement("h2");
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
        btn.addEventListener("click", () => { this.host.launchUrl(getPurchaseUrl()); });
        card.appendChild(btn);

        const sub = document.createElement("p");
        sub.className = "trial-subtext";
        sub.textContent = "Purchase on Microsoft AppSource to unlock all features instantly.";
        card.appendChild(sub);

        this.overlayEl.appendChild(card);
    }

    private buildBubbleIconSvg(svgEl: SVGSVGElement, color: string, opacity: number): void {
        const ns = "http://www.w3.org/2000/svg";
        const mkC = (cx: string, cy: string, r: string, col: string) => {
            const c = document.createElementNS(ns, "circle");
            c.setAttribute("cx", cx); c.setAttribute("cy", cy); c.setAttribute("r", r);
            c.setAttribute("fill", col); c.setAttribute("opacity", opacity.toString());
            svgEl.appendChild(c);
        };
        mkC("28", "28", "14", color);
        mkC("10", "18", "7",  "#F97316");
        mkC("44", "14", "5",  "#3B82F6");
        mkC("46", "38", "8",  "#8B5CF6");
        mkC("14", "42", "5",  "#10B981");
    }

    // ── Empty state ────────────────────────────────────────────────────────────
    private buildEmptyState(): HTMLDivElement {
        const wrap = this.mkDiv("briqlab-empty");

        const iconSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        iconSvg.setAttribute("viewBox", "0 0 56 56");
        iconSvg.setAttribute("fill", "none");
        iconSvg.classList.add("empty-icon");
        this.buildBubbleIconSvg(iconSvg, "#0D9488", 0.4);
        wrap.appendChild(iconSvg);

        const t = document.createElement("p");
        t.className = "empty-title";
        t.textContent = "Connect your data";
        wrap.appendChild(t);

        const b = document.createElement("p");
        b.className = "empty-body";
        b.textContent = "Add Category, X Value, Y Value and Bubble Size to get started";
        wrap.appendChild(b);

        return wrap;
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
                        dataItems: [{ displayName: "Briqlab Drill Down Bubble", value: "" }],
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
    
                const dv = options.dataViews?.[0];
                if (!dv?.categorical?.categories?.length || !dv.categorical.values?.length) {
                    this.renderEmpty();
                    return;
                }
    
                this.currentData = this.parseData(dv);
                if (!this.currentData.length) {
                    this.renderEmpty();
                    return;
                }
    
                const emptyEl = this.chartAreaEl.querySelector(".briqlab-empty");
                if (emptyEl) { emptyEl.remove(); }
    
                this.render();
    
            } catch (err) {
                console.error("[BriqlabDrillBubble] update error:", err);
            }
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    // ── Data parsing ───────────────────────────────────────────────────────────
    private parseData(dv: DataView): BubblePoint[] {
        const cats        = dv.categorical!.categories!;
        const vals        = dv.categorical!.values!;
        const catCol      = cats[0];
        const colorGrpCol = cats.length > 1 ? cats[1] : null;

        let xCol:    powerbi.DataViewValueColumn | null = null;
        let yCol:    powerbi.DataViewValueColumn | null = null;
        let sizeCol: powerbi.DataViewValueColumn | null = null;
        const extraCols: powerbi.DataViewValueColumn[] = [];

        for (let i = 0; i < vals.length; i++) {
            const col   = vals[i];
            const roles = col.source.roles ?? {};
            if (roles["xAxis"]      && !xCol)    { xCol    = col; continue; }
            if (roles["yAxis"]      && !yCol)    { yCol    = col; continue; }
            if (roles["bubbleSize"] && !sizeCol) { sizeCol = col; continue; }
            if (roles["tooltipValues"])           { extraCols.push(col); }
        }

        if (!sizeCol) { return []; }

        const sizeVals = sizeCol.values as number[];
        const xVals    = xCol    ? (xCol.values    as number[]) : null;
        const yVals    = yCol    ? (yCol.values    as number[]) : null;

        const minR = clamp(this.settings.chartStyle.minBubbleSize.value, 4, 100);
        const maxR = clamp(this.settings.chartStyle.maxBubbleSize.value, 4, 200);

        const rawSizes  = sizeVals.map(v => Math.abs((v as number) ?? 0));
        const sizeMin   = d3.min(rawSizes) ?? 0;
        const sizeMax   = d3.max(rawSizes) ?? 1;
        const sizeRange = sizeMax - sizeMin || 1;

        const palette     = this.buildPalette();
        const colorGroups = colorGrpCol
            ? Array.from(new Set(colorGrpCol.values.map(v => String(v ?? "(blank)"))))
            : [];

        const result: BubblePoint[] = [];

        for (let i = 0; i < catCol.values.length; i++) {
            const raw = rawSizes[i];
            if (raw <= 0) { continue; }

            const t      = (raw - sizeMin) / sizeRange;
            const radius = minR + t * (maxR - minR);

            const grpName  = colorGrpCol ? String(colorGrpCol.values[i] ?? "(blank)") : "";
            const colorIdx = colorGroups.length > 0 ? colorGroups.indexOf(grpName) : i;
            const color    = palette[colorIdx % palette.length];

            const extra = extraCols.map(col => ({
                displayName: col.source.displayName || "",
                value:       fmtVal((col.values[i] as number) ?? 0)
            }));

            result.push({
                name:        String(catCol.values[i] ?? "(blank)"),
                xVal:        xVals ? ((xVals[i] as number) ?? 0) : i,
                yVal:        yVals ? ((yVals[i] as number) ?? 0) : raw,
                sizeVal:     raw,
                radius,
                color,
                colorGroup:  grpName,
                tooltipExtra: extra,
                selectionId: this.host.createSelectionIdBuilder()
                    .withCategory(catCol, i)
                    .createSelectionId()
            });
        }

        return result;
    }

    private buildPalette(): string[] {
        const c = this.settings.bubbleColors;
        return [
            c.color1.value.value,  c.color2.value.value,  c.color3.value.value,
            c.color4.value.value,  c.color5.value.value,  c.color6.value.value,
            c.color7.value.value,  c.color8.value.value,  c.color9.value.value,
            c.color10.value.value
        ];
    }

    // ── Main render ────────────────────────────────────────────────────────────
    private render(): void {
        const { width, height } = this.viewport;
        const isTiny = width < 120 || height < 120;
        this.root.classList.toggle("tiny", isTiny);

        const legendPos  = this.settings.legendSettings.legendPosition.value.value as string;
        const showLegend = this.settings.legendSettings.showLegend.value && legendPos !== "None";

        this.layoutLegend(legendPos, showLegend && !isTiny);

        let legendH = 0;
        let legendW = 0;
        if (showLegend && !isTiny) {
            if (legendPos === "Left" || legendPos === "Right") { legendW = 140; }
            else { legendH = 56; }
        }

        const hasDrill = this.drillStack.length > 0;
        const bcH      = hasDrill && !isTiny ? 26 : 0;
        const chartW   = width  - legendW;
        const chartH   = height - legendH;

        this.svg.attr("width", chartW).attr("height", chartH);

        const mode = this.settings.chartStyle.layoutMode.value.value as string;
        if (mode === "Packed") {
            this.renderPacked(chartW, chartH, bcH, isTiny);
        } else {
            this.renderScatter(chartW, chartH, bcH, isTiny);
        }

        this.renderBreadcrumb();

        if (showLegend && !isTiny) {
            this.renderLegend();
        } else {
            this.clearEl(this.legendEl);
        }
    }

    // ── Scatter mode ───────────────────────────────────────────────────────────
    private renderScatter(chartW: number, chartH: number, bcH: number, isTiny: boolean): void {
        const showX = this.settings.xAxisSettings.showAxis.value && !isTiny;
        const showY = this.settings.yAxisSettings.showAxis.value && !isTiny;

        const margin = {
            top:    bcH + 14,
            right:  20,
            bottom: showX ? 44 : 16,
            left:   showY ? 54 : 16
        };

        const plotW = Math.max(10, chartW - margin.left - margin.right);
        const plotH = Math.max(10, chartH - margin.top  - margin.bottom);

        const xs = this.buildLinearScale(
            this.currentData.map(d => d.xVal),
            [0, plotW],
            this.settings.xAxisSettings.axisMin.value,
            this.settings.xAxisSettings.axisMax.value
        );
        const ys = this.buildLinearScale(
            this.currentData.map(d => d.yVal),
            [plotH, 0],
            this.settings.yAxisSettings.axisMin.value,
            this.settings.yAxisSettings.axisMax.value
        );

        // Axis groups positioned
        this.gXAxis.attr("transform", `translate(${margin.left},${margin.top + plotH})`);
        this.gYAxis.attr("transform", `translate(${margin.left},${margin.top})`);
        this.gGrid.attr("transform",  `translate(${margin.left},${margin.top})`);
        this.gBubbles.attr("transform", `translate(${margin.left},${margin.top})`);
        this.gLabels.attr("transform",  `translate(${margin.left},${margin.top})`);
        this.gParent.selectAll("*").remove();

        if (showX) {
            const xAxis = d3.axisBottom(xs).ticks(5).tickSizeOuter(0);
            this.gXAxis.call(xAxis as unknown as (sel: d3.Selection<SVGGElement, unknown, null, undefined>) => void);
            this.gXAxis.attr("class", "axis-group");
            this.gXAxis.selectAll(".x-axis-label").remove();
            const xLabel = this.settings.xAxisSettings.axisLabel.value.trim();
            if (xLabel) {
                this.gXAxis.append("text")
                    .attr("class", "axis-label x-axis-label")
                    .attr("x", plotW / 2)
                    .attr("y", 38)
                    .text(xLabel);
            }
            this.gGrid.selectAll(".x-grid").remove();
            this.gGrid.selectAll<SVGLineElement, number>(".x-grid")
                .data(xs.ticks(5))
                .enter().append("line")
                .attr("class", "axis-grid-line x-grid")
                .attr("x1", d => xs(d)).attr("x2", d => xs(d))
                .attr("y1", 0).attr("y2", plotH);
        } else {
            this.gXAxis.selectAll("*").remove();
            this.gGrid.selectAll(".x-grid").remove();
        }

        if (showY) {
            const yAxis = d3.axisLeft(ys).ticks(5).tickSizeOuter(0);
            this.gYAxis.call(yAxis as unknown as (sel: d3.Selection<SVGGElement, unknown, null, undefined>) => void);
            this.gYAxis.attr("class", "axis-group");
            this.gYAxis.selectAll(".y-axis-label").remove();
            const yLabel = this.settings.yAxisSettings.axisLabel.value.trim();
            if (yLabel) {
                this.gYAxis.append("text")
                    .attr("class", "axis-label y-axis-label")
                    .attr("transform", "rotate(-90)")
                    .attr("x", -(plotH / 2))
                    .attr("y", -44)
                    .text(yLabel);
            }
            this.gGrid.selectAll(".y-grid").remove();
            this.gGrid.selectAll<SVGLineElement, number>(".y-grid")
                .data(ys.ticks(5))
                .enter().append("line")
                .attr("class", "axis-grid-line y-grid")
                .attr("x1", 0).attr("x2", plotW)
                .attr("y1", d => ys(d)).attr("y2", d => ys(d));
        } else {
            this.gYAxis.selectAll("*").remove();
            this.gGrid.selectAll(".y-grid").remove();
        }

        const opacity = clamp(this.settings.chartStyle.bubbleOpacity.value, 10, 100) / 100;
        const isFirst = this.gBubbles.selectAll(".bubble-circle").size() === 0;

        const circles = this.gBubbles
            .selectAll<SVGCircleElement, BubblePoint>(".bubble-circle")
            .data(this.currentData, d => d.name);

        circles.exit()
            .transition().duration(dur(250))
            .attr("r", 0)
            .remove();

        const entered = circles.enter()
            .append<SVGCircleElement>("circle")
            .attr("class",        "bubble-circle")
            .attr("cx",           d => xs(d.xVal))
            .attr("cy",           d => ys(d.yVal))
            .attr("r",            isFirst ? 0 : d => d.radius)
            .attr("fill",         d => d.color)
            .attr("fill-opacity", opacity)
            .attr("stroke",       d => d.color)
            .attr("stroke-width", 1)
            .on("click",      (event: MouseEvent, d) => this.onBubbleClick(event, d))
            .on("mouseenter", (event: MouseEvent, d) => this.onBubbleEnter(event, d))
            .on("mouseleave", (event: MouseEvent, d) => this.onBubbleLeave(event, d));

        const merged = entered.merge(circles);
        merged
            .attr("fill",         d => d.color)
            .attr("stroke",       d => d.color)
            .attr("fill-opacity", opacity);

        if (isFirst) {
            merged.each((d, i, nodes) => {
                d3.select(nodes[i])
                    .transition()
                    .delay(dur(i * 40))
                    .duration(dur(500))
                    .ease(d3.easeCubicOut)
                    .attr("r",  d.radius)
                    .attr("cx", xs(d.xVal))
                    .attr("cy", ys(d.yVal));
            });
        } else {
            merged
                .transition().duration(dur(400)).ease(d3.easeCubicOut)
                .attr("cx", d => xs(d.xVal))
                .attr("cy", d => ys(d.yVal))
                .attr("r",  d => d.radius);
        }

        this.applySelectionState();

        this.gLabels.selectAll("*").remove();
        if (this.settings.labelSettings.showLabels.value && !isTiny) {
            this.renderScatterLabels(xs, ys);
        }
    }

    private buildLinearScale(
        vals: number[],
        range: [number, number],
        settingMin: number,
        settingMax: number
    ): d3.ScaleLinear<number, number> {
        const dataMin = d3.min(vals) ?? 0;
        const dataMax = d3.max(vals) ?? 1;
        const pad     = (dataMax - dataMin) * 0.1 || 1;
        const lo      = settingMin !== 0 ? settingMin : dataMin - pad;
        const hi      = settingMax !== 0 ? settingMax : dataMax + pad;
        return d3.scaleLinear().domain([lo, hi]).range(range).nice();
    }

    private renderScatterLabels(
        xs: d3.ScaleLinear<number, number>,
        ys: d3.ScaleLinear<number, number>
    ): void {
        const fontSize = this.settings.labelSettings.labelFontSize.value;
        const placed: { x: number; y: number; w: number; h: number }[] = [];

        this.currentData.forEach(d => {
            const cx    = xs(d.xVal);
            const cy    = ys(d.yVal);
            const label = truncate(d.name, 14);
            const lw    = label.length * (fontSize * 0.55) + 4;
            const lh    = fontSize + 4;

            const candidates = [
                { x: cx,        y: cy - d.radius - 6 },
                { x: cx,        y: cy + d.radius + lh },
                { x: cx + d.radius + 6, y: cy },
                { x: cx - d.radius - 6, y: cy }
            ];

            let best = candidates[0];
            for (const cand of candidates) {
                const rect = { x: cand.x - lw / 2, y: cand.y - lh / 2, w: lw, h: lh };
                const overlaps = placed.some(p =>
                    rect.x < p.x + p.w && rect.x + rect.w > p.x &&
                    rect.y < p.y + p.h && rect.y + rect.h > p.y
                );
                if (!overlaps) { best = cand; break; }
            }
            placed.push({ x: best.x - lw / 2, y: best.y - lh / 2, w: lw, h: lh });

            this.gLabels.append("text")
                .attr("class",      "bubble-label bubble-label-outside")
                .attr("x",          best.x)
                .attr("y",          best.y)
                .attr("dy",         "0.35em")
                .style("font-size", `${fontSize}px`)
                .text(label);
        });
    }

    // ── Packed mode ────────────────────────────────────────────────────────────
    private renderPacked(chartW: number, chartH: number, bcH: number, isTiny: boolean): void {
        this.gXAxis.selectAll("*").remove();
        this.gYAxis.selectAll("*").remove();
        this.gGrid.selectAll("*").remove();

        const availH  = chartH - bcH - 8;
        const packDia = Math.min(chartW, availH) * 0.92;
        const packR   = packDia / 2;
        const cx      = chartW / 2;
        const cy      = bcH + availH / 2 + 4;

        // Build D3 hierarchy for pack layout
        interface HierarchyDatum { name: string; value?: number; children?: HierarchyDatum[] }
        const children: HierarchyDatum[] = this.currentData.map(d => ({ name: d.name, value: d.sizeVal }));
        const root = d3.hierarchy<HierarchyDatum>({ name: "root", children })
            .sum(d => d.value ?? 0);

        const pack = d3.pack<HierarchyDatum>().size([packDia, packDia]).padding(3);
        const packed = pack(root);

        // Map packed positions back to data points
        const posMap = new Map<string, { px: number; py: number; pr: number }>();
        packed.leaves().forEach(leaf => {
            posMap.set(leaf.data.name, {
                px: leaf.x + (cx - packR),
                py: leaf.y + (cy - packR),
                pr: leaf.r
            });
        });

        // Parent ring for drill context
        this.gParent.selectAll("*").remove();
        if (this.drillStack.length > 0) {
            this.gParent.append("circle")
                .attr("class", "packed-parent-ring")
                .attr("cx", cx)
                .attr("cy", cy)
                .attr("r",  packR + 10);

            const frame = this.drillStack[this.drillStack.length - 1];
            this.gParent.append("text")
                .attr("class", "packed-parent-label")
                .attr("x", cx)
                .attr("y", cy - packR - 14)
                .text(truncate(frame.label, 30));
        }

        this.gBubbles.attr("transform", "");
        this.gLabels.attr("transform",  "");

        const opacity = clamp(this.settings.chartStyle.bubbleOpacity.value, 10, 100) / 100;
        const isFirst = this.gBubbles.selectAll(".bubble-circle").size() === 0;

        const circles = this.gBubbles
            .selectAll<SVGCircleElement, BubblePoint>(".bubble-circle")
            .data(this.currentData, d => d.name);

        circles.exit()
            .transition().duration(dur(300))
            .attr("r", 0)
            .remove();

        const entered = circles.enter()
            .append<SVGCircleElement>("circle")
            .attr("class",        "bubble-circle")
            .attr("fill",         d => d.color)
            .attr("fill-opacity", opacity)
            .attr("stroke",       d => d.color)
            .attr("stroke-width", 1)
            .on("click",      (event: MouseEvent, d) => this.onBubbleClick(event, d))
            .on("mouseenter", (event: MouseEvent, d) => this.onBubbleEnter(event, d))
            .on("mouseleave", (event: MouseEvent, d) => this.onBubbleLeave(event, d));

        const merged = entered.merge(circles);
        merged
            .attr("fill",         d => d.color)
            .attr("stroke",       d => d.color)
            .attr("fill-opacity", opacity);

        if (isFirst) {
            // Emerge from centre
            merged
                .attr("cx", cx).attr("cy", cy).attr("r", 0)
                .each((d, i, nodes) => {
                    const pos = posMap.get(d.name);
                    if (!pos) { return; }
                    d3.select(nodes[i])
                        .transition()
                        .delay(dur(i * 30))
                        .duration(dur(500))
                        .ease(d3.easeCubicOut)
                        .attr("cx", pos.px)
                        .attr("cy", pos.py)
                        .attr("r",  pos.pr);
                });
        } else {
            // Smooth reposition
            merged
                .transition().duration(dur(450)).ease(d3.easeCubicOut)
                .attr("cx", d => posMap.get(d.name)?.px ?? cx)
                .attr("cy", d => posMap.get(d.name)?.py ?? cy)
                .attr("r",  d => posMap.get(d.name)?.pr ?? d.radius);
        }

        this.applySelectionState();

        this.gLabels.selectAll("*").remove();
        if (this.settings.labelSettings.showLabels.value && !isTiny) {
            this.renderPackedLabels(posMap);
        }
    }

    private renderPackedLabels(posMap: Map<string, { px: number; py: number; pr: number }>): void {
        const fontSize = this.settings.labelSettings.labelFontSize.value;

        this.currentData.forEach(d => {
            const pos = posMap.get(d.name);
            if (!pos) { return; }

            if (pos.pr > 20) {
                // Label inside the bubble
                const maxChars = Math.floor(pos.pr * 1.8 / (fontSize * 0.6));
                const label    = truncate(d.name, Math.max(3, maxChars));
                const fs       = Math.min(fontSize, pos.pr * 0.42);

                this.gLabels.append("text")
                    .attr("class",      "bubble-label bubble-label-inside")
                    .attr("x",          pos.px)
                    .attr("y",          pos.py)
                    .style("font-size", `${fs}px`)
                    .text(label);
            } else if (pos.pr > 8) {
                // Small label above bubble
                this.gLabels.append("text")
                    .attr("class",      "bubble-label bubble-label-outside")
                    .attr("x",          pos.px)
                    .attr("y",          pos.py - pos.pr - 4)
                    .style("font-size", `${Math.min(fontSize, 9)}px`)
                    .text(truncate(d.name, 8));
            }
        });
    }

    // ── Selection state ────────────────────────────────────────────────────────
    private applySelectionState(): void {
        const hasSel = this.selectedIds.size > 0;
        this.gBubbles.selectAll<SVGCircleElement, BubblePoint>(".bubble-circle")
            .classed("bubble-circle-faded",    d => hasSel && !this.selectedIds.has(this.selKey(d)))
            .classed("bubble-circle-selected", d => hasSel &&  this.selectedIds.has(this.selKey(d)));
    }

    private selKey(d: BubblePoint): string {
        return ((d.selectionId as unknown as Record<string, unknown>)["key"] as string) || d.name;
    }

    // ── Interactions ───────────────────────────────────────────────────────────
    private onBubbleClick(event: MouseEvent, d: BubblePoint): void {
        try {
            const idKey   = this.selKey(d);
            const isMulti = event.ctrlKey || event.metaKey;
            const wasSel  = this.selectedIds.has(idKey);

            if (isMulti) {
                wasSel ? this.selectedIds.delete(idKey) : this.selectedIds.add(idKey);
            } else {
                if (wasSel && this.selectedIds.size === 1) {
                    this.selectedIds.clear();
                } else {
                    this.selectedIds.clear();
                    this.selectedIds.add(idKey);
                }
            }

            if (this.selectedIds.size > 0) {
                const ids = this.currentData
                    .filter(p => this.selectedIds.has(this.selKey(p)))
                    .map(p => p.selectionId);
                this.selMgr.select(ids, isMulti).then(() => this.applySelectionState());
            } else {
                this.selMgr.clear().then(() => this.applySelectionState());
            }

            this.applySelectionState();
            event.stopPropagation();
        } catch (err) {
            console.warn("[BriqlabDrillBubble] click:", err);
        }
    }

    private onBubbleEnter(event: MouseEvent, d: BubblePoint): void {
        d3.select(event.currentTarget as SVGCircleElement)
            .raise()
            .transition().duration(dur(150))
            .attr("stroke-width", "2.5")
            .attr("stroke", "#fff");

        if (this.settings.tooltipSettings.showTooltip.value) {
            if (this.tooltipHideTimer) { clearTimeout(this.tooltipHideTimer); this.tooltipHideTimer = null; }
            this.tooltipShowTimer = setTimeout(() => this.showTooltip(event, d), 150);
        }
    }

    private onBubbleLeave(event: MouseEvent, d: BubblePoint): void {
        d3.select(event.currentTarget as SVGCircleElement)
            .transition().duration(dur(150))
            .attr("stroke-width", "1")
            .attr("stroke", d.color);

        if (this.tooltipShowTimer) { clearTimeout(this.tooltipShowTimer); this.tooltipShowTimer = null; }
        this.tooltipHideTimer = setTimeout(() => this.tooltipEl.classList.remove("visible"), 100);
    }

    // ── Drill navigation ───────────────────────────────────────────────────────
    private drillUp(): void {
        if (!this.drillStack.length) { return; }
        const frame      = this.drillStack.pop()!;
        this.currentData = frame.data;
        this.selectedIds.clear();
        this.selMgr.clear();
        this.render();
    }

    private drillUpTo(depth: number): void {
        if (depth === 0 && this.drillStack.length) {
            this.currentData = this.drillStack[0].data;
            this.drillStack  = [];
        } else {
            while (this.drillStack.length > depth) {
                const frame      = this.drillStack.pop()!;
                this.currentData = frame.data;
            }
        }
        this.selectedIds.clear();
        this.selMgr.clear();
        this.render();
    }

    // ── Breadcrumb ─────────────────────────────────────────────────────────────
    private renderBreadcrumb(): void {
        this.clearEl(this.breadcrumbEl);

        if (this.drillStack.length === 0) {
            this.breadcrumbEl.classList.add("hidden");
            this.drillUpBtn.classList.add("hidden");
            return;
        }

        this.breadcrumbEl.classList.remove("hidden");
        this.drillUpBtn.classList.remove("hidden");

        const addCrumb = (label: string, depth: number | null) => {
            const span = document.createElement("span");
            span.className = depth !== null ? "crumb" : "crumb crumb-current";
            span.textContent = label;
            if (depth !== null) {
                const cap = depth;
                span.addEventListener("click", () => this.drillUpTo(cap));
            }
            this.breadcrumbEl.appendChild(span);
        };
        const addSep = () => {
            const sep = document.createElement("span");
            sep.className = "crumb-sep";
            sep.textContent = "\u203A";
            this.breadcrumbEl.appendChild(sep);
        };

        addCrumb("All", 0);
        this.drillStack.forEach((frame, i) => {
            addSep();
            const isLast = i === this.drillStack.length - 1;
            addCrumb(frame.label, isLast ? null : i + 1);
        });
    }

    // ── Legend ─────────────────────────────────────────────────────────────────
    private renderLegend(): void {
        this.clearEl(this.legendEl);
        const fontSize = this.settings.legendSettings.legendFontSize.value;
        this.legendEl.style.fontSize = `${fontSize}px`;

        const hasGroups = this.currentData.some(d => d.colorGroup !== "");
        const entries: { name: string; color: string }[] = [];
        const seen = new Set<string>();

        this.currentData.forEach(d => {
            const key = hasGroups ? d.colorGroup : d.name;
            if (seen.has(key)) { return; }
            seen.add(key);
            entries.push({ name: key || d.name, color: d.color });
        });

        entries.forEach(e => {
            const item   = this.mkDiv("briqlab-legend-item");
            const swatch = this.mkDiv("legend-swatch");
            swatch.style.background = e.color;
            const nameEl = this.mkDiv("legend-name");
            nameEl.textContent = truncate(e.name, 20);
            item.appendChild(swatch);
            item.appendChild(nameEl);
            this.legendEl.appendChild(item);
        });
    }

    private layoutLegend(pos: string, show: boolean): void {
        if (!show || pos === "None") {
            this.legendEl.className = "briqlab-legend legend-hidden";
            return;
        }
        this.legendEl.className = "briqlab-legend";
        if (pos === "Top") {
            this.legendEl.classList.add("legend-top");
            this.contentEl.style.flexDirection = "column";
            this.legendEl.style.cssText = "order:-1;width:100%;max-height:70px;flex:0 0 auto;";
        } else if (pos === "Bottom") {
            this.legendEl.classList.add("legend-bottom");
            this.contentEl.style.flexDirection = "column";
            this.legendEl.style.cssText = "order:1;width:100%;max-height:70px;flex:0 0 auto;";
        } else if (pos === "Left") {
            this.legendEl.classList.add("legend-left");
            this.contentEl.style.flexDirection = "row";
            this.legendEl.style.cssText = "order:-1;width:140px;flex:0 0 140px;max-height:100%;";
        } else if (pos === "Right") {
            this.legendEl.classList.add("legend-right");
            this.contentEl.style.flexDirection = "row";
            this.legendEl.style.cssText = "order:1;width:140px;flex:0 0 140px;max-height:100%;";
        }
    }

    // ── Empty state ────────────────────────────────────────────────────────────
    private renderEmpty(): void {
        this.gBubbles.selectAll("*").remove();
        this.gLabels.selectAll("*").remove();
        this.gXAxis.selectAll("*").remove();
        this.gYAxis.selectAll("*").remove();
        this.gGrid.selectAll("*").remove();
        this.gParent.selectAll("*").remove();
        this.clearEl(this.legendEl);
        this.breadcrumbEl.classList.add("hidden");
        this.drillUpBtn.classList.add("hidden");

        if (!this.chartAreaEl.querySelector(".briqlab-empty")) {
            this.chartAreaEl.appendChild(this.buildEmptyState());
        }
    }

    // ── Tooltip ────────────────────────────────────────────────────────────────
    private showTooltip(event: MouseEvent, d: BubblePoint): void {
        this.clearEl(this.tooltipEl);

        const header = this.mkDiv("tooltip-header");
        const dot    = this.mkDiv("tooltip-dot");
        dot.style.background = d.color;
        const nameEl = this.mkDiv("tooltip-name");
        nameEl.textContent = d.name;
        header.appendChild(dot);
        header.appendChild(nameEl);
        this.tooltipEl.appendChild(header);

        this.tooltipEl.appendChild(this.mkTooltipRow("X Value",     fmtVal(d.xVal)));
        this.tooltipEl.appendChild(this.mkTooltipRow("Y Value",     fmtVal(d.yVal)));
        this.tooltipEl.appendChild(this.mkTooltipRow("Bubble Size", fmtVal(d.sizeVal)));

        if (d.colorGroup) {
            this.tooltipEl.appendChild(this.mkTooltipRow("Group", d.colorGroup));
        }

        if (d.tooltipExtra.length) {
            this.tooltipEl.appendChild(this.mkDiv("tooltip-divider"));
            d.tooltipExtra.forEach(e => {
                const row = this.mkDiv("tooltip-extra-row");
                const lbl = document.createElement("span");
                lbl.textContent = e.displayName;
                const val = document.createElement("span");
                val.textContent = e.value;
                row.appendChild(lbl);
                row.appendChild(val);
                this.tooltipEl.appendChild(row);
            });
        }

        this.positionTooltip(event);
        this.tooltipEl.classList.add("visible");
    }

    private mkTooltipRow(label: string, value: string): HTMLDivElement {
        const row = this.mkDiv("tooltip-row");
        const lbl = this.mkDiv("tooltip-row-label");
        lbl.textContent = label;
        const val = this.mkDiv("tooltip-row-value");
        val.textContent = value;
        row.appendChild(lbl);
        row.appendChild(val);
        return row;
    }

    private positionTooltip(event: MouseEvent): void {
        const gap   = 12;
        const rootR = this.root.getBoundingClientRect();
        let x = event.clientX - rootR.left + gap;
        let y = event.clientY - rootR.top  + gap;
        if (x + 200 > rootR.width)  { x = event.clientX - rootR.left - 200 - gap; }
        if (y + 140 > rootR.height) { y = event.clientY - rootR.top  - 140 - gap; }
        this.tooltipEl.style.cssText = `position:absolute;left:${x}px;top:${y}px;`;
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
