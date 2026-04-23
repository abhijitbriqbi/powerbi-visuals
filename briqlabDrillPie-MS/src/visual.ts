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
interface DataPoint {
    name:         string;
    value:        number;
    pct:          number;
    color:        string;
    selectionId:  ISelectionId;
    tooltipExtra: { displayName: string; value: string }[];
}

interface DrillFrame {
    label:     string;
    data:      DataPoint[];
    parentArc: { startAngle: number; endAngle: number; color: string } | null;
}

// ── Constants ──────────────────────────────────────────────────────────────────
const TRIAL_KEY     = "briqlab_trial_drillpie_start";
const PRO_STORE_KEY = "briqlab_drillpie_prokey";
const TRIAL_MS      = 4 * 24 * 60 * 60 * 1000;

// ── Helpers ────────────────────────────────────────────────────────────────────
function fmtVal(v: number): string {
    const a = Math.abs(v);
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

function midAngle(d: d3.PieArcDatum<DataPoint>): number {
    return d.startAngle + (d.endAngle - d.startAngle) / 2;
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
    private readonly svg:        d3.Selection<SVGSVGElement, unknown, null, undefined>;
    private readonly gParentArc: d3.Selection<SVGGElement, unknown, null, undefined>;
    private readonly gArcs:      d3.Selection<SVGGElement, unknown, null, undefined>;
    private readonly gLabels:    d3.Selection<SVGGElement, unknown, null, undefined>;
    private readonly gCenter:    d3.Selection<SVGGElement, unknown, null, undefined>;

    // State
    private settings!:   VisualFormattingSettingsModel;
    private viewport:    powerbi.IViewport = { width: 300, height: 300 };
    private currentData: DataPoint[]       = [];
    private selectedIds: Set<string>       = new Set();
    private drillStack:  DrillFrame[]      = [];

    // Arc generators (recalculated on each render)
    private arcGen!: d3.Arc<unknown, d3.PieArcDatum<DataPoint>>;
    private arcHov!: d3.Arc<unknown, d3.PieArcDatum<DataPoint>>;
    private outerR   = 100;
    private innerR   = 0;
    private cx       = 0;
    private cy       = 0;

    // Trial / Pro
    private isPro  = false;
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
        this.root.classList.add("briqlab-drill-pie");

        // Visual content (gets blurred on trial expiry)
        this.contentEl = this.mkDiv("briqlab-visual-content");
        this.root.appendChild(this.contentEl);

        // Breadcrumb
        this.breadcrumbEl = this.mkDiv("briqlab-breadcrumb hidden");
        this.contentEl.appendChild(this.breadcrumbEl);

        // Chart area
        this.chartAreaEl = this.mkDiv("briqlab-chart-area");
        this.contentEl.appendChild(this.chartAreaEl);

        // Drill-up button
        this.drillUpBtn = document.createElement("button");
        this.drillUpBtn.className = "briqlab-drill-up hidden";
        this.drillUpBtn.textContent = "\u2190 Drill Up";
        this.drillUpBtn.addEventListener("click", () => this.drillUp());
        this.chartAreaEl.appendChild(this.drillUpBtn);

        // SVG
        this.svg = d3.select(this.chartAreaEl)
            .append<SVGSVGElement>("svg")
            .attr("class", "briqlab-svg");

        this.gParentArc = this.svg.append("g").attr("class", "g-parent-arc");
        this.gArcs      = this.svg.append("g").attr("class", "g-arcs");
        this.gLabels    = this.svg.append("g").attr("class", "g-labels");
        this.gCenter    = this.svg.append("g").attr("class", "g-center");

        // Double-click SVG background → drill up
        this.svg.on("dblclick", () => { if (this.drillStack.length > 0) this.drillUp(); });
        // Single click on SVG background → clear selection
        this.svg.on("click", (event: MouseEvent) => {
            const t = event.target as SVGElement;
            if (t === this.svg.node() || t.tagName === "svg") {
                this.selectedIds.clear();
                this.selMgr.clear();
                this.applySelectionState();
            }
        });

        // Legend
        this.legendEl = this.mkDiv("briqlab-legend legend-bottom");
        this.contentEl.appendChild(this.legendEl);

        // Tooltip (outside contentEl — won't be blurred by CSS)
        this.tooltipEl = this.mkDiv("briqlab-tooltip");
        this.root.appendChild(this.tooltipEl);

        // Badges
        this.trialBadge = this.mkDiv("briqlab-trial-badge hidden");
        this.root.appendChild(this.trialBadge);

        this.proBadge = this.mkDiv("briqlab-pro-badge hidden");
        this.root.appendChild(this.proBadge);

        this.keyErrorEl = this.mkDiv("briqlab-key-error hidden");
        this.keyErrorEl.textContent = "\u2717 Invalid key";
        this.root.appendChild(this.keyErrorEl);

        // Trial overlay (outside contentEl — not blurred)
        this.overlayEl = this.mkDiv("briqlab-trial-overlay hidden");
        this.buildTrialOverlay();
        this.root.appendChild(this.overlayEl);

        // Restore any previously stored Pro key
        this.restoreProKey();
    }

    // ── Build trial overlay (DOM only, no innerHTML) ───────────────────────────
    private buildTrialOverlay(): void {
        const card = this.mkDiv("briqlab-trial-card");

        // Lock icon via SVG
        const iconSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        iconSvg.setAttribute("viewBox", "0 0 56 56");
        iconSvg.setAttribute("fill", "none");
        iconSvg.classList.add("trial-icon");
        this.buildPieIconSvg(iconSvg, "#0D9488", 0.5);
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

    // ── Build a simple SVG pie icon (safe, no innerHTML) ──────────────────────
    private buildPieIconSvg(
        svgEl: SVGSVGElement,
        color: string,
        opacity: number
    ): void {
        const ns = "http://www.w3.org/2000/svg";

        const circle = document.createElementNS(ns, "circle");
        circle.setAttribute("cx", "28"); circle.setAttribute("cy", "28"); circle.setAttribute("r", "22");
        circle.setAttribute("stroke", color); circle.setAttribute("stroke-width", "2.5");
        circle.setAttribute("opacity", opacity.toString());
        svgEl.appendChild(circle);

        const makeL = (x2: string, y2: string, col: string) => {
            const line = document.createElementNS(ns, "line");
            line.setAttribute("x1", "28"); line.setAttribute("y1", "28");
            line.setAttribute("x2", x2);   line.setAttribute("y2", y2);
            line.setAttribute("stroke", col); line.setAttribute("stroke-width", "2.5");
            line.setAttribute("stroke-linecap", "round");
            line.setAttribute("opacity", opacity.toString());
            svgEl.appendChild(line);
        };
        makeL("28", "6", color);
        makeL("47", "38", color);
        makeL("9",  "38", "#F97316");
    }

    // ── Build empty state (DOM only) ──────────────────────────────────────────
    private buildEmptyState(): HTMLDivElement {
        const wrap = this.mkDiv("briqlab-empty");

        const iconSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        iconSvg.setAttribute("viewBox", "0 0 56 56");
        iconSvg.setAttribute("fill", "none");
        iconSvg.classList.add("empty-icon");
        this.buildPieIconSvg(iconSvg, "#0D9488", 0.4);
        wrap.appendChild(iconSvg);

        const t = document.createElement("p");
        t.className = "empty-title";
        t.textContent = "Connect your data";
        wrap.appendChild(t);

        const b = document.createElement("p");
        b.className = "empty-body";
        b.textContent = "Add a Category and Values field to get started";
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
                        dataItems: [{ displayName: "Briqlab Drill Down Pie", value: "" }],
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
    
                // Remove any empty-state placeholder
                const emptyEl = this.chartAreaEl.querySelector(".briqlab-empty");
                if (emptyEl) { emptyEl.remove(); }
    
                this.render();
    
            } catch (err) {
                console.error("[BriqlabDrillPie] update error:", err);
            }
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    // ── Data parsing ───────────────────────────────────────────────────────────
    private parseData(dv: DataView): DataPoint[] {
        const cats    = dv.categorical!.categories![0];
        const vals    = dv.categorical!.values!;
        const measure = vals[0];
        const extras  = vals.slice(1);

        const rawVals = measure.values as number[];
        const total   = rawVals.reduce((a, v) => a + (v ?? 0), 0);
        const palette = this.buildPalette();

        return cats.values
            .map((name, i) => {
                const val   = (rawVals[i] as number) ?? 0;
                const extra = extras.map(col => ({
                    displayName: col.source.displayName || "",
                    value:       fmtVal((col.values[i] as number) ?? 0)
                }));
                return {
                    name:         (name ?? "(blank)").toString(),
                    value:        val,
                    pct:          total === 0 ? 0 : (val / total) * 100,
                    color:        palette[i % palette.length],
                    tooltipExtra: extra,
                    selectionId:  this.host.createSelectionIdBuilder()
                        .withCategory(cats, i)
                        .createSelectionId()
                };
            })
            .filter(d => d.value > 0);
    }

    private buildPalette(): string[] {
        const s = this.settings.segments;
        return [
            s.color1.value.value,  s.color2.value.value,  s.color3.value.value,
            s.color4.value.value,  s.color5.value.value,  s.color6.value.value,
            s.color7.value.value,  s.color8.value.value,  s.color9.value.value,
            s.color10.value.value
        ];
    }

    // ── Main render ────────────────────────────────────────────────────────────
    private render(): void {
        const { width, height } = this.viewport;
        const isTiny = width < 120 || height < 120;
        this.root.classList.toggle("tiny", isTiny);

        const legendPos   = this.settings.legendSettings.legendPosition.value.value as string;
        const showLegend  = this.settings.legendSettings.showLegend.value && legendPos !== "None";

        this.layoutLegend(legendPos, showLegend && !isTiny);

        // Reserve space for legend
        let legendH = 0;
        let legendW = 0;
        if (showLegend && !isTiny) {
            if (legendPos === "Left" || legendPos === "Right") legendW = 140;
            else legendH = 56;
        }

        const hasDrill = this.drillStack.length > 0;
        const bcH      = hasDrill && !isTiny ? 26 : 0;

        const chartW = width  - legendW;
        const chartH = height - legendH;

        this.svg.attr("width", chartW).attr("height", chartH);

        this.cx     = chartW / 2;
        this.cy     = bcH + (chartH - bcH) / 2;
        this.outerR = Math.min(chartW, chartH - bcH) * 0.38;

        const isDonut   = this.settings.chartStyle.chartType.value.value === "Donut";
        const innerPct  = Math.max(0.40, Math.min(0.70, this.settings.chartStyle.innerRadius.value / 100));
        this.innerR     = isDonut ? this.outerR * innerPct : 0;
        const padRad    = (this.settings.chartStyle.padAngle.value * Math.PI) / 1000;

        this.arcGen = d3.arc<d3.PieArcDatum<DataPoint>>()
            .innerRadius(this.innerR)
            .outerRadius(this.outerR)
            .cornerRadius(this.innerR > 0 ? 4 : 2);

        this.arcHov = d3.arc<d3.PieArcDatum<DataPoint>>()
            .innerRadius(this.innerR)
            .outerRadius(this.outerR + 10)
            .cornerRadius(this.innerR > 0 ? 4 : 2);

        this.gArcs.attr("transform",      `translate(${this.cx},${this.cy})`);
        this.gLabels.attr("transform",    `translate(${this.cx},${this.cy})`);
        this.gCenter.attr("transform",    `translate(${this.cx},${this.cy})`);
        this.gParentArc.attr("transform", `translate(${this.cx},${this.cy})`);

        const pieGen = d3.pie<DataPoint>()
            .value(d => d.value)
            .padAngle(padRad)
            .sort(null);

        const arcData = pieGen(this.currentData);

        this.renderArcs(arcData);
        this.renderParentArc();

        if (this.settings.labelSettings.showLabels.value && !isTiny) {
            this.renderLabels(arcData);
        } else {
            this.gLabels.selectAll("*").remove();
        }

        if (isDonut && this.settings.centerDisplay.showCenter.value) {
            const total = this.currentData.reduce((a, d) => a + d.value, 0);
            this.renderCenter(total, null);
        } else {
            this.gCenter.selectAll("*").remove();
        }

        this.renderBreadcrumb();

        if (showLegend && !isTiny) {
            this.renderLegend();
        } else {
            this.clearEl(this.legendEl);
        }
    }

    // ── Arc rendering ──────────────────────────────────────────────────────────
    private renderArcs(arcData: d3.PieArcDatum<DataPoint>[]): void {
        const firstRender = this.gArcs.selectAll(".arc-path").size() === 0;

        const paths = this.gArcs
            .selectAll<SVGPathElement, d3.PieArcDatum<DataPoint>>(".arc-path")
            .data(arcData, d => d.data.name);

        paths.exit()
            .transition().duration(dur(300))
            .style("opacity", 0)
            .remove();

        const entered = paths.enter()
            .append<SVGPathElement>("path")
            .attr("class", "arc-path")
            .attr("fill",         d => d.data.color)
            .attr("stroke",       "#ffffff")
            .attr("stroke-width", "1.5")
            .style("opacity", firstRender ? "0" : "1")
            .on("click",      (event: MouseEvent, d) => this.onSegmentClick(event, d))
            .on("contextmenu", (event: MouseEvent, d) => {
                event.preventDefault();
                event.stopPropagation();
                this.selMgr.showContextMenu(
                    d.data.selectionId,
                    { x: event.clientX, y: event.clientY }
                );
            })
            .on("mouseenter", (event: MouseEvent, d) => this.onSegmentHover(event, d))
            .on("mouseleave", (event: MouseEvent, d) => this.onSegmentLeave(event, d));

        const merged = entered.merge(paths);
        merged.attr("fill", d => d.data.color);

        if (firstRender) {
            // Sweep-in from 12-o'clock
            merged
                .attr("d", d => {
                    const zero = { ...d, startAngle: 0, endAngle: 0 };
                    return this.arcGen(zero) ?? "";
                })
                .transition().duration(0).style("opacity", "1");

            merged.each((d, i, nodes) => {
                const node = nodes[i] as SVGPathElement;
                const capturedD = d;
                const capturedI = i;
                d3.select(node)
                    .transition()
                    .delay(dur(capturedI * 50))
                    .duration(dur(600))
                    .ease(d3.easeCubicOut)
                    .attrTween("d", () => {
                        const from = { ...capturedD, startAngle: 0, endAngle: 0 };
                        const interp = d3.interpolate(from, capturedD);
                        return (t: number) => this.arcGen(interp(t)) ?? "";
                    })
                    .on("end", function(this: SVGPathElement) {
                        (this as SVGPathElement & { _prev?: d3.PieArcDatum<DataPoint> })._prev = capturedD;
                    });
            });
        } else {
            // Morph animation for data updates
            const capturedInnerR = this.innerR;
            const capturedOuterR = this.outerR;
            merged
                .style("opacity", "1")
                .transition()
                .duration(dur(500))
                .ease(d3.easeCubicOut)
                .attrTween("d", function(this: SVGPathElement, d: d3.PieArcDatum<DataPoint>) {
                    const node   = this as SVGPathElement & { _prev?: d3.PieArcDatum<DataPoint> };
                    const prev   = node._prev ?? d;
                    const interp = d3.interpolate(prev, d);
                    node._prev   = d;
                    const snap   = d3.arc<d3.PieArcDatum<DataPoint>>()
                        .innerRadius(capturedInnerR)
                        .outerRadius(capturedOuterR)
                        .cornerRadius(4);
                    return (t: number) => snap(interp(t)) ?? "";
                });
        }

        this.applySelectionState();
    }

    // ── Parent arc ring (drill context) ───────────────────────────────────────
    private renderParentArc(): void {
        this.gParentArc.selectAll("*").remove();
        if (this.drillStack.length === 0) return;

        const frame = this.drillStack[this.drillStack.length - 1];
        if (!frame.parentArc) return;

        const rOut = this.outerR + 18;
        const rIn  = this.outerR + 4;

        const bgArc = d3.arc<{ startAngle: number; endAngle: number }>()
            .innerRadius(rIn).outerRadius(rOut).cornerRadius(2);

        this.gParentArc.append("path")
            .datum({ startAngle: 0, endAngle: Math.PI * 2 })
            .attr("fill", "#E2E8F0")
            .attr("d", d => bgArc(d) ?? "");

        this.gParentArc.append("path")
            .datum(frame.parentArc)
            .attr("class", "parent-arc")
            .attr("fill",  frame.parentArc.color)
            .attr("d",     d => bgArc(d) ?? "")
            .on("click",   () => this.drillUp());
    }

    // ── Labels ─────────────────────────────────────────────────────────────────
    private renderLabels(arcData: d3.PieArcDatum<DataPoint>[]): void {
        this.gLabels.selectAll("*").remove();

        const threshold = this.settings.labelSettings.labelThreshold.value;
        const fontSize  = this.settings.labelSettings.labelFontSize.value;
        const fmtMode   = this.settings.labelSettings.labelFormat.value.value as string;
        const labelR    = this.outerR + 28;

        arcData.forEach(d => {
            if (d.data.pct < threshold) return;

            const angle = midAngle(d) - Math.PI / 2;
            const right = Math.cos(angle) >= 0;

            const innerPt = [Math.cos(angle) * (this.outerR + 4), Math.sin(angle) * (this.outerR + 4)];
            const bendPt  = [Math.cos(angle) * (labelR - 8),      Math.sin(angle) * (labelR - 8)];
            const endPt   = [right ? bendPt[0] + 12 : bendPt[0] - 12, bendPt[1]];

            this.gLabels.append("polyline")
                .attr("class", "label-connector")
                .attr("points", `${innerPt[0]},${innerPt[1]} ${bendPt[0]},${bendPt[1]} ${endPt[0]},${endPt[1]}`);

            const short  = truncate(d.data.name, 16);
            const pctStr = d.data.pct.toFixed(1) + "%";
            let text = "";
            if (fmtMode === "Name (%)") text = `${short} (${pctStr})`;
            else if (fmtMode === "% only") text = pctStr;
            else text = short;

            this.gLabels.append("text")
                .attr("class",       "segment-label")
                .attr("x",           endPt[0])
                .attr("y",           endPt[1])
                .attr("text-anchor", right ? "start" : "end")
                .attr("dy",          "0.35em")
                .style("font-size",  `${fontSize}px`)
                .text(text);
        });
    }

    // ── Center display (donut mode) ────────────────────────────────────────────
    private renderCenter(total: number, hovered: DataPoint | null): void {
        const vSize = this.settings.centerDisplay.centerValueSize.value;
        const lSize = this.settings.centerDisplay.centerLabelSize.value;
        const label = this.settings.centerDisplay.centerLabel.value;

        this.gCenter.selectAll("*").remove();

        if (hovered) {
            this.gCenter.append("text")
                .attr("class", "center-segment-name")
                .attr("y",     -(vSize * 0.75))
                .style("font-size", `${lSize + 1}px`)
                .text(truncate(hovered.name, 20));

            this.gCenter.append("text")
                .attr("class", "center-value")
                .attr("y",     0)
                .style("font-size", `${vSize}px`)
                .text(fmtVal(hovered.value));

            this.gCenter.append("text")
                .attr("class", "center-pct")
                .attr("y",     vSize * 0.75)
                .style("font-size", `${lSize}px`)
                .text(`${hovered.pct.toFixed(1)}% of total`);
        } else {
            this.gCenter.append("text")
                .attr("class", "center-label")
                .attr("y",     -(vSize * 0.45))
                .style("font-size", `${lSize}px`)
                .text(label);

            this.gCenter.append("text")
                .attr("class", "center-value")
                .attr("y",     vSize * 0.3)
                .style("font-size", `${vSize}px`)
                .text(fmtVal(total));
        }
    }

    // ── Legend ─────────────────────────────────────────────────────────────────
    private renderLegend(): void {
        this.clearEl(this.legendEl);
        const fontSize = this.settings.legendSettings.legendFontSize.value;
        this.legendEl.style.fontSize = `${fontSize}px`;

        this.currentData.forEach(d => {
            const item   = this.mkDiv("briqlab-legend-item");
            const swatch = this.mkDiv("legend-swatch");
            swatch.style.background = d.color;

            const nameEl = this.mkDiv("legend-name");
            nameEl.textContent = truncate(d.name, 20);

            const valEl = this.mkDiv("legend-value");
            valEl.textContent = fmtVal(d.value);

            item.appendChild(swatch);
            item.appendChild(nameEl);
            item.appendChild(valEl);

            item.addEventListener("click", (e: MouseEvent) => {
                // Create a minimal PieArcDatum for the click handler
                const fakeDatum = {
                    data: d, value: d.value,
                    index: 0, startAngle: 0, endAngle: 0, padAngle: 0
                } as d3.PieArcDatum<DataPoint>;
                this.onSegmentClick(e, fakeDatum);
            });

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

    // ── Breadcrumb ────────────────────────────────────────────────────────────
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
                const captureDepth = depth;
                span.addEventListener("click", () => this.drillUpTo(captureDepth));
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

    // ── Empty state ────────────────────────────────────────────────────────────
    private renderEmpty(): void {
        this.gArcs.selectAll("*").remove();
        this.gLabels.selectAll("*").remove();
        this.gCenter.selectAll("*").remove();
        this.gParentArc.selectAll("*").remove();
        this.clearEl(this.legendEl);
        this.breadcrumbEl.classList.add("hidden");
        this.drillUpBtn.classList.add("hidden");

        if (!this.chartAreaEl.querySelector(".briqlab-empty")) {
            this.chartAreaEl.appendChild(this.buildEmptyState());
        }
    }

    // ── Interactions ───────────────────────────────────────────────────────────
    private onSegmentClick(event: MouseEvent, d: d3.PieArcDatum<DataPoint>): void {
        try {
            const idKey       = this.selKey(d.data);
            const isMulti     = event.ctrlKey || event.metaKey;
            const wasSelected = this.selectedIds.has(idKey);

            if (isMulti) {
                wasSelected ? this.selectedIds.delete(idKey) : this.selectedIds.add(idKey);
            } else {
                if (wasSelected && this.selectedIds.size === 1) {
                    this.selectedIds.clear();
                } else {
                    this.selectedIds.clear();
                    this.selectedIds.add(idKey);
                }
            }

            if (this.selectedIds.size > 0) {
                const ids = this.currentData
                    .filter(pt => this.selectedIds.has(this.selKey(pt)))
                    .map(pt => pt.selectionId);
                this.selMgr.select(ids, isMulti).then(() => this.applySelectionState());
            } else {
                this.selMgr.clear().then(() => this.applySelectionState());
            }

            this.applySelectionState();
            event.stopPropagation();
        } catch (err) {
            console.warn("[BriqlabDrillPie] click:", err);
        }
    }

    private onSegmentHover(event: MouseEvent, d: d3.PieArcDatum<DataPoint>): void {
        const target = event.currentTarget as SVGPathElement;
        d3.select(target)
            .transition().duration(dur(200))
            .attr("d", this.arcHov(d) ?? "");

        const isDonut = this.settings.chartStyle.chartType.value.value === "Donut";
        if (isDonut && this.settings.centerDisplay.showCenter.value) {
            const total = this.currentData.reduce((a, p) => a + p.value, 0);
            this.renderCenter(total, d.data);
        }

        if (this.settings.tooltipSettings.showTooltip.value) {
            if (this.tooltipHideTimer) { clearTimeout(this.tooltipHideTimer); this.tooltipHideTimer = null; }
            this.tooltipShowTimer = setTimeout(() => this.showTooltip(event, d.data), 150);
        }
    }

    private onSegmentLeave(event: MouseEvent, d: d3.PieArcDatum<DataPoint>): void {
        const target = event.currentTarget as SVGPathElement;
        d3.select(target)
            .transition().duration(dur(200))
            .attr("d", this.arcGen(d) ?? "");

        const isDonut = this.settings.chartStyle.chartType.value.value === "Donut";
        if (isDonut && this.settings.centerDisplay.showCenter.value) {
            const total = this.currentData.reduce((a, p) => a + p.value, 0);
            this.renderCenter(total, null);
        }

        if (this.tooltipShowTimer) { clearTimeout(this.tooltipShowTimer); this.tooltipShowTimer = null; }
        this.tooltipHideTimer = setTimeout(() => this.tooltipEl.classList.remove("visible"), 100);
    }

    private applySelectionState(): void {
        const hasSel = this.selectedIds.size > 0;
        this.gArcs.selectAll<SVGPathElement, d3.PieArcDatum<DataPoint>>(".arc-path")
            .style("opacity", d => {
                if (!hasSel) return "1";
                return this.selectedIds.has(this.selKey(d.data)) ? "1" : "0.4";
            })
            .style("transform", d => {
                return hasSel && this.selectedIds.has(this.selKey(d.data))
                    ? "scale(1.04)" : "scale(1)";
            });
    }

    private selKey(d: DataPoint): string {
        return ((d.selectionId as unknown as Record<string, unknown>)["key"] as string) || d.name;
    }

    // ── Drill navigation ───────────────────────────────────────────────────────
    private drillUp(): void {
        if (!this.drillStack.length) return;
        const frame      = this.drillStack.pop()!;
        this.currentData = frame.data;
        this.selectedIds.clear();
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
        this.render();
    }

    // ── Tooltip (no innerHTML) ────────────────────────────────────────────────
    private showTooltip(event: MouseEvent, d: DataPoint): void {
        this.clearEl(this.tooltipEl);

        const total  = this.currentData.reduce((a, p) => a + p.value, 0);
        const pctStr = total > 0 ? (d.value / total * 100).toFixed(1) + "%" : "0.0%";

        // Header
        const header = this.mkDiv("tooltip-header");
        const dot    = this.mkDiv("tooltip-dot");
        dot.style.background = d.color;
        const nameEl = this.mkDiv("tooltip-name");
        nameEl.textContent = d.name;
        header.appendChild(dot);
        header.appendChild(nameEl);
        this.tooltipEl.appendChild(header);

        // Value row
        this.tooltipEl.appendChild(this.mkTooltipRow("Value", fmtVal(d.value)));
        this.tooltipEl.appendChild(this.mkTooltipRow("% of total", pctStr));

        // Extra tooltip fields
        if (d.tooltipExtra.length) {
            const div = this.mkDiv("tooltip-divider");
            this.tooltipEl.appendChild(div);
            d.tooltipExtra.forEach(e => {
                const row  = this.mkDiv("tooltip-extra-row");
                const lbl  = document.createElement("span");
                lbl.textContent = e.displayName;
                const val  = document.createElement("span");
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
        const row  = this.mkDiv("tooltip-row");
        const lbl  = this.mkDiv("tooltip-row-label");
        lbl.textContent = label;
        const val  = this.mkDiv("tooltip-row-value");
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
        if (x + 200 > rootR.width)  x = event.clientX - rootR.left - 200 - gap;
        if (y + 120 > rootR.height) y = event.clientY - rootR.top  - 120 - gap;
        this.tooltipEl.style.cssText =
            `position:absolute;left:${x}px;top:${y}px;`;
    }

    // ── Trial / Pro key ────────────────────────────────────────────────────────
    private getTrialStatus(): { active: boolean; daysLeft: number; expired: boolean } {
        try {
            let raw = localStorage.getItem(TRIAL_KEY);
            if (!raw) { raw = Date.now().toString(); localStorage.setItem(TRIAL_KEY, raw); }
            const elapsed  = Date.now() - parseInt(raw, 10);
            const daysLeft = Math.max(0, Math.ceil((TRIAL_MS - elapsed) / 86400000));
            return { active: elapsed <= TRIAL_MS, daysLeft, expired: elapsed > TRIAL_MS };
        } catch {
            return { active: true, daysLeft: 4, expired: false };
        }
    }

    private handleProKey(): void {
        checkMicrosoftLicence(this.host).then(p => this._msUpdateLicenceUI(p)).catch(() => this._msUpdateLicenceUI(false));
    }

    private async validateKey(key: string): Promise<boolean> {
        return checkMicrosoftLicence(this.host);
    }

    private restoreProKey(): void {
        try {
            const stored = localStorage.getItem(PRO_STORE_KEY);
            if (!stored) return;
            const parsed = JSON.parse(stored) as { key?: string };
            const key    = parsed?.key;
            if (!key) return;
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

    // ── DOM utilities ─────────────────────────────────────────────────────────
    private mkDiv(cls: string): HTMLDivElement {
        const el = document.createElement("div");
        el.className = cls;
        return el;
    }

    private clearEl(el: HTMLElement): void {
        while (el.firstChild) { el.removeChild(el.firstChild); }
    }
}
