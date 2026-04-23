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
import { getPurchaseUrl, getButtonText } from "./trialManager";

// ── Types ──────────────────────────────────────────────────────────────────────
interface DataPoint {
    name:        string;
    value:       number;
    pct:         number;
    color:       string;
    selectionId: ISelectionId;
}

// ── Constants ──────────────────────────────────────────────────────────────────
const DEFAULT_PALETTE = [
    "#0D9488", "#F97316", "#3B82F6", "#8B5CF6", "#10B981",
    "#EF4444", "#F59E0B", "#EC4899", "#06B6D4", "#84CC16"
];

const TRIAL_MS        = 4 * 24 * 60 * 60 * 1000;
const LS_TRIAL_START  = "briqlab_trial_piechart_start";
const LS_PRO_KEY      = "briqlab_piechart_prokey";

const LEGEND_SIDE_W   = 175;
const HOVER_EXPAND    = 10;
const MAX_LABEL_CHARS = 14;

// ── Helpers ────────────────────────────────────────────────────────────────────
function clamp(v: number, lo: number, hi: number): number {
    return Math.max(lo, Math.min(hi, v));
}

function fmt(v: number): string {
    const a = Math.abs(v);
    if (a >= 1e6) { const r = (v / 1e6).toFixed(1); return (r.endsWith(".0") ? r.slice(0, -2) : r) + "M"; }
    if (a >= 1e3) { const r = (v / 1e3).toFixed(1); return (r.endsWith(".0") ? r.slice(0, -2) : r) + "K"; }
    return v.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

function truncate(s: string, max: number): string {
    return s.length > max ? s.slice(0, max - 1) + "\u2026" : s;
}

function makeLabelText(dp: DataPoint, format: string): string {
    const name = truncate(dp.name, MAX_LABEL_CHARS);
    const pctStr = `${dp.pct.toFixed(1)}%`;
    const valStr = fmt(dp.value);
    switch (format) {
        case "name_value": return `${name}: ${valStr}`;
        case "pct":        return pctStr;
        case "name":       return name;
        case "value":      return valStr;
        default:           return `${name}: ${pctStr}`;
    }
}

function getTrialDaysRemaining(): number {
    try {
        const raw = localStorage.getItem(LS_TRIAL_START);
        if (!raw) {
            localStorage.setItem(LS_TRIAL_START, String(Date.now()));
            return 4;
        }
        const start = parseInt(raw, 10);
        if (isNaN(start)) return 0;
        const elapsed = Date.now() - start;
        const days    = Math.floor(elapsed / (24 * 60 * 60 * 1000));
        return Math.max(0, 4 - days);
    } catch {
        return 4;
    }
}

function isTrialExpired(): boolean {
    try {
        const raw = localStorage.getItem(LS_TRIAL_START);
        if (!raw) return false;
        const start = parseInt(raw, 10);
        if (isNaN(start)) return false;
        return (Date.now() - start) >= TRIAL_MS;
    } catch {
        return false;
    }
}

// ── Visual ─────────────────────────────────────────────────────────────────────
export class Visual implements IVisual {
    private readonly host:   IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private readonly selMgr: ISelectionManager;
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private _handlersAttached = false;
    private readonly fmtSvc: FormattingSettingsService;

    // DOM
    private readonly root:          HTMLElement;
    private readonly visualContent: HTMLElement;
    private readonly chartArea:     HTMLElement;
    private readonly legendEl:      HTMLElement;
    private readonly trialBadge:    HTMLElement;
    private readonly proBadge:      HTMLElement;
    private readonly keyError:      HTMLElement;
    private readonly trialOverlay:  HTMLElement;
    private readonly svg: d3.Selection<SVGSVGElement, unknown, null, undefined>;

    // State
    private settings!: VisualFormattingSettingsModel;
    private vp:   powerbi.IViewport    = { width: 300, height: 300 };
    private pts:  DataPoint[]          = [];
    private total: number              = 0;

    // Pro key
    private curKey:   string              = "";
    private isPro:    boolean             = false;
    private keyCache: Map<string, boolean> = new Map();

    // ── Constructor ────────────────────────────────────────────────────────────
    constructor(options: VisualConstructorOptions) {
        this.host   = options.host;
        this.renderingManager = options.host.eventService;
        this.selMgr = options.host.createSelectionManager();
        this.tooltipSvc = this.host.tooltipService;
        this.fmtSvc = new FormattingSettingsService();

        this.root = options.element;
        this.root.classList.add("briqlab-pie");

        this.visualContent = document.createElement("div");
        this.visualContent.className = "briqlab-visual-content";
        this.root.appendChild(this.visualContent);

        this.chartArea = document.createElement("div");
        this.chartArea.className = "pie-chart-area";
        this.visualContent.appendChild(this.chartArea);

        this.legendEl = document.createElement("div");
        this.legendEl.className = "pie-legend";
        this.visualContent.appendChild(this.legendEl);

        this.svg = d3.select(this.chartArea)
            .append<SVGSVGElement>("svg")
            .attr("class", "pie-svg");

        this.trialBadge = document.createElement("div");
        this.trialBadge.className = "briqlab-trial-badge hidden";
        this.root.appendChild(this.trialBadge);

        this.proBadge = document.createElement("div");
        this.proBadge.className = "briqlab-pro-badge hidden";
        this.proBadge.textContent = "\u2713 Pro Active";
        this.root.appendChild(this.proBadge);

        this.keyError = document.createElement("div");
        this.keyError.className = "briqlab-key-error hidden";
        this.keyError.textContent = "\u2717 Invalid key";
        this.root.appendChild(this.keyError);

        this.trialOverlay = document.createElement("div");
        this.trialOverlay.className = "briqlab-trial-overlay hidden";
        this.root.appendChild(this.trialOverlay);

        const card = document.createElement("div");
        card.className = "briqlab-trial-card";
        this.trialOverlay.appendChild(card);

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
        btn.addEventListener("click", () => {
            this.host.launchUrl(getPurchaseUrl());
        });
        card.appendChild(btn);

        const sub = document.createElement("p");
        sub.className = "trial-subtext";
        sub.textContent = "Purchase on Microsoft AppSource to unlock all features instantly.";
        card.appendChild(sub);

        try {
            const stored = localStorage.getItem(LS_PRO_KEY);
            if (stored) {
                this.curKey = stored;
                this.validateKey(stored).then(ok => {
                    this.isPro = ok;
                    this.render();
                });
            }
        } catch {
            // localStorage unavailable
        }
    }

    // ── Pro key validation ─────────────────────────────────────────────────────
    private async validateKey(key: string): Promise<boolean> {
        return checkMicrosoftLicence(this.host);
    }

    // ── Colour palette from settings ───────────────────────────────────────────
    private getPalette(): string[] {
        const cs = this.settings.colorSettings;
        return [
            cs.color1.value?.value  ?? DEFAULT_PALETTE[0],
            cs.color2.value?.value  ?? DEFAULT_PALETTE[1],
            cs.color3.value?.value  ?? DEFAULT_PALETTE[2],
            cs.color4.value?.value  ?? DEFAULT_PALETTE[3],
            cs.color5.value?.value  ?? DEFAULT_PALETTE[4],
            cs.color6.value?.value  ?? DEFAULT_PALETTE[5],
            cs.color7.value?.value  ?? DEFAULT_PALETTE[6],
            cs.color8.value?.value  ?? DEFAULT_PALETTE[7],
            cs.color9.value?.value  ?? DEFAULT_PALETTE[8],
            cs.color10.value?.value ?? DEFAULT_PALETTE[9]
        ];
    }

    // ── update ─────────────────────────────────────────────────────────────────
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
                        dataItems: [{ displayName: "Briqlab Pie Chart", value: "" }],
                        identities: [],
                        coordinates: [e.clientX, e.clientY],
                        isTouchEvent: false
                    });
                });
                this.root.addEventListener("mouseleave", () => {
                    this.tooltipSvc.hide({ isTouchEvent: false, immediately: false });
                });
            }
            this.vp       = options.viewport;
            this.settings = this.fmtSvc.populateFormattingSettingsModel(
                VisualFormattingSettingsModel, options.dataViews?.[0]
            );
    
            this.pts   = [];
            this.total = 0;
            const dv   = options.dataViews?.[0];
    
            if (dv?.categorical?.categories?.length && dv.categorical.values?.length) {
                const cats   = dv.categorical.categories[0];
                const vals   = dv.categorical.values[0];
                const palette = this.getPalette();
                const sum    = (vals.values as number[]).reduce((a, v) => a + (v ?? 0), 0);
                this.total   = sum;
    
                this.pts = cats.values.map((name, i) => {
                    const val = (vals.values[i] as number) ?? 0;
                    return {
                        name:        (name ?? "(blank)").toString(),
                        value:       val,
                        pct:         sum === 0 ? 0 : (val / sum) * 100,
                        color:       palette[i % palette.length],
                        selectionId: this.host.createSelectionIdBuilder()
                            .withCategory(cats, i)
                            .createSelectionId()
                    };
                });
            }
    
            // Pro key
            const key = ""; // MS cert: pro key field removed
            if (key && key !== this.curKey) {
                this.curKey = key;
                this.isPro  = false;
                this.validateKey(key).then(ok => {
                    this.isPro = ok;
                    this.render();
                });
            } else if (!key && !this.curKey) {
                this.isPro = false;
            }
    
            this.render();
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    // ── render ─────────────────────────────────────────────────────────────────
    private render(): void {
        if (!this.settings) return;

        const { width, height } = this.vp;
        const s = this.settings;

        // ── Trial / Pro UI ───────────────────────────────────────────────────
        const expired = isTrialExpired();
        const days    = expired ? 0 : getTrialDaysRemaining();

        if (!this.isPro && !expired) {
            this.trialBadge.textContent = `Trial: ${days} day${days === 1 ? "" : "s"} remaining`;
            this.trialBadge.classList.remove("hidden");
        } else {
            this.trialBadge.classList.add("hidden");
        }

        if (this.isPro) {
            this.proBadge.classList.remove("hidden");
            this.keyError.classList.add("hidden");
        } else {
            this.proBadge.classList.add("hidden");
            const k = ""; // MS cert: pro key field removed
            if (k && this.keyCache.has(k) && !this.keyCache.get(k)) {
                this.keyError.classList.remove("hidden");
            } else {
                this.keyError.classList.add("hidden");
            }
        }

        if (expired && !this.isPro) {
            this.visualContent.classList.add("blurred");
            this.trialOverlay.classList.remove("hidden");
        } else {
            this.visualContent.classList.remove("blurred");
            this.trialOverlay.classList.add("hidden");
        }

        // ── Settings ──────────────────────────────────────────────────────────
        const donutMode     = s.pieSettings.donutMode.value;
        const innerRadiusPct = clamp(s.pieSettings.innerRadius.value, 10, 80) / 100;
        const showCenter    = s.pieSettings.showCenterText.value;
        const centerLabel   = s.pieSettings.centerLabel.value ?? "Total";
        const sortOrder     = String(s.pieSettings.sortOrder?.value?.value ?? "original");
        const startAngleDeg = s.pieSettings.startAngle.value ?? 0;
        const borderColor   = s.pieSettings.borderColor.value?.value ?? "#ffffff";
        const borderWidth   = clamp(s.pieSettings.borderWidth.value ?? 2, 0, 10);
        const minLabelPct   = clamp(s.pieSettings.minLabelPct.value ?? 4, 0, 50);

        const showLabels    = s.labelSettings.showLabels.value;
        const labelFontSize = clamp(s.labelSettings.labelFontSize.value, 8, 18);
        const labelFormat   = String(s.labelSettings.labelFormat?.value?.value ?? "name_pct");
        const labelColor    = s.labelSettings.labelColor?.value?.value ?? "#374151";
        const boldLabels    = s.labelSettings.boldLabels?.value ?? false;
        const fontFamily    = String(s.labelSettings.fontFamily?.value?.value ?? "Segoe UI");

        const showLegend    = s.legendSettings.showLegend.value;
        const rawPos        = String(s.legendSettings.legendPosition.value?.value ?? "Right").toLowerCase();
        const legendPos     = (["left", "top", "bottom", "right"] as const).includes(
            rawPos as "left" | "top" | "bottom" | "right"
        ) ? (rawPos as "left" | "top" | "bottom" | "right") : "right";
        const legendFontSize = clamp(s.legendSettings.legendFontSize?.value ?? 11, 8, 18);

        // ── Sort ──────────────────────────────────────────────────────────────
        let sorted = [...this.pts];
        if (sortOrder === "desc") sorted.sort((a, b) => b.value - a.value);
        else if (sortOrder === "asc") sorted.sort((a, b) => a.value - b.value);

        // ── Layout ────────────────────────────────────────────────────────────
        const hasLegend = showLegend && sorted.length > 0;
        const itemH     = legendFontSize * 2.2;
        const legendHMax = Math.min(sorted.length * itemH, height * 0.28);

        let chartW = width, chartH = height;
        let legW   = 0,     legH   = 0;

        if (hasLegend) {
            if (legendPos === "left" || legendPos === "right") {
                legW   = Math.min(LEGEND_SIDE_W, Math.floor(width * 0.35));
                chartW = width - legW;
            } else {
                legH   = legendHMax;
                chartH = height - legH;
            }
        }

        this.visualContent.style.cssText =
            `width:${width}px;height:${height}px;display:flex;` +
            `flex-direction:${legendPos === "left"   ? "row-reverse"    :
                              legendPos === "top"    ? "column-reverse" :
                              legendPos === "bottom" ? "column"         : "row"};`;

        this.chartArea.style.cssText =
            `width:${chartW}px;height:${chartH}px;flex-shrink:0;position:relative;`;

        if (hasLegend) {
            this.legendEl.style.cssText =
                `display:flex;flex-shrink:0;` +
                (legendPos === "left" || legendPos === "right"
                    ? `width:${legW}px;height:${chartH}px;`
                    : `width:${width}px;height:${legH}px;`);
        } else {
            this.legendEl.style.cssText = "display:none;";
        }

        // ── SVG ───────────────────────────────────────────────────────────────
        this.svg.attr("width", chartW).attr("height", chartH);
        this.svg.selectAll("*").remove();

        const cx = chartW / 2;
        const cy = chartH / 2;

        const outerR = Math.min(chartW, chartH) / 2 * (showLabels ? 0.65 : 0.82);
        const innerR = donutMode ? outerR * innerRadiusPct : 0;
        const hoverR = outerR + HOVER_EXPAND;

        // Convert startAngle degrees → radians (D3 pie uses 0 = top)
        const startAngleRad = (startAngleDeg * Math.PI) / 180;

        const arcGen = d3.arc<d3.DefaultArcObject>();

        const makeArc = (sa: number, ea: number, r: number): d3.DefaultArcObject => ({
            startAngle: sa, endAngle: ea,
            innerRadius: innerR, outerRadius: r, padAngle: 0
        });

        // ── Empty state ───────────────────────────────────────────────────────
        if (sorted.length === 0) {
            this.svg.append("circle")
                .attr("cx", cx).attr("cy", cy).attr("r", outerR)
                .attr("fill", "#F3F4F6");
            this.svg.append("text")
                .attr("x", cx).attr("y", cy)
                .attr("text-anchor", "middle")
                .attr("dominant-baseline", "middle")
                .attr("class", "pie-empty-msg")
                .text("Add Category & Value fields");
            this.buildLegend([], legendFontSize);
            return;
        }

        // ── Pie layout ────────────────────────────────────────────────────────
        const pie     = d3.pie<DataPoint>()
            .value(d => d.value)
            .sort(null)
            .padAngle(0)
            .startAngle(startAngleRad);
        const pieData = pie(sorted);

        // ── Segments ──────────────────────────────────────────────────────────
        const arcsG = this.svg.append("g")
            .attr("class", "pie-arcs")
            .attr("transform", `translate(${cx},${cy})`);

        const segs = arcsG
            .selectAll<SVGPathElement, d3.PieArcDatum<DataPoint>>("path")
            .data(pieData)
            .join("path")
            .attr("class",        "pie-seg")
            .attr("fill",         d => d.data.color)
            .attr("stroke",       borderColor)
            .attr("stroke-width", borderWidth)
            .attr("d", d => arcGen(makeArc(d.startAngle, d.startAngle, outerR)) ?? "");

        // Sweep-in animation
        segs.transition()
            .duration(650)
            .ease(d3.easeCubicOut)
            .attrTween("d", function(d) {
                const interp = d3.interpolate(
                    makeArc(d.startAngle, d.startAngle, outerR),
                    makeArc(d.startAngle, d.endAngle,   outerR)
                );
                return (t: number) => arcGen(interp(t)) ?? "";
            });

        // ── Centre hub (donut mode) ────────────────────────────────────────────
        const hubFS  = clamp(innerR * 0.50, 9, 22);
        const labFS2 = clamp(innerR * 0.28, 7, 13);

        const showHub = donutMode && showCenter && innerR > 20;
        const hubG = showHub
            ? this.svg.append("g").attr("transform", `translate(${cx},${cy})`)
            : null;

        if (hubG) {
            hubG.append("circle").attr("r", innerR).attr("class", "pie-hub-bg");
        }

        const hubValEl = hubG
            ? hubG.append("text")
                .attr("class",       "pie-hub-val")
                .attr("text-anchor", "middle")
                .attr("y",           hubFS * 0.35)
                .style("font-size",  `${hubFS}px`)
                .style("font-family", fontFamily)
                .text(fmt(this.total))
            : null;

        const hubLabEl = hubG
            ? hubG.append("text")
                .attr("class",       "pie-hub-lab")
                .attr("text-anchor", "middle")
                .attr("y",           hubFS * 0.35 + labFS2 + 2)
                .style("font-size",  `${labFS2}px`)
                .style("font-family", fontFamily)
                .text(centerLabel)
            : null;

        // ── Labels + connector lines ──────────────────────────────────────────
        if (showLabels) {
            const labG = this.svg.append("g")
                .attr("class", "pie-labels")
                .attr("transform", `translate(${cx},${cy})`);

            const midR  = outerR * 0.75;
            const bendR = outerR * 1.06;
            const endR  = outerR * 1.14;
            const textR = outerR * 1.18;

            pieData.forEach(d => {
                if (d.data.pct < minLabelPct) return;

                const mid     = (d.startAngle + d.endAngle) / 2;
                const sinMid  = Math.sin(mid);
                const cosMid  = Math.cos(mid);
                const isRight = sinMid >= 0;

                const ax = sinMid * midR,  ay = -cosMid * midR;
                const bx = sinMid * bendR, by = -cosMid * bendR;
                const ex = sinMid * endR,  ey = -cosMid * endR;
                const tx = sinMid * textR, ty = -cosMid * textR;

                labG.append("polyline")
                    .attr("class", "pie-connector")
                    .attr("points",
                        `${ax.toFixed(1)},${ay.toFixed(1)} ` +
                        `${bx.toFixed(1)},${by.toFixed(1)} ` +
                        `${ex.toFixed(1)},${ey.toFixed(1)}`);

                labG.append("text")
                    .attr("class",       "pie-seg-label")
                    .attr("x",           tx + (isRight ? 4 : -4))
                    .attr("y",           ty + labelFontSize * 0.35)
                    .attr("text-anchor", isRight ? "start" : "end")
                    .style("font-size",  `${labelFontSize}px`)
                    .style("font-family", fontFamily)
                    .style("fill",        labelColor)
                    .style("font-weight", boldLabels ? "bold" : "normal")
                    .text(makeLabelText(d.data, labelFormat));
            });
        }

        // ── Hover interactions ────────────────────────────────────────────────
        const self = this;

        segs.on("mouseover", function(_, d) {
            d3.select<SVGPathElement, d3.PieArcDatum<DataPoint>>(this)
                .raise()
                .transition().duration(150)
                .attr("d", arcGen(makeArc(d.startAngle, d.endAngle, hoverR)) ?? "");

            segs.filter(p => p !== d)
                .transition().duration(150)
                .style("opacity", "0.55");

            if (hubValEl) hubValEl.text(fmt(d.data.value));
            if (hubLabEl) hubLabEl.text(`${d.data.pct.toFixed(1)}%`);

            self.legendEl.querySelectorAll<HTMLElement>(".leg-item")
                .forEach((e, i) => e.classList.toggle("active", i === d.index));
        });

        segs.on("mouseout", function(_, d) {
            d3.select<SVGPathElement, d3.PieArcDatum<DataPoint>>(this)
                .transition().duration(150)
                .attr("d", arcGen(makeArc(d.startAngle, d.endAngle, outerR)) ?? "");

            segs.transition().duration(150).style("opacity", null);

            if (hubValEl) hubValEl.text(fmt(self.total));
            if (hubLabEl) hubLabEl.text(centerLabel);

            self.legendEl.querySelectorAll<HTMLElement>(".leg-item")
                .forEach(e => e.classList.remove("active"));
        });

        segs.on("click", function(event: MouseEvent, d) {
            self.selMgr
                .select(d.data.selectionId, event.ctrlKey || event.metaKey)
                .then((ids: ISelectionId[]) => {
                    if (ids.length === 0) {
                        segs.style("opacity", null);
                    } else {
                        segs.style("opacity", (p: d3.PieArcDatum<DataPoint>) =>
                            ids.some((id: ISelectionId) => id.equals(p.data.selectionId))
                                ? "1" : "0.3"
                        );
                    }
                });
            event.stopPropagation();
        });

        segs.on("contextmenu", function(event: MouseEvent, d) {
            event.preventDefault();
            event.stopPropagation();
            self.selMgr.showContextMenu(
                d.data.selectionId,
                { x: event.clientX, y: event.clientY }
            );
        });

        this.svg.on("click", () => {
            this.selMgr.clear().then(() => segs.style("opacity", null));
        });

        // ── Legend ────────────────────────────────────────────────────────────
        this.buildLegend(hasLegend ? sorted : [], legendFontSize);
    }

    // ── Legend DOM builder ─────────────────────────────────────────────────────
    private buildLegend(pts: DataPoint[], fontSize: number): void {
        while (this.legendEl.firstChild) {
            this.legendEl.removeChild(this.legendEl.firstChild);
        }

        for (const dp of pts) {
            const item = document.createElement("div");
            item.className = "leg-item";
            item.style.fontSize = `${fontSize}px`;

            const swatch = document.createElement("span");
            swatch.className = "leg-swatch";
            swatch.style.backgroundColor = dp.color;
            item.appendChild(swatch);

            const name = document.createElement("span");
            name.className = "leg-name";
            name.textContent = dp.name;
            item.appendChild(name);

            const val = document.createElement("span");
            val.className = "leg-val";
            val.textContent = `${fmt(dp.value)} (${dp.pct.toFixed(1)}%)`;
            item.appendChild(val);

            this.legendEl.appendChild(item);
        }
    }

    // ── Formatting model ───────────────────────────────────────────────────────

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
}
