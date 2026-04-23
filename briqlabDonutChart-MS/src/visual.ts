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

// ── Constants ──────────────────────────────────────────────────────────────────
const DEFAULT_PALETTE: string[] = [
    "#0D9488", "#F97316", "#3B82F6", "#8B5CF6", "#10B981",
    "#EF4444", "#F59E0B", "#EC4899", "#06B6D4", "#84CC16"
];
const LEGEND_SIDE_W  = 175;
const HOVER_EXPAND   = 10;
const MAX_LABEL_CHARS = 14;

const TRIAL_KEY     = "briqlab_trial_donut_start";
const PRO_STORE_KEY = "briqlab_donut_prokey";
const TRIAL_MS      = 4 * 24 * 60 * 60 * 1000;

// ── Types ──────────────────────────────────────────────────────────────────────
interface DataPoint {
    name:        string;
    value:       number;
    pct:         number;
    color:       string;
    selectionId: ISelectionId;
}

// ── Helpers ────────────────────────────────────────────────────────────────────
function clamp(v: number, lo: number, hi: number): number {
    return Math.max(lo, Math.min(hi, v));
}

function fmt(v: number, decimals = 1, useAbbrev = true): string {
    const a = Math.abs(v);
    if (useAbbrev) {
        if (a >= 1e6) { const r = (v / 1e6).toFixed(decimals); return (r.endsWith(".0") ? r.slice(0, -2) : r) + "M"; }
        if (a >= 1e3) { const r = (v / 1e3).toFixed(decimals); return (r.endsWith(".0") ? r.slice(0, -2) : r) + "K"; }
    }
    return v.toLocaleString(undefined, { maximumFractionDigits: decimals });
}

function truncate(s: string, max: number): string {
    return s.length > max ? s.slice(0, max - 1) + "\u2026" : s;
}

function makeLabelText(name: string, pct: number, value: number, format: string): string {
    const n = truncate(name, MAX_LABEL_CHARS);
    const pctStr = `${pct.toFixed(1)}%`;
    const valStr = fmt(value);
    switch (format) {
        case "name_value": return `${n}: ${valStr}`;
        case "pct":        return pctStr;
        case "name":       return n;
        case "value":      return valStr;
        default:           return `${n}: ${pctStr}`;
    }
}

function trialDaysRemaining(): number {
    let start = parseInt(localStorage.getItem(TRIAL_KEY) ?? "", 10);
    if (isNaN(start)) {
        start = Date.now();
        localStorage.setItem(TRIAL_KEY, String(start));
    }
    const elapsed = Date.now() - start;
    const remaining = Math.ceil((TRIAL_MS - elapsed) / (24 * 60 * 60 * 1000));
    return remaining;
}

// ── Visual ─────────────────────────────────────────────────────────────────────
export class Visual implements IVisual {
    private readonly host:    IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private readonly selMgr:  ISelectionManager;
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private _handlersAttached = false;
    private readonly fmtSvc:  FormattingSettingsService;

    // DOM nodes
    private readonly root:         HTMLElement;
    private readonly contentEl:    HTMLElement;
    private readonly chartArea:    HTMLElement;
    private readonly legendEl:     HTMLElement;
    private readonly trialBadge:   HTMLElement;
    private readonly proBadge:     HTMLElement;
    private readonly keyError:     HTMLElement;
    private readonly trialOverlay: HTMLElement;
    private readonly svg: d3.Selection<SVGSVGElement, unknown, null, undefined>;

    // State
    private settings!: VisualFormattingSettingsModel;
    private vp:        powerbi.IViewport = { width: 300, height: 300 };
    private pts:       DataPoint[]       = [];
    private total:     number            = 0;

    // Pro / trial
    private curKey:    string = "";
    private isPro:     boolean = false;
    private keyCache:  Map<string, boolean> = new Map();

    // ── Constructor ────────────────────────────────────────────────────────────
    constructor(options: VisualConstructorOptions) {
        this.host   = options.host;
        this.renderingManager = options.host.eventService;
        this.selMgr = options.host.createSelectionManager();
        this.tooltipSvc = this.host.tooltipService;
        this.fmtSvc = new FormattingSettingsService();

        this.root = options.element;
        this.root.classList.add("briqlab-donut");

        // .briqlab-visual-content
        this.contentEl = document.createElement("div");
        this.contentEl.className = "briqlab-visual-content";
        this.root.appendChild(this.contentEl);

        // .donut-chart-area
        this.chartArea = document.createElement("div");
        this.chartArea.className = "donut-chart-area";
        this.contentEl.appendChild(this.chartArea);

        // .donut-legend
        this.legendEl = document.createElement("div");
        this.legendEl.className = "donut-legend";
        this.contentEl.appendChild(this.legendEl);

        // SVG inside chart area
        this.svg = d3.select(this.chartArea)
            .append<SVGSVGElement>("svg")
            .attr("class", "donut-svg");

        // Trial badge
        this.trialBadge = document.createElement("div");
        this.trialBadge.className = "briqlab-trial-badge hidden";
        this.root.appendChild(this.trialBadge);

        // Pro badge
        this.proBadge = document.createElement("div");
        this.proBadge.className = "briqlab-pro-badge hidden";
        this.proBadge.textContent = "\u2713 Pro Active";
        this.root.appendChild(this.proBadge);

        // Key error
        this.keyError = document.createElement("div");
        this.keyError.className = "briqlab-key-error hidden";
        this.keyError.textContent = "\u2717 Invalid key";
        this.root.appendChild(this.keyError);

        // Trial overlay
        this.trialOverlay = this.buildTrialOverlay();
        this.root.appendChild(this.trialOverlay);

        // Restore stored pro key silently.
        // NOTE: do NOT set this.curKey here — if we did, update() would see
        // (rawKey="" !== curKey=storedKey) and call localStorage.removeItem,
        // erasing the key on every page load.  Instead validate async-only.
        try {
            const storedKey = localStorage.getItem(PRO_STORE_KEY) ?? "";
            if (storedKey) {
                this.isPro = true; // optimistic until validated
                this.validateKey(storedKey).then(ok => {
                    this.isPro = ok;
                    if (!ok) localStorage.removeItem(PRO_STORE_KEY);
                    this.updateTrialUI();
                });
            }
        } catch { /* ignore */ }
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

    // ── Build trial overlay (no innerHTML) ────────────────────────────────────
    private buildTrialOverlay(): HTMLElement {
        const overlay = document.createElement("div");
        overlay.className = "briqlab-trial-overlay hidden";

        const card = document.createElement("div");
        card.className = "briqlab-trial-card";

        // Icon (lock SVG via inline textContent — pure DOM)
        const iconDiv = document.createElement("div");
        iconDiv.className = "trial-icon";
        const iconSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        iconSvg.setAttribute("viewBox", "0 0 24 24");
        iconSvg.setAttribute("fill", "none");
        iconSvg.setAttribute("width", "40");
        iconSvg.setAttribute("height", "40");
        const iconPath = document.createElementNS("http://www.w3.org/2000/svg", "path");
        iconPath.setAttribute("d", "M12 2a5 5 0 0 1 5 5v3h1a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2v-8a2 2 0 0 1 2-2h1V7a5 5 0 0 1 5-5zm0 2a3 3 0 0 0-3 3v3h6V7a3 3 0 0 0-3-3z");
        iconPath.setAttribute("fill", "#F97316");
        iconSvg.appendChild(iconPath);
        iconDiv.appendChild(iconSvg);

        const title = document.createElement("p");
        title.className = "trial-title";
        title.textContent = "Free trial ended";

        const body = document.createElement("p");
        body.className = "trial-body";
        body.textContent = "Activate Briqlab Pro to continue using this visual and unlock all features.";

        const btn = document.createElement("button");
        btn.className = "trial-btn";
        btn.textContent = getButtonText();
        btn.addEventListener("click", (e: MouseEvent) => {
            e.stopPropagation();
            this.host.launchUrl(getPurchaseUrl());
        });

        const sub = document.createElement("p");
        sub.className = "trial-subtext";
        sub.textContent = "Purchase on Microsoft AppSource to unlock all features instantly.";

        card.appendChild(iconDiv);
        card.appendChild(title);
        card.appendChild(body);
        card.appendChild(btn);
        card.appendChild(sub);
        overlay.appendChild(card);

        return overlay;
    }

    // ── Pro validation ─────────────────────────────────────────────────────────
    private async validateKey(key: string): Promise<boolean> {
        return checkMicrosoftLicence(this.host);
    }

    // ── Trial / pro UI state ───────────────────────────────────────────────────
    private updateTrialUI(): void {
        checkMicrosoftLicence(this.host).then(p => this._msUpdateLicenceUI(p)).catch(() => this._msUpdateLicenceUI(false));
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
                        dataItems: [{ displayName: "Briqlab Donut Chart", value: "" }],
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
    
            // Extract categorical data
            this.pts   = [];
            this.total = 0;
            const dv   = options.dataViews?.[0];
    
            if (dv?.categorical?.categories?.length && dv.categorical.values?.length) {
                const cats    = dv.categorical.categories[0];
                const vals    = dv.categorical.values[0];
                const palette = this.getPalette();
                const sum     = (vals.values as number[]).reduce((acc, v) => acc + (v ?? 0), 0);
                this.total    = sum;
    
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
    
            // Handle pro key change
            const rawKey = ""; // MS cert: pro key field removed
            if (rawKey !== this.curKey) {
                this.curKey = rawKey;
                this.isPro  = false;
                this.keyError.classList.add("hidden");
                if (rawKey) {
                    localStorage.setItem(PRO_STORE_KEY, rawKey);
                    this.validateKey(rawKey).then(ok => {
                        this.isPro = ok;
                        if (!ok) {
                            this.keyError.classList.remove("hidden");
                        }
                        this.updateTrialUI();
                        this.render();
                    });
                } else {
                    localStorage.removeItem(PRO_STORE_KEY);
                }
            }
    
            this.updateTrialUI();
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

        // ── Settings ──────────────────────────────────────────────────────────
        const innerRadiusPct    = clamp(s.donutSettings.innerRadius.value, 10, 80);
        const outerRadiusPct    = clamp(s.donutSettings.outerRadius.value, 40, 95) / 100;
        const sortOrder         = String(s.donutSettings.sortOrder?.value?.value ?? "original");
        const startAngleDeg     = s.donutSettings.startAngle.value ?? 0;
        const showCenter        = s.donutSettings.showCenter.value;
        const centerLabel       = s.donutSettings.centerLabel.value || "Total";
        const centerDecimals    = clamp(s.donutSettings.centerValueDecimals.value ?? 0, 0, 4);
        const centerFmtType     = String(s.donutSettings.centerValueFormat?.value?.value ?? "auto");
        const borderWidth       = clamp(s.donutSettings.borderWidth.value, 0, 10);
        const minLabelPct       = clamp(s.donutSettings.minLabelPct.value ?? 4, 0, 50);

        const showLabels        = s.labelSettings.showLabels.value;
        const labelFontSize     = clamp(s.labelSettings.labelFontSize.value, 8, 18);
        const labelFormat       = String(s.labelSettings.labelFormat?.value?.value ?? "name_pct");
        const labelColor        = s.labelSettings.labelColor?.value?.value ?? "#374151";
        const boldLabels        = s.labelSettings.boldLabels?.value ?? false;
        const fontFamily        = String(s.labelSettings.fontFamily?.value?.value ?? "Segoe UI");

        const showLegend        = s.legendSettings.showLegend.value;
        const rawPos            = String(s.legendSettings.legendPosition.value?.value ?? "Right").toLowerCase();
        const legendPos = (["left", "top", "bottom", "right"] as const).includes(
            rawPos as "left" | "top" | "bottom" | "right"
        ) ? (rawPos as "left" | "top" | "bottom" | "right") : "right";
        const legendFontSize    = clamp(s.legendSettings.legendFontSize?.value ?? 11, 8, 18);

        // ── Sort ──────────────────────────────────────────────────────────────
        let sorted = [...this.pts];
        if (sortOrder === "desc") sorted.sort((a, b) => b.value - a.value);
        else if (sortOrder === "asc") sorted.sort((a, b) => a.value - b.value);

        // ── Layout ────────────────────────────────────────────────────────────
        const hasLegend  = showLegend && sorted.length > 0;
        const itemH      = legendFontSize * 2.2;
        const legendHMax = Math.min(sorted.length * itemH, height * 0.30);

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

        const flexDir = legendPos === "left"   ? "row-reverse"    :
                        legendPos === "top"    ? "column-reverse" :
                        legendPos === "bottom" ? "column"         : "row";

        this.contentEl.style.cssText =
            `width:${width}px;height:${height}px;display:flex;flex-direction:${flexDir};position:relative;`;

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

        // ── SVG setup ─────────────────────────────────────────────────────────
        this.svg.attr("width", chartW).attr("height", chartH);
        this.svg.selectAll("*").remove();

        const cx = chartW / 2;
        const cy = chartH / 2;
        const maxR   = Math.min(chartW, chartH) / 2 * (showLabels ? 0.65 : 0.82);
        const outerR = maxR * outerRadiusPct;
        const innerR = outerR * (innerRadiusPct / 100);
        const hoverR = outerR + HOVER_EXPAND;

        const startAngleRad = (startAngleDeg * Math.PI) / 180;

        const arcGen  = d3.arc<d3.DefaultArcObject>();
        const makeArc = (sa: number, ea: number, r: number): d3.DefaultArcObject => ({
            startAngle: sa, endAngle: ea,
            innerRadius: innerR, outerRadius: r, padAngle: 0
        });

        // ── Empty state ───────────────────────────────────────────────────────
        if (sorted.length === 0) {
            this.svg.append("circle").attr("cx", cx).attr("cy", cy).attr("r", outerR).attr("fill", "#F3F4F6");
            this.svg.append("circle").attr("cx", cx).attr("cy", cy).attr("r", innerR).attr("fill", "#ffffff");
            this.svg.append("text")
                .attr("x", cx).attr("y", cy)
                .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                .attr("class", "donut-empty-msg").text("Add Category & Value fields");
            this.buildLegend([], legendFontSize);
            return;
        }

        // ── Pie layout ────────────────────────────────────────────────────────
        const pie     = d3.pie<DataPoint>().value(d => d.value).sort(null).padAngle(0).startAngle(startAngleRad);
        const pieData = pie(sorted);

        // ── Segments ──────────────────────────────────────────────────────────
        const arcsG = this.svg.append("g")
            .attr("class", "donut-arcs")
            .attr("transform", `translate(${cx},${cy})`);

        const segs = arcsG
            .selectAll<SVGPathElement, d3.PieArcDatum<DataPoint>>("path")
            .data(pieData)
            .join("path")
            .attr("class",        "donut-seg")
            .attr("fill",         d => d.data.color)
            .attr("stroke",       "#ffffff")
            .attr("stroke-width", borderWidth)
            .style("cursor",      "pointer")
            .attr("d", d => arcGen(makeArc(d.startAngle, d.startAngle, outerR)) ?? "");

        segs.transition().duration(650).ease(d3.easeCubicOut)
            .attrTween("d", function(d) {
                const interp = d3.interpolate(
                    makeArc(d.startAngle, d.startAngle, outerR),
                    makeArc(d.startAngle, d.endAngle,   outerR)
                );
                return (t: number) => arcGen(interp(t)) ?? "";
            });

        // ── Center text ───────────────────────────────────────────────────────
        const cValFS  = clamp(innerR * 0.38, 12, 44);
        const cLabFS  = clamp(innerR * 0.20, 9,  16);
        const centerG = this.svg.append("g").attr("transform", `translate(${cx},${cy})`);

        const centerValText = fmt(this.total, centerDecimals, centerFmtType === "auto");

        const cValEl = showCenter
            ? centerG.append("text")
                .attr("class",       "donut-center-val")
                .attr("text-anchor", "middle")
                .attr("y",           cValFS * 0.35)
                .style("font-size",  `${cValFS}px`)
                .style("font-family", fontFamily)
                .text(centerValText)
            : null;

        const cLabEl = showCenter
            ? centerG.append("text")
                .attr("class",       "donut-center-lab")
                .attr("text-anchor", "middle")
                .attr("y",           cValFS * 0.35 + cLabFS + 3)
                .style("font-size",  `${cLabFS}px`)
                .style("font-family", fontFamily)
                .text(centerLabel)
            : null;

        // ── Labels + connector lines ──────────────────────────────────────────
        if (showLabels) {
            const labG = this.svg.append("g")
                .attr("class", "donut-labels")
                .attr("transform", `translate(${cx},${cy})`);

            const midR   = (innerR + outerR) / 2;
            const lineR1 = outerR  * 1.04;
            const lineR2 = outerR  * 1.13;
            const textR  = outerR  * 1.17;

            pieData.forEach(d => {
                if (d.data.pct < minLabelPct) return;

                const mid     = (d.startAngle + d.endAngle) / 2;
                const sinMid  = Math.sin(mid);
                const cosMid  = Math.cos(mid);
                const isRight = sinMid >= 0;

                const ax = sinMid * midR,   ay = -cosMid * midR;
                const bx = sinMid * lineR1, by = -cosMid * lineR1;
                const ex = sinMid * lineR2, ey = -cosMid * lineR2;
                const tx = sinMid * textR,  ty = -cosMid * textR;

                labG.append("polyline")
                    .attr("class", "donut-connector")
                    .attr("points",
                        `${ax.toFixed(1)},${ay.toFixed(1)} ` +
                        `${bx.toFixed(1)},${by.toFixed(1)} ` +
                        `${ex.toFixed(1)},${ey.toFixed(1)}`);

                labG.append("text")
                    .attr("class",        "donut-pct-label")
                    .attr("x",            tx + (isRight ? 4 : -4))
                    .attr("y",            ty + labelFontSize * 0.35)
                    .attr("text-anchor",  isRight ? "start" : "end")
                    .style("font-size",   `${labelFontSize}px`)
                    .style("font-family", fontFamily)
                    .style("fill",        labelColor)
                    .style("font-weight", boldLabels ? "bold" : "normal")
                    .text(makeLabelText(d.data.name, d.data.pct, d.data.value, labelFormat));
            });
        }

        // ── Hover ─────────────────────────────────────────────────────────────
        const self = this;

        segs.on("mouseover", function(_, d) {
            d3.select<SVGPathElement, d3.PieArcDatum<DataPoint>>(this)
                .raise().transition().duration(150)
                .attr("d", arcGen(makeArc(d.startAngle, d.endAngle, hoverR)) ?? "");

            segs.filter(p => p !== d).transition().duration(150).style("opacity", "0.6");

            if (cValEl) cValEl.text(fmt(d.data.value, centerDecimals, centerFmtType === "auto"));
            if (cLabEl) cLabEl.text(`${d.data.pct.toFixed(1)}%`);

            self.legendEl.querySelectorAll<HTMLElement>(".leg-item")
                .forEach((e, i) => e.classList.toggle("active", i === d.index));
        });

        segs.on("mouseout", function(_, d) {
            d3.select<SVGPathElement, d3.PieArcDatum<DataPoint>>(this)
                .transition().duration(150)
                .attr("d", arcGen(makeArc(d.startAngle, d.endAngle, outerR)) ?? "");

            segs.transition().duration(150).style("opacity", null);

            if (cValEl) cValEl.text(centerValText);
            if (cLabEl) cLabEl.text(centerLabel);

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
                                ? "1" : "0.35"
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

    // ── Format model ───────────────────────────────────────────────────────────

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
