"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import * as d3 from "d3";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions       = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual                   = powerbi.extensibility.visual.IVisual;
import IVisualHost               = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager         = powerbi.extensibility.ISelectionManager;
import ISelectionId              = powerbi.visuals.ISelectionId;

import { VisualFormattingSettingsModel } from "./settings";
import { checkMicrosoftLicence, resetLicenceCache } from "./licenceManager";
import { getTrialDaysRemaining, isTrialExpired, getPurchaseUrl, getButtonText } from "./trialManager";

// ── Constants ──────────────────────────────────────────────────────────────────
const TRIAL_KEY   = "briqlab_trial_barchart_start";
const PRO_LS_KEY  = "briqlab_barchart_prokey";
const TRIAL_MS    = 4 * 24 * 60 * 60 * 1000;
const KEY_CACHE: Map<string, boolean> = new Map();

// ── Types ──────────────────────────────────────────────────────────────────────
interface DataPoint {
    category:    string;
    value:       number;
    comparison?: number;
    selectionId: ISelectionId;
    index:       number;
}

interface BarEntry {
    pt:     DataPoint;
    series: string;
    color:  string;
    value:  number;
}

// ── Helpers ────────────────────────────────────────────────────────────────────
function fmtVal(v: number): string {
    const a = Math.abs(v);
    if (a >= 1_000_000) {
        const r = (v / 1_000_000).toFixed(1);
        return (r.endsWith(".0") ? r.slice(0, -2) : r) + "M";
    }
    if (a >= 1_000) {
        const r = (v / 1_000).toFixed(1);
        return (r.endsWith(".0") ? r.slice(0, -2) : r) + "K";
    }
    return v.toLocaleString("en-US");
}

function topRoundedPath(x: number, y: number, w: number, h: number, r: number): string {
    if (h <= 0 || w <= 0) return "";
    const e = Math.min(r, Math.abs(h) / 2, Math.abs(w) / 2);
    if (e <= 0) return `M${x},${y + h} L${x},${y} L${x + w},${y} L${x + w},${y + h} Z`;
    return `M${x},${y + h} L${x},${y + e} Q${x},${y} ${x + e},${y} L${x + w - e},${y} Q${x + w},${y} ${x + w},${y + e} L${x + w},${y + h} Z`;
}

function rightRoundedPath(x: number, y: number, w: number, h: number, r: number): string {
    if (h <= 0 || w <= 0) return "";
    const e = Math.min(r, Math.abs(w) / 2, Math.abs(h) / 2);
    if (e <= 0) return `M${x},${y} L${x + w},${y} L${x + w},${y + h} L${x},${y + h} Z`;
    return `M${x},${y} L${x + w - e},${y} Q${x + w},${y} ${x + w},${y + e} L${x + w},${y + h - e} Q${x + w},${y + h} ${x + w - e},${y + h} L${x},${y + h} Z`;
}

// ── Visual ─────────────────────────────────────────────────────────────────────
export class Visual implements IVisual {
    private host:             IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private selectionManager: ISelectionManager;
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private _handlersAttached = false;
    private fmtService:       FormattingSettingsService;
    private settings!:        VisualFormattingSettingsModel;

    // DOM elements
    private root:        HTMLElement;
    private contentEl:   HTMLElement;
    private overlayEl:   HTMLElement;
    private trialBadge:  HTMLElement;
    private proBadge:    HTMLElement;
    private keyErrorEl:  HTMLElement;
    private tooltip:     HTMLElement;
    private ttName:      HTMLElement;
    private ttValue:     HTMLElement;

    private lastData:    DataPoint[] = [];
    private lastOptions: VisualUpdateOptions | null = null;
    private proValid:    boolean = false;

    constructor(options: VisualConstructorOptions) {
        // Fixed: use options.host (not options.element)
        this.host             = options.host;
        this.renderingManager = options.host.eventService;
        this.fmtService       = new FormattingSettingsService();
        this.selectionManager = options.host.createSelectionManager();
        this.tooltipSvc = this.host.tooltipService;

        // ── Root DOM ──────────────────────────────────────────────────────────
        // root: .briqlab-bar [position:relative]
        this.root = document.createElement("div");
        this.root.className = "briqlab-bar";
        options.element.appendChild(this.root);

        // .briqlab-visual-content  ← holds chart SVG, gets blur class
        this.contentEl = document.createElement("div");
        this.contentEl.className = "briqlab-visual-content";
        this.root.appendChild(this.contentEl);

        // .briqlab-trial-overlay  ← absolute, z-index 100, NOT blurred
        this.overlayEl = document.createElement("div");
        this.overlayEl.className = "briqlab-trial-overlay hidden";
        this.root.appendChild(this.overlayEl);

        // trial card inside overlay
        const card = document.createElement("div");
        card.className = "briqlab-trial-card";
        this.overlayEl.appendChild(card);

        const titleEl = document.createElement("p");
        titleEl.className = "trial-title";
        titleEl.textContent = "Free trial ended";
        card.appendChild(titleEl);

        const bodyEl = document.createElement("p");
        bodyEl.className = "trial-body";
        bodyEl.textContent = "Your 4-day free trial has ended. Purchase Briqlab Pro on Microsoft AppSource to continue.";
        card.appendChild(bodyEl);

        const btnEl = document.createElement("button");
        btnEl.className = "trial-btn";
        btnEl.textContent = getButtonText();
        btnEl.addEventListener("click", () => {
            this.host.launchUrl(getPurchaseUrl());
        });
        card.appendChild(btnEl);

        const subEl = document.createElement("p");
        subEl.className = "trial-subtext";
        subEl.textContent = "Purchase on Microsoft AppSource to unlock all features instantly.";
        card.appendChild(subEl);

        // .briqlab-trial-badge (bottom-left, outside contentEl)
        this.trialBadge = document.createElement("div");
        this.trialBadge.className = "briqlab-trial-badge hidden";
        this.root.appendChild(this.trialBadge);

        // .briqlab-pro-badge (bottom-right, outside contentEl)
        this.proBadge = document.createElement("div");
        this.proBadge.className = "briqlab-pro-badge hidden";
        this.proBadge.textContent = "\u2713 Pro Active";
        this.root.appendChild(this.proBadge);

        // .briqlab-key-error (near bottom-right, outside contentEl)
        this.keyErrorEl = document.createElement("div");
        this.keyErrorEl.className = "briqlab-key-error hidden";
        this.keyErrorEl.textContent = "\u2717 Invalid key";
        this.root.appendChild(this.keyErrorEl);

        // ── Tooltip ───────────────────────────────────────────────────────────
        this.tooltip = document.createElement("div");
        this.tooltip.className = "bar-tooltip";
        this.tooltip.style.display = "none";
        this.ttName  = document.createElement("div");
        this.ttName.className  = "tt-name";
        this.ttValue = document.createElement("div");
        this.ttValue.className = "tt-value";
        this.tooltip.appendChild(this.ttName);
        this.tooltip.appendChild(this.ttValue);
        this.contentEl.appendChild(this.tooltip);

        // ── Trial init ────────────────────────────────────────────────────────
        this.initTrial();

        // ── Restore saved pro key ──────────────────────────────────────────────
        try {
            const savedKey = localStorage.getItem(PRO_LS_KEY);
            if (savedKey) {
                this.validateKey(savedKey, true);
            }
        } catch (_e) {
            // localStorage not available
        }
    }

    // ── Trial management ──────────────────────────────────────────────────────
    private initTrial(): void {
        try {
            let start = localStorage.getItem(TRIAL_KEY);
            if (!start) {
                start = String(Date.now());
                localStorage.setItem(TRIAL_KEY, start);
            }
            const elapsed = Date.now() - parseInt(start, 10);
            const daysLeft = Math.ceil((TRIAL_MS - elapsed) / (24 * 60 * 60 * 1000));

            if (elapsed < TRIAL_MS) {
                // Still in trial
                this.trialBadge.textContent = "Trial: " + daysLeft + " day" + (daysLeft === 1 ? "" : "s") + " remaining";
            }
        } catch (_e) {
            // localStorage not available — treat as day 1
        }
    }

    private getTrialStatus(): "active" | "expired" {
        try {
            const start = localStorage.getItem(TRIAL_KEY);
            if (!start) return "active";
            const elapsed = Date.now() - parseInt(start, 10);
            return elapsed < TRIAL_MS ? "active" : "expired";
        } catch (_e) {
            return "active";
        }
    }

    private updateTrialUI(): void {
        checkMicrosoftLicence(this.host).then(p => this._msUpdateLicenceUI(p)).catch(() => this._msUpdateLicenceUI(false));
    }

    private validateKey(key: string, silent: boolean): void {
        checkMicrosoftLicence(this.host).then(p => { this.proValid = p; this._msUpdateLicenceUI(p); this.render(); }).catch(() => this._msUpdateLicenceUI(false));
    }

    // ── Update ────────────────────────────────────────────────────────────────
    public update(options: VisualUpdateOptions): void {
        this.renderingManager.renderingStarted(options);
        try {
            // AppSource: attach context-menu + tooltip once
            if (!this._handlersAttached) {
                this._handlersAttached = true;
                this.root.addEventListener("contextmenu", (e: MouseEvent) => {
                    e.preventDefault();
                    this.selectionManager.showContextMenu(
                        null as unknown as powerbi.visuals.ISelectionId,
                        { x: e.clientX, y: e.clientY }
                    );
                });
                this.root.addEventListener("mousemove", (e: MouseEvent) => {
                    this.tooltipSvc.show({
                        dataItems: [{ displayName: "Briqlab Bar Chart", value: "" }],
                        identities: [],
                        coordinates: [e.clientX, e.clientY],
                        isTouchEvent: false
                    });
                });
                this.root.addEventListener("mouseleave", () => {
                    this.tooltipSvc.hide({ isTouchEvent: false, immediately: false });
                });
            }
            this.settings    = this.fmtService.populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews[0]);
            this.lastOptions = options;
            this.lastData    = this.parseData(options);
    
            const proKey = ""; // MS cert: pro key field removed
    
            if (proKey === "") {
                this.proValid = false;
                this.keyErrorEl.classList.add("hidden");
                this.updateTrialUI();
                this.render();
            } else {
                this.validateKey(proKey, false);
                // render immediately with current state while validation is in-flight
                this.updateTrialUI();
                this.render();
            }
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    // ── Parse data ────────────────────────────────────────────────────────────
    private parseData(options: VisualUpdateOptions): DataPoint[] {
        const dv = options.dataViews?.[0];
        if (!dv?.categorical) return [];
        const cat = dv.categorical;
        const cats = cat.categories?.[0];
        if (!cats?.values?.length) return [];

        const vals  = cat.values ?? [];
        let measureCol: powerbi.DataViewValueColumn | undefined;
        let compareCol: powerbi.DataViewValueColumn | undefined;

        for (const col of vals) {
            if (col.source.roles?.["measure"])           measureCol = col;
            if (col.source.roles?.["comparisonMeasure"]) compareCol = col;
        }

        return cats.values.map((catVal, i) => {
            const host = this.host as unknown as Record<string, unknown>;
            const builder = (host["createSelectionIdBuilder"] as () => powerbi.visuals.ISelectionIdBuilder)();
            const selId = builder.withCategory(cats, i).createSelectionId();
            return {
                category:    String(catVal ?? ""),
                value:       measureCol ? Number(measureCol.values[i] ?? 0) : 0,
                comparison:  compareCol ? Number(compareCol.values[i] ?? 0) : undefined,
                selectionId: selId,
                index:       i
            };
        });
    }

    // ── Number formatter using format settings ────────────────────────────────
    private fmtNum(v: number): string {
        const s = this.settings;
        const decimals = Number(s.numberFormat?.decimalPlaces?.value ?? 1);
        const useKM    = s.numberFormat?.useKM?.value ?? true;
        const prefix   = s.numberFormat?.prefix?.value ?? "";
        const suffix   = s.numberFormat?.suffix?.value ?? "";
        const a = Math.abs(v);
        let body: string;
        if (useKM) {
            if (a >= 1_000_000_000) body = (v / 1_000_000_000).toFixed(decimals) + "B";
            else if (a >= 1_000_000) body = (v / 1_000_000).toFixed(decimals) + "M";
            else if (a >= 1_000)     body = (v / 1_000).toFixed(decimals) + "K";
            else                     body = v.toFixed(decimals);
        } else {
            body = v.toLocaleString(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
        }
        return `${prefix}${body}${suffix}`;
    }

    // ── Render ────────────────────────────────────────────────────────────────
    private render(): void {
        const data = this.lastData;
        const opts = this.lastOptions;
        if (!opts) return;

        const s         = this.settings;
        const orient    = String(s.chartSettings.orientation.value?.value ?? "vertical");
        const barMode   = String(s.chartSettings.barMode?.value?.value ?? "grouped");
        const barColor  = s.chartSettings.barColor.value?.value              ?? "#0D9488";
        const cmpColor  = s.chartSettings.comparisonColor.value?.value       ?? "#F97316";
        const s3Color   = s.chartSettings.series3Color?.value?.value         ?? "#3B82F6";
        const cornerR   = Math.max(0, Math.min(8, Number(s.chartSettings.cornerRadius.value ?? 4)));
        const barPad    = Math.max(0, Math.min(50, Number(s.chartSettings.barPadding?.value ?? 25))) / 100;
        const showBg    = s.chartSettings.showBackground?.value ?? false;
        const bgColor   = s.chartSettings.backgroundColor?.value?.value ?? "#F8FAFA";
        const showLbl   = s.labelSettings.showLabels.value ?? false;
        const lblFz     = Number(s.labelSettings.labelFontSize.value ?? 10);
        const lblFormat = String(s.labelSettings.labelFormat?.value?.value ?? "auto");
        const lblColor  = s.labelSettings.labelColor?.value?.value ?? "#374151";
        const boldLbls  = s.labelSettings.boldLabels?.value ?? false;
        const fontFam   = String(s.fontSettings?.fontFamily?.value?.value ?? "Segoe UI");
        const boldAxis  = s.fontSettings?.boldAxis?.value ?? false;
        const italicAxis = s.fontSettings?.italicAxis?.value ?? false;
        const xFz       = Number(s.axisSettings.xFontSize.value  ?? 10);
        const yFz       = Number(s.axisSettings.yFontSize.value  ?? 10);
        const xTitle    = s.axisSettings.xTitle?.value ?? "";
        const yTitle    = s.axisSettings.yTitle?.value ?? "";
        const showGrid  = s.axisSettings.showGridlines.value ?? true;
        const gridColor = s.axisSettings.gridlineColor?.value?.value ?? "#E2E8F0";
        const showZero  = s.axisSettings.showZeroLine?.value ?? false;
        const hasComp   = data.some(d => d.comparison !== undefined);

        const colors = [barColor, cmpColor, s3Color];

        const W = opts.viewport.width;
        const H = opts.viewport.height;

        // Background
        this.contentEl.style.backgroundColor = showBg ? bgColor : "";

        // ── Clear ─────────────────────────────────────────────────────────────
        d3.select(this.contentEl).selectAll<SVGSVGElement, unknown>("svg.bar-svg").remove();

        if (!data.length) {
            const emptySvg = d3.select(this.contentEl)
                .append("svg").attr("class", "bar-svg")
                .attr("width", W).attr("height", H);
            emptySvg.append("text").attr("class", "bar-empty-msg")
                .attr("x", W / 2).attr("y", H / 2)
                .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                .style("font-family", fontFam)
                .text("Add Category and Values to get started");
            return;
        }

        const hasXTitle = xTitle.trim().length > 0;
        const hasYTitle = yTitle.trim().length > 0;

        const margin = orient === "vertical"
            ? { top: 20, right: 16, bottom: (data.length > 6 ? 70 : 40) + (hasXTitle ? 20 : 0), left: 52 + (hasYTitle ? 16 : 0) }
            : { top: 16, right: 52, bottom: 36 + (hasXTitle ? 16 : 0), left: 90 + (hasYTitle ? 0 : 0) };

        const innerW = Math.max(10, W - margin.left - margin.right);
        const innerH = Math.max(10, H - margin.top  - margin.bottom);

        const svgEl = d3.select(this.contentEl)
            .append("svg").attr("class", "bar-svg")
            .attr("width", W).attr("height", H)
            .style("font-family", fontFam);

        const g = svgEl.append("g")
            .attr("transform", `translate(${margin.left},${margin.top})`);

        // Axis titles
        if (yTitle && orient === "vertical") {
            svgEl.append("text")
                .attr("class", "axis-title")
                .attr("transform", `translate(${margin.left - 36},${margin.top + innerH / 2}) rotate(-90)`)
                .attr("text-anchor", "middle")
                .style("font-size", `${yFz}px`)
                .style("font-family", fontFam)
                .text(yTitle);
        }
        if (xTitle) {
            svgEl.append("text")
                .attr("class", "axis-title")
                .attr("x", margin.left + innerW / 2)
                .attr("y", H - 4)
                .attr("text-anchor", "middle")
                .style("font-size", `${xFz}px`)
                .style("font-family", fontFam)
                .text(xTitle);
        }

        const renderArgs = { g, data, innerW, innerH, colors, cornerR, barPad, showLbl, lblFz, lblFormat, lblColor, boldLbls, fontFam, boldAxis, italicAxis, xFz, yFz, showGrid, gridColor, showZero, hasComp };

        if (barMode === "stacked" || barMode === "stacked100") {
            if (orient === "vertical") {
                this.renderStackedVertical(renderArgs, barMode === "stacked100");
            } else {
                this.renderStackedHorizontal(renderArgs, barMode === "stacked100");
            }
        } else {
            if (orient === "vertical") {
                this.renderVertical(renderArgs);
            } else {
                this.renderHorizontal(renderArgs);
            }
        }
    }

    // ── Render args interface ────────────────────────────────────────────────
    private applyAxisStyle(sel: d3.Selection<SVGGElement, unknown, null, undefined>, fz: number, fontFam: string, bold: boolean, italic: boolean): void {
        sel.selectAll("text")
            .style("font-size", `${fz}px`)
            .style("font-family", fontFam)
            .style("font-weight", bold ? "bold" : "normal")
            .style("font-style", italic ? "italic" : "normal");
    }

    private makeLabelText(v: number, total: number, fmt: string): string {
        if (fmt === "pct") return `${((v / total) * 100).toFixed(1)}%`;
        if (fmt === "value") return v.toLocaleString();
        return this.fmtNum(v);
    }

    // ── Stacked vertical ────────────────────────────────────────────────────
    private renderStackedVertical(args: Record<string, unknown>, normalize: boolean): void {
        const { g, data, innerW, innerH, colors, cornerR, barPad, showLbl, lblFz, lblColor, boldLbls, fontFam, boldAxis, italicAxis, xFz, yFz, showGrid, gridColor, showZero, hasComp } =
            args as { g: d3.Selection<SVGGElement, unknown, null, undefined>; data: DataPoint[]; innerW: number; innerH: number; colors: string[]; cornerR: number; barPad: number; showLbl: boolean; lblFz: number; lblFormat: string; lblColor: string; boldLbls: boolean; fontFam: string; boldAxis: boolean; italicAxis: boolean; xFz: number; yFz: number; showGrid: boolean; gridColor: string; showZero: boolean; hasComp: boolean };

        const seriesKeys = hasComp ? ["value", "comparison"] : ["value"];

        const catScale = d3.scaleBand().domain(data.map(d => d.category)).range([0, innerW]).padding(barPad);

        const totals = data.map(d => (d.value ?? 0) + (hasComp ? (d.comparison ?? 0) : 0));
        const maxTotal = normalize ? 100 : Math.max(...totals, 0);

        const yScale = d3.scaleLinear().domain([0, maxTotal * 1.05]).range([innerH, 0]).nice();

        if (showGrid) {
            g.append("g").selectAll<SVGLineElement, number>("line").data(yScale.ticks(5)).join("line")
                .attr("x1", 0).attr("x2", innerW)
                .attr("y1", d => yScale(d)).attr("y2", d => yScale(d))
                .attr("stroke", gridColor).attr("stroke-width", 0.5).attr("stroke-dasharray", "3,3");
        }
        if (showZero) {
            g.append("line").attr("x1", 0).attr("x2", innerW).attr("y1", yScale(0)).attr("y2", yScale(0))
                .attr("stroke", "#374151").attr("stroke-width", 1);
        }

        const yAxis = d3.axisLeft(yScale).ticks(5).tickFormat(d => normalize ? `${d}%` : this.fmtNum(Number(d)));
        const yG = g.append("g").attr("class", "axis-y").call(yAxis);
        this.applyAxisStyle(yG, yFz, fontFam, boldAxis, italicAxis);

        const xAxis = d3.axisBottom(catScale).tickSizeOuter(0);
        const xG = g.append("g").attr("class", "axis-x").attr("transform", `translate(0,${innerH})`).call(xAxis);
        this.applyAxisStyle(xG, xFz, fontFam, boldAxis, italicAxis);
        if (data.length > 6) {
            xG.selectAll("text").attr("transform", "rotate(-45)").attr("text-anchor", "end").attr("dx", "-0.4em").attr("dy", "0.6em");
        }

        // Stacked bars
        let yOffsets = new Array(data.length).fill(0);
        seriesKeys.forEach((key, si) => {
            data.forEach((d, i) => {
                const raw = key === "value" ? d.value : (d.comparison ?? 0);
                const totalVal = totals[i] || 1;
                const v = normalize ? (raw / totalVal) * 100 : raw;
                const x = catScale(d.category) ?? 0;
                const w = catScale.bandwidth();
                const y0 = yOffsets[i];
                const y1 = y0 + v;
                const sy = yScale(y1);
                const sh = yScale(y0) - sy;
                const r = si === seriesKeys.length - 1 ? Math.min(cornerR, sh / 2, w / 2) : 0;
                const path = r > 0
                    ? `M${x},${sy + sh} L${x},${sy + r} Q${x},${sy} ${x + r},${sy} L${x + w - r},${sy} Q${x + w},${sy} ${x + w},${sy + r} L${x + w},${sy + sh} Z`
                    : `M${x},${sy} L${x + w},${sy} L${x + w},${sy + sh} L${x},${sy + sh} Z`;
                g.append("path").attr("d", path).attr("fill", (colors as string[])[si] ?? "#0D9488").attr("opacity", 0.9);
                if (showLbl && sh > lblFz + 2) {
                    g.append("text")
                        .attr("x", x + w / 2).attr("y", sy + sh / 2)
                        .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                        .style("font-size", `${lblFz}px`).style("fill", lblColor)
                        .style("font-weight", boldLbls ? "bold" : "normal")
                        .style("font-family", fontFam)
                        .text(normalize ? `${v.toFixed(1)}%` : this.fmtNum(v));
                }
                yOffsets[i] = y1;
            });
        });
    }

    // ── Stacked horizontal ───────────────────────────────────────────────────
    private renderStackedHorizontal(args: Record<string, unknown>, normalize: boolean): void {
        const { g, data, innerW, innerH, colors, cornerR, barPad, showLbl, lblFz, lblColor, boldLbls, fontFam, boldAxis, italicAxis, xFz, yFz, showGrid, gridColor, hasComp } =
            args as { g: d3.Selection<SVGGElement, unknown, null, undefined>; data: DataPoint[]; innerW: number; innerH: number; colors: string[]; cornerR: number; barPad: number; showLbl: boolean; lblFz: number; lblFormat: string; lblColor: string; boldLbls: boolean; fontFam: string; boldAxis: boolean; italicAxis: boolean; xFz: number; yFz: number; showGrid: boolean; gridColor: string; showZero: boolean; hasComp: boolean };

        const seriesKeys = hasComp ? ["value", "comparison"] : ["value"];
        const catScale = d3.scaleBand().domain(data.map(d => d.category)).range([0, innerH]).padding(barPad);
        const totals = data.map(d => (d.value ?? 0) + (hasComp ? (d.comparison ?? 0) : 0));
        const maxTotal = normalize ? 100 : Math.max(...totals, 0);
        const xScale = d3.scaleLinear().domain([0, maxTotal * 1.05]).range([0, innerW]).nice();

        if (showGrid) {
            g.append("g").selectAll<SVGLineElement, number>("line").data(xScale.ticks(5)).join("line")
                .attr("x1", d => xScale(d)).attr("x2", d => xScale(d))
                .attr("y1", 0).attr("y2", innerH)
                .attr("stroke", gridColor).attr("stroke-width", 0.5).attr("stroke-dasharray", "3,3");
        }

        const yAxis = d3.axisLeft(catScale).tickSizeOuter(0);
        const yG = g.append("g").attr("class", "axis-y").call(yAxis);
        this.applyAxisStyle(yG, yFz, fontFam, boldAxis, italicAxis);
        const xAxis = d3.axisBottom(xScale).ticks(5).tickFormat(d => normalize ? `${d}%` : this.fmtNum(Number(d)));
        const xG = g.append("g").attr("class", "axis-x").attr("transform", `translate(0,${innerH})`).call(xAxis);
        this.applyAxisStyle(xG, xFz, fontFam, boldAxis, italicAxis);

        let xOffsets = new Array(data.length).fill(0);
        seriesKeys.forEach((key, si) => {
            data.forEach((d, i) => {
                const raw = key === "value" ? d.value : (d.comparison ?? 0);
                const totalVal = totals[i] || 1;
                const v = normalize ? (raw / totalVal) * 100 : raw;
                const y = catScale(d.category) ?? 0;
                const h = catScale.bandwidth();
                const x0 = xOffsets[i];
                const bw = xScale(v);
                const r = si === seriesKeys.length - 1 ? Math.min(cornerR, bw / 2, h / 2) : 0;
                const path = r > 0
                    ? `M${x0},${y} L${x0 + bw - r},${y} Q${x0 + bw},${y} ${x0 + bw},${y + r} L${x0 + bw},${y + h - r} Q${x0 + bw},${y + h} ${x0 + bw - r},${y + h} L${x0},${y + h} Z`
                    : `M${x0},${y} L${x0 + bw},${y} L${x0 + bw},${y + h} L${x0},${y + h} Z`;
                g.append("path").attr("d", path).attr("fill", (colors as string[])[si] ?? "#0D9488").attr("opacity", 0.9);
                if (showLbl && bw > 30) {
                    g.append("text")
                        .attr("x", x0 + bw / 2).attr("y", y + h / 2)
                        .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                        .style("font-size", `${lblFz}px`).style("fill", lblColor)
                        .style("font-weight", boldLbls ? "bold" : "normal")
                        .style("font-family", fontFam)
                        .text(normalize ? `${v.toFixed(1)}%` : this.fmtNum(v));
                }
                xOffsets[i] = x0 + bw;
            });
        });
    }

    // ── Vertical bars ─────────────────────────────────────────────────────────
    private renderVertical(args: Record<string, unknown>): void {
        const { g, data, innerW, innerH, colors, cornerR, barPad, showLbl, lblFz, lblFormat, lblColor, boldLbls, fontFam, boldAxis, italicAxis, xFz, yFz, showGrid, gridColor, showZero, hasComp } =
            args as { g: d3.Selection<SVGGElement, unknown, null, undefined>; data: DataPoint[]; innerW: number; innerH: number; colors: string[]; cornerR: number; barPad: number; showLbl: boolean; lblFz: number; lblFormat: string; lblColor: string; boldLbls: boolean; fontFam: string; boldAxis: boolean; italicAxis: boolean; xFz: number; yFz: number; showGrid: boolean; gridColor: string; showZero: boolean; hasComp: boolean };

        const series = hasComp ? ["primary", "comparison"] : ["primary"];

        const catScale = d3.scaleBand()
            .domain(data.map(d => d.category))
            .range([0, innerW])
            .padding(barPad);

        const grpScale = d3.scaleBand()
            .domain(series)
            .range([0, catScale.bandwidth()])
            .padding(0.08);

        const allVals = data.flatMap(d => hasComp ? [d.value, d.comparison ?? 0] : [d.value]);
        const maxVal  = Math.max(...allVals, 0);

        const valScale = d3.scaleLinear()
            .domain([0, maxVal * 1.1])
            .range([innerH, 0])
            .nice();

        if (showGrid) {
            g.append("g").attr("class", "gridlines")
                .selectAll<SVGLineElement, number>("line")
                .data(valScale.ticks(5)).join("line")
                .attr("x1", 0).attr("x2", innerW)
                .attr("y1", d => valScale(d)).attr("y2", d => valScale(d))
                .attr("stroke", gridColor).attr("stroke-width", 0.5).attr("stroke-dasharray", "3,3");
        }
        if (showZero) {
            g.append("line").attr("x1", 0).attr("x2", innerW)
                .attr("y1", valScale(0)).attr("y2", valScale(0))
                .attr("stroke", "#374151").attr("stroke-width", 1.5);
        }

        const yG = g.append("g").attr("class", "axis-y")
            .call(d3.axisLeft(valScale).ticks(5).tickFormat(d => this.fmtNum(Number(d))));
        this.applyAxisStyle(yG, yFz, fontFam, boldAxis, italicAxis);

        const xG = g.append("g").attr("class", "axis-x")
            .attr("transform", `translate(0,${innerH})`)
            .call(d3.axisBottom(catScale).tickSizeOuter(0));
        this.applyAxisStyle(xG, xFz, fontFam, boldAxis, italicAxis);
        if (data.length > 6) {
            xG.selectAll("text").attr("transform", "rotate(-45)").attr("text-anchor", "end").attr("dx", "-0.4em").attr("dy", "0.6em");
        }

        const entries: BarEntry[] = data.flatMap(pt => {
            const c0 = (colors as string[])[0] ?? "#0D9488";
            const c1 = (colors as string[])[1] ?? "#F97316";
            const arr: BarEntry[] = [{ pt, series: "primary", color: c0, value: pt.value }];
            if (hasComp && pt.comparison !== undefined) {
                arr.push({ pt, series: "comparison", color: c1, value: pt.comparison });
            }
            return arr;
        });

        const self = this;
        const totalSum = data.reduce((a, d) => a + d.value, 0) || 1;

        const barPaths = g.append("g").attr("class", "bars")
            .selectAll<SVGPathElement, BarEntry>("path")
            .data(entries).join("path")
            .attr("class", "bar-path")
            .attr("fill", e => e.color)
            .attr("d", "");

        barPaths
            .on("mouseover", function(event: MouseEvent, e: BarEntry) {
                d3.selectAll<SVGPathElement, BarEntry>(".bar-path")
                    .filter(b => b !== e)
                    .transition().duration(150).style("opacity", "0.5");
                self.ttName.textContent  = e.pt.category + (hasComp ? ` (${e.series})` : "");
                self.ttValue.textContent = self.fmtNum(e.value);
                self.tooltip.style.display = "block";
                self.positionTooltip(event);
            })
            .on("mousemove", function(event: MouseEvent) {
                self.positionTooltip(event);
            })
            .on("mouseout", function() {
                d3.selectAll<SVGPathElement, BarEntry>(".bar-path")
                    .transition().duration(150).style("opacity", "1");
                self.tooltip.style.display = "none";
            })
            .on("click", function(event: MouseEvent, e: BarEntry) {
                self.selectionManager.select(e.pt.selectionId, event.ctrlKey || event.metaKey);
                event.stopPropagation();
            })
            .on("contextmenu", function(event: MouseEvent, e: BarEntry) {
                event.preventDefault();
                event.stopPropagation();
                self.selectionManager.showContextMenu(
                    e.pt.selectionId,
                    { x: event.clientX, y: event.clientY }
                );
            });

        barPaths.transition().duration(650).ease(d3.easeCubicOut)
            .attrTween("d", function(e: BarEntry) {
                const gx   = (catScale(e.pt.category) ?? 0) + (grpScale(e.series) ?? 0);
                const bw   = grpScale.bandwidth();
                const endY = valScale(Math.max(0, e.value));
                const endH = innerH - endY;
                const interp = d3.interpolate(0, endH);
                return (t: number) => topRoundedPath(gx, innerH - interp(t), bw, interp(t), cornerR);
            });

        if (showLbl) {
            const lblG = g.append("g").attr("class", "labels");
            entries.forEach(e => {
                const gx  = (catScale(e.pt.category) ?? 0) + (grpScale(e.series) ?? 0);
                const bw  = grpScale.bandwidth();
                const yv  = valScale(Math.max(0, e.value));
                lblG.append("text")
                    .attr("class", "bar-label")
                    .attr("x", gx + bw / 2)
                    .attr("y", yv - 4)
                    .attr("text-anchor", "middle")
                    .style("font-size", `${lblFz}px`)
                    .style("fill", lblColor)
                    .style("font-weight", boldLbls ? "bold" : "normal")
                    .style("font-family", fontFam)
                    .attr("opacity", 0)
                    .text(this.makeLabelText(e.value, totalSum, lblFormat))
                    .transition().delay(650).duration(200).attr("opacity", 1);
            });
        }
    }

    // ── Horizontal bars ───────────────────────────────────────────────────────
    private renderHorizontal(args: Record<string, unknown>): void {
        const { g, data, innerW, innerH, colors, cornerR, barPad, showLbl, lblFz, lblFormat, lblColor, boldLbls, fontFam, boldAxis, italicAxis, xFz, yFz, showGrid, gridColor, showZero, hasComp } =
            args as { g: d3.Selection<SVGGElement, unknown, null, undefined>; data: DataPoint[]; innerW: number; innerH: number; colors: string[]; cornerR: number; barPad: number; showLbl: boolean; lblFz: number; lblFormat: string; lblColor: string; boldLbls: boolean; fontFam: string; boldAxis: boolean; italicAxis: boolean; xFz: number; yFz: number; showGrid: boolean; gridColor: string; showZero: boolean; hasComp: boolean };

        const series = hasComp ? ["primary", "comparison"] : ["primary"];

        const catScale = d3.scaleBand()
            .domain(data.map(d => d.category))
            .range([0, innerH])
            .padding(barPad);

        const grpScale = d3.scaleBand()
            .domain(series)
            .range([0, catScale.bandwidth()])
            .padding(0.08);

        const allVals = data.flatMap(d => hasComp ? [d.value, d.comparison ?? 0] : [d.value]);
        const maxVal  = Math.max(...allVals, 0);

        const valScale = d3.scaleLinear()
            .domain([0, maxVal * 1.1])
            .range([0, innerW])
            .nice();

        if (showGrid) {
            g.append("g").attr("class", "gridlines")
                .selectAll<SVGLineElement, number>("line")
                .data(valScale.ticks(5)).join("line")
                .attr("x1", d => valScale(d)).attr("x2", d => valScale(d))
                .attr("y1", 0).attr("y2", innerH)
                .attr("stroke", gridColor).attr("stroke-width", 0.5).attr("stroke-dasharray", "3,3");
        }
        if (showZero) {
            g.append("line").attr("x1", valScale(0)).attr("x2", valScale(0))
                .attr("y1", 0).attr("y2", innerH)
                .attr("stroke", "#374151").attr("stroke-width", 1.5);
        }

        const yG = g.append("g").attr("class", "axis-y")
            .call(d3.axisLeft(catScale).tickSizeOuter(0));
        this.applyAxisStyle(yG, yFz, fontFam, boldAxis, italicAxis);

        const xG = g.append("g").attr("class", "axis-x")
            .attr("transform", `translate(0,${innerH})`)
            .call(d3.axisBottom(valScale).ticks(5).tickFormat(d => this.fmtNum(Number(d))));
        this.applyAxisStyle(xG, xFz, fontFam, boldAxis, italicAxis);

        const entries: BarEntry[] = data.flatMap(pt => {
            const c0 = (colors as string[])[0] ?? "#0D9488";
            const c1 = (colors as string[])[1] ?? "#F97316";
            const arr: BarEntry[] = [{ pt, series: "primary", color: c0, value: pt.value }];
            if (hasComp && pt.comparison !== undefined) {
                arr.push({ pt, series: "comparison", color: c1, value: pt.comparison });
            }
            return arr;
        });

        const self = this;
        const totalSum = data.reduce((a, d) => a + d.value, 0) || 1;

        const barPaths = g.append("g").attr("class", "bars")
            .selectAll<SVGPathElement, BarEntry>("path")
            .data(entries).join("path")
            .attr("class", "bar-path")
            .attr("fill", e => e.color)
            .attr("d", "");

        barPaths
            .on("mouseover", function(event: MouseEvent, e: BarEntry) {
                d3.selectAll<SVGPathElement, BarEntry>(".bar-path")
                    .filter(b => b !== e)
                    .transition().duration(150).style("opacity", "0.5");
                self.ttName.textContent  = e.pt.category + (hasComp ? ` (${e.series})` : "");
                self.ttValue.textContent = fmtVal(e.value);
                self.tooltip.style.display = "block";
                self.positionTooltip(event);
            })
            .on("mousemove", function(event: MouseEvent) {
                self.positionTooltip(event);
            })
            .on("mouseout", function() {
                d3.selectAll<SVGPathElement, BarEntry>(".bar-path")
                    .transition().duration(150).style("opacity", "1");
                self.tooltip.style.display = "none";
            })
            .on("click", function(event: MouseEvent, e: BarEntry) {
                self.selectionManager.select(e.pt.selectionId, event.ctrlKey || event.metaKey);
                event.stopPropagation();
            })
            .on("contextmenu", function(event: MouseEvent, e: BarEntry) {
                event.preventDefault();
                event.stopPropagation();
                self.selectionManager.showContextMenu(
                    e.pt.selectionId,
                    { x: event.clientX, y: event.clientY }
                );
            });

        barPaths.transition().duration(650).ease(d3.easeCubicOut)
            .attrTween("d", function(e: BarEntry) {
                const gy  = (catScale(e.pt.category) ?? 0) + (grpScale(e.series) ?? 0);
                const bh  = grpScale.bandwidth();
                const endW = valScale(Math.max(0, e.value));
                const interp = d3.interpolate(0, endW);
                return (t: number) => rightRoundedPath(0, gy, interp(t), bh, cornerR);
            });

        const totalSum2 = data.reduce((a, d) => a + d.value, 0) || 1;
        if (showLbl) {
            const lblG = g.append("g").attr("class", "labels");
            entries.forEach(e => {
                const gy  = (catScale(e.pt.category) ?? 0) + (grpScale(e.series) ?? 0);
                const bh  = grpScale.bandwidth();
                const xv  = valScale(Math.max(0, e.value));
                lblG.append("text")
                    .attr("class", "bar-label")
                    .attr("text-anchor", "start")
                    .attr("x", xv + 4)
                    .attr("y", gy + bh / 2)
                    .attr("dominant-baseline", "middle")
                    .style("font-size", `${lblFz}px`)
                    .style("fill", lblColor)
                    .style("font-weight", boldLbls ? "bold" : "normal")
                    .style("font-family", fontFam)
                    .attr("opacity", 0)
                    .text(this.makeLabelText(e.value, totalSum2, lblFormat))
                    .transition().delay(650).duration(200).attr("opacity", 1);
            });
        }
    }

    // ── Tooltip positioning ───────────────────────────────────────────────────
    private positionTooltip(event: MouseEvent): void {
        const areaRect = this.contentEl.getBoundingClientRect();
        const ttW = this.tooltip.offsetWidth  || 100;
        const ttH = this.tooltip.offsetHeight || 44;
        let tx = event.clientX - areaRect.left + 12;
        let ty = event.clientY - areaRect.top  - ttH / 2;
        if (tx + ttW > areaRect.width)  tx = event.clientX - areaRect.left - ttW - 12;
        if (ty < 0)                      ty = 4;
        if (ty + ttH > areaRect.height)  ty = areaRect.height - ttH - 4;
        this.tooltip.style.left = `${tx}px`;
        this.tooltip.style.top  = `${ty}px`;
    }

    // ── Formatting model ──────────────────────────────────────────────────────

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
        return this.fmtService.buildFormattingModel(this.settings);
    }
}
