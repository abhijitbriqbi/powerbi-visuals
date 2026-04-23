"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import * as d3 from "d3";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions      = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual                  = powerbi.extensibility.visual.IVisual;
import IVisualHost              = powerbi.extensibility.visual.IVisualHost;

import { VisualFormattingSettingsModel } from "./settings";
import { checkMicrosoftLicence, resetLicenceCache } from "./licenceManager";
import { getTrialDaysRemaining, isTrialExpired, getPurchaseUrl, getButtonText } from "./trialManager";

const CHART_COLORS = [
    "#0D9488","#F97316","#3B82F6","#8B5CF6","#10B981",
    "#EF4444","#F59E0B","#EC4899","#06B6D4","#84CC16"
];

const TRIAL_MS   = 4 * 24 * 60 * 60 * 1000;
const TRIAL_KEY  = "briqlab_trial_BriqlabMekkoChart_start";
const CACHED_KEY = "briqlab_mekkochart_prokey";

interface MekkoCell {
    xCategory: string;
    segment:   string;
    xValue:    number;
    yValue:    number;
}

function fmtNum(v: number): string {
    if (Math.abs(v) >= 1e6) return `${(v / 1e6).toFixed(1)}M`;
    if (Math.abs(v) >= 1e3) return `${(v / 1e3).toFixed(1)}K`;
    return v.toLocaleString();
}

export class Visual implements IVisual {
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private selMgr!:      powerbi.extensibility.ISelectionManager;
    private _handlersAttached = false;
    private readonly host: IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private readonly fmtSvc: FormattingSettingsService;
    private settings!: VisualFormattingSettingsModel;

    private root!:      HTMLElement;
    private contentEl!: HTMLElement;
    private svgEl!:     d3.Selection<SVGSVGElement, unknown, null, undefined>;
    private trialBadge!: HTMLElement;
    private proBadge!:   HTMLElement;
    private overlay!:    HTMLElement;

    private vp: powerbi.IViewport = { width: 400, height: 300 };
    private curKey = "";
    private isPro  = false;
    private keyCache: Map<string, boolean> = new Map();
    private trialStart = 0;
    private cells: MekkoCell[] = [];

    constructor(options: VisualConstructorOptions) {
        this.host   = options.host;
        this.renderingManager = options.host.eventService;
        this.selMgr     = this.host.createSelectionManager();
        this.tooltipSvc = this.host.tooltipService;
        this.fmtSvc = new FormattingSettingsService();
        this.root   = options.element;
        this.root.classList.add("briqlab-mekko");
        this.buildDOM();
        this.initTrial();
    }

    private buildDOM(): void {
        this.contentEl = document.createElement("div");
        this.contentEl.className = "mekko-content";
        this.root.appendChild(this.contentEl);

        this.svgEl = d3.select(this.contentEl)
            .append<SVGSVGElement>("svg")
            .attr("class", "mekko-svg");

        this.trialBadge = document.createElement("div");
        this.trialBadge.className = "briq-trial-badge hidden";
        this.root.appendChild(this.trialBadge);

        this.proBadge = document.createElement("div");
        this.proBadge.className = "briq-pro-badge hidden";
        this.proBadge.textContent = "✓ Pro Active";
        this.root.appendChild(this.proBadge);

        this.overlay = document.createElement("div");
        this.overlay.className = "briq-trial-overlay hidden";
        this.root.appendChild(this.overlay);
        this.buildOverlay();
    }

    private buildOverlay(): void {
        const card = document.createElement("div");
        card.className = "briq-trial-card";

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

        this.overlay.appendChild(card);
    }

    private initTrial(): void {
        try {
            const s = localStorage.getItem(TRIAL_KEY);
            this.trialStart = s ? parseInt(s, 10) : Date.now();
            if (!s) localStorage.setItem(TRIAL_KEY, String(this.trialStart));
            const k = localStorage.getItem(CACHED_KEY);
            if (k) {
                this.curKey = k;
                this.validateKey(k).then(ok => { this.isPro = ok; this.render(); });
            }
        } catch {
            this.trialStart = Date.now();
        }
    }

    private isExpired(): boolean { return Date.now() - this.trialStart >= TRIAL_MS; }
    private daysLeft():  number  { return Math.max(0, 4 - Math.floor((Date.now() - this.trialStart) / (24 * 60 * 60 * 1000))); }

    private async validateKey(key: string): Promise<boolean> {
        return checkMicrosoftLicence(this.host);
    }

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
                        dataItems: [{ displayName: "Briqlab Mekko Chart", value: "" }],
                        identities: [],
                        coordinates: [e.clientX, e.clientY],
                        isTouchEvent: false
                    });
                });
                this.root.addEventListener("mouseleave", () => {
                    this.tooltipSvc.hide({ isTouchEvent: false, immediately: false });
                });
            }
            this.vp = options.viewport;
            this.settings = this.fmtSvc.populateFormattingSettingsModel(
                VisualFormattingSettingsModel, options.dataViews?.[0]
            );
    
            this.cells = [];
            const dv = options.dataViews?.[0];
            if (dv?.categorical) {
                const cats = dv.categorical.categories ?? [];
                const vals = dv.categorical.values    ?? [];
                const xCatCol  = cats.find(c => (c.source.roles as Record<string, unknown>)["xCategory"]);
                const segCol   = cats.find(c => (c.source.roles as Record<string, unknown>)["segment"]);
                const xValCol  = vals.find(c => (c.source.roles as Record<string, unknown>)["xValue"]);
                const yValCol  = vals.find(c => (c.source.roles as Record<string, unknown>)["yValue"]);
    
                if (xCatCol && segCol) {
                    for (let i = 0; i < xCatCol.values.length; i++) {
                        const xv = xValCol ? Number(xValCol.values[i]) : 0;
                        const yv = yValCol ? Number(yValCol.values[i]) : 0;
                        this.cells.push({
                            xCategory: String(xCatCol.values[i] ?? ""),
                            segment:   String(segCol.values[i]  ?? ""),
                            xValue:    isNaN(xv) ? 0 : xv,
                            yValue:    isNaN(yv) ? 0 : yv
                        });
                    }
                }
            }
    
            const key = ""; // MS cert: pro key field removed
            if (key && key !== this.curKey) {
                this.curKey = key;
                this.isPro  = false;
                this.validateKey(key).then(ok => { this.isPro = ok; this.render(); });
            }
    
            this.render();
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    private render(): void {
        if (!this.settings) return;

        const expired = this.isExpired();
        if (!this.isPro && !expired) {
            this.trialBadge.textContent = `Trial: ${this.daysLeft()} days remaining`;
            this.trialBadge.classList.remove("hidden");
        } else {
            this.trialBadge.classList.add("hidden");
        }
        this.isPro ? this.proBadge.classList.remove("hidden") : this.proBadge.classList.add("hidden");
        if (expired && !this.isPro) {
            this.contentEl.classList.add("blurred");
            this.overlay.classList.remove("hidden");
        } else {
            this.contentEl.classList.remove("blurred");
            this.overlay.classList.add("hidden");
        }

        const cs = this.settings.chartSettings;
        const ls = this.settings.legendSettings;
        const { width, height } = this.vp;

        const barGap           = Math.min(8, Math.max(0, cs.barGap.value ?? 2));
        const showSegLabels    = cs.showSegmentLabels.value;
        const labelThreshold   = cs.labelThreshold.value ?? 8;
        const showXScale       = cs.showXScale.value;
        const fontFamily       = String(cs.fontFamily?.value?.value ?? "Segoe UI");
        const showLegend       = ls.showLegend.value;
        const legendPosition   = String(ls.legendPosition?.value?.value ?? "Bottom");

        this.svgEl.attr("width", width).attr("height", height);
        this.svgEl.selectAll("*").remove();

        if (this.cells.length === 0) {
            this.svgEl.append("text")
                .attr("x", width / 2).attr("y", height / 2)
                .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                .attr("class", "mekko-empty").attr("font-family", fontFamily)
                .text("Add X Category, Segment, X Value & Y Value fields");
            return;
        }

        const xCategories = Array.from(new Set(this.cells.map(c => c.xCategory)));
        const segments    = Array.from(new Set(this.cells.map(c => c.segment)));
        const segColor    = new Map<string, string>();
        segments.forEach((seg, i) => segColor.set(seg, CHART_COLORS[i % CHART_COLORS.length]));

        // Per xCategory: xValue (bar width), yMap (segment -> yValue sum)
        const xValueMap  = new Map<string, number>();
        const yValueMap  = new Map<string, Map<string, number>>();
        for (const cell of this.cells) {
            if (!xValueMap.has(cell.xCategory) || cell.xValue > 0) {
                xValueMap.set(cell.xCategory, cell.xValue);
            }
            let yMap = yValueMap.get(cell.xCategory);
            if (!yMap) { yMap = new Map(); yValueMap.set(cell.xCategory, yMap); }
            yMap.set(cell.segment, (yMap.get(cell.segment) ?? 0) + cell.yValue);
        }

        const totalXValue = Array.from(xValueMap.values()).reduce((a, b) => a + b, 0);

        const legendH = showLegend ? 24 : 0;
        const xAxisH  = showXScale ? 36 : 0;
        const padL    = 4; const padR = 4; const padT = 4;

        const chartTop    = legendPosition === "Top" && showLegend ? padT + legendH : padT;
        const chartBottom = legendPosition === "Bottom" && showLegend ? height - xAxisH - legendH : height - xAxisH;
        const chartH      = chartBottom - chartTop;
        const chartW      = width - padL - padR;

        const g = this.svgEl.append("g").attr("transform", `translate(${padL},${chartTop})`);

        let xCursor = 0;
        xCategories.forEach(xCat => {
            const xv   = xValueMap.get(xCat) ?? 0;
            const barW = totalXValue > 0 ? (xv / totalXValue) * (chartW - barGap * (xCategories.length - 1)) : 0;
            const yMap = yValueMap.get(xCat) ?? new Map<string, number>();
            const totalY = Array.from(yMap.values()).reduce((a, b) => a + b, 0);

            let yCursor = 0;
            segments.forEach(seg => {
                const yv  = yMap.get(seg) ?? 0;
                if (yv <= 0) return;
                const segH = totalY > 0 ? (yv / totalY) * chartH : 0;
                const color = segColor.get(seg) ?? "#0D9488";
                const pctOfCat   = totalY > 0 ? (yv / totalY) * 100 : 0;
                const pctOfTotal = totalY > 0 && totalXValue > 0 ? (yv * xv) / (totalY * totalXValue) * 100 : 0;

                const rect = g.append("rect")
                    .attr("x", xCursor)
                    .attr("y", yCursor)
                    .attr("width", Math.max(0, barW))
                    .attr("height", Math.max(0, segH))
                    .attr("fill", color)
                    .attr("class", "mekko-segment");

                rect.append("title")
                    .text(`${xCat} / ${seg}\nValue: ${fmtNum(yv)}\n% of category: ${pctOfCat.toFixed(1)}%\n% of total: ${pctOfTotal.toFixed(1)}%`);

                // Segment label
                if (showSegLabels && barW > 40 && pctOfCat >= labelThreshold && segH > 14) {
                    g.append("text")
                        .attr("x", xCursor + barW / 2)
                        .attr("y", yCursor + segH / 2)
                        .attr("text-anchor", "middle")
                        .attr("dominant-baseline", "middle")
                        .attr("class", "mekko-seg-label")
                        .attr("font-family", fontFamily)
                        .text(`${pctOfCat.toFixed(0)}%`);
                }

                yCursor += segH;
            });

            // X axis label
            if (showXScale) {
                g.append("text")
                    .attr("x", xCursor + barW / 2)
                    .attr("y", chartH + 14)
                    .attr("text-anchor", "middle")
                    .attr("class", "mekko-xcat-label")
                    .attr("font-family", fontFamily)
                    .text(xCat);

                g.append("text")
                    .attr("x", xCursor + barW / 2)
                    .attr("y", chartH + 28)
                    .attr("text-anchor", "middle")
                    .attr("class", "mekko-xval-label")
                    .attr("font-family", fontFamily)
                    .text(fmtNum(xv));
            }

            xCursor += barW + barGap;
        });

        // Legend
        if (showLegend) {
            const legY = legendPosition === "Top"
                ? padT
                : height - legendH + 4;
            const totalLegW = segments.length * 90;
            let legX = Math.max(4, (width - totalLegW) / 2);
            segments.forEach(seg => {
                const color = segColor.get(seg) ?? "#0D9488";
                this.svgEl.append("rect")
                    .attr("x", legX).attr("y", legY + 2)
                    .attr("width", 10).attr("height", 10)
                    .attr("rx", 2).attr("fill", color);
                this.svgEl.append("text")
                    .attr("x", legX + 14).attr("y", legY + 7)
                    .attr("dominant-baseline", "middle")
                    .attr("class", "mekko-legend-label")
                    .attr("font-family", fontFamily)
                    .text(seg);
                legX += 90;
            });
        }
    }

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
