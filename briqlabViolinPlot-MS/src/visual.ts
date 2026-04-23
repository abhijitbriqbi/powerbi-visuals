"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";
import * as d3 from "d3";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;
import DataView = powerbi.DataView;

import { VisualFormattingSettingsModel } from "./settings";
import { checkMicrosoftLicence, resetLicenceCache } from "./licenceManager";
import { getTrialDaysRemaining, isTrialExpired, getPurchaseUrl, getButtonText } from "./trialManager";

const CHART_COLORS = [
    "#0D9488","#F97316","#3B82F6","#8B5CF6","#10B981",
    "#EF4444","#F59E0B","#EC4899","#06B6D4","#84CC16"
];

const TRIAL_MS      = 4 * 24 * 60 * 60 * 1000;
const TRIAL_KEY     = "briqlab_trial_ViolinPlot_start";
const PRO_STORE_KEY = "briqlab_violinplot_prokey";

interface ViolinStats {
    category:      string;
    values:        number[];
    mean:          number;
    median:        number;
    q1:            number;
    q3:            number;
    stdDev:        number;
    min:           number;
    max:           number;
    whiskerLow:    number;
    whiskerHigh:   number;
    outliers:      number[];
    densityPoints: Array<[number, number]>;
    selectionId:   ISelectionId;
}

function getTrial(): { daysLeft: number; expired: boolean } {
    try {
        let raw = localStorage.getItem(TRIAL_KEY);
        if (!raw) { raw = String(Date.now()); localStorage.setItem(TRIAL_KEY, raw); }
        const elapsed  = Date.now() - parseInt(raw, 10);
        const daysLeft = Math.max(0, Math.ceil((TRIAL_MS - elapsed) / 86400000));
        return { daysLeft, expired: elapsed > TRIAL_MS };
    } catch { return { daysLeft: 4, expired: false }; }
}

function computeKDE(values: number[], bandwidth: number, points: number[]): number[] {
    return points.map(x => {
        const sum = values.reduce((acc,v) => {
            const u = (x-v)/bandwidth;
            return acc + Math.exp(-0.5*u*u) / Math.sqrt(2*Math.PI);
        }, 0);
        return sum / (values.length * bandwidth);
    });
}

function computeStats(category: string, rawValues: number[], bwMode: string, selId: ISelectionId): ViolinStats {
    const values  = rawValues.slice().sort((a,b) => a-b);
    const n       = values.length;
    const mean    = values.reduce((s,v) => s+v, 0) / n;
    const median  = n % 2 === 0 ? (values[n/2-1]+values[n/2])/2 : values[Math.floor(n/2)];
    const q1      = values[Math.floor(n*0.25)];
    const q3      = values[Math.floor(n*0.75)];
    const variance= values.reduce((s,v) => s+(v-mean)**2, 0) / n;
    const stdDev  = Math.sqrt(variance);
    const iqr     = q3 - q1;
    const whiskerLow  = Math.max(values[0],  q1 - 1.5*iqr);
    const whiskerHigh = Math.min(values[n-1], q3 + 1.5*iqr);
    const outliers= values.filter(v => v < whiskerLow || v > whiskerHigh);
    let bw        = 1.06 * stdDev * Math.pow(n, -0.2);
    if (bwMode === "Fine")   bw *= 0.5;
    if (bwMode === "Coarse") bw *= 2.0;
    if (bw === 0) bw = 0.01;
    const minVal  = values[0];
    const maxVal  = values[n-1];
    const range   = maxVal - minVal || 1;
    const samplePts: number[] = d3.range(50).map(i => minVal + (i/49)*range);
    const densities = computeKDE(values, bw, samplePts);
    const densityPoints: Array<[number,number]> = samplePts.map((y,i) => [y, densities[i]]);
    return { category, values, mean, median, q1, q3, stdDev, min: minVal, max: maxVal, whiskerLow, whiskerHigh, outliers, densityPoints, selectionId: selId };
}

export class Visual implements IVisual {
    private readonly host:    IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private readonly selMgr:  ISelectionManager;
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private _handlersAttached = false;
    private readonly fmtSvc:  FormattingSettingsService;

    private readonly root:       HTMLElement;
    private readonly contentEl:  HTMLDivElement;
    private readonly chartEl:    HTMLDivElement;
    private readonly tooltipEl:  HTMLDivElement;
    private readonly trialBadge: HTMLDivElement;
    private readonly proBadge:   HTMLDivElement;
    private readonly keyErrorEl: HTMLDivElement;
    private readonly overlayEl:  HTMLDivElement;

    private settings!:   VisualFormattingSettingsModel;
    private vp:          powerbi.IViewport = { width: 400, height: 300 };
    private selectedCats:Set<string> = new Set();

    private isPro   = false;
    private lastKey = "";
    private readonly keyCache: Map<string, boolean> = new Map();

    constructor(options: VisualConstructorOptions) {
        this.host   = options.host;
        this.renderingManager = options.host.eventService;
        this.selMgr = options.host.createSelectionManager();
        this.tooltipSvc = this.host.tooltipService;
        this.fmtSvc = new FormattingSettingsService();

        this.root = options.element;
        this.root.classList.add("briqlab-violin");
        this.root.style.position = "relative";
        this.root.style.overflow = "hidden";

        this.contentEl = this.mkDiv("briqlab-visual-content");
        this.root.appendChild(this.contentEl);

        this.chartEl = this.mkDiv("briq-violin-container");
        this.contentEl.appendChild(this.chartEl);

        this.tooltipEl = this.mkDiv("briq-tooltip");
        this.tooltipEl.style.display = "none";
        this.root.appendChild(this.tooltipEl);

        this.trialBadge = this.mkDiv("briqlab-trial-badge hidden");
        this.root.appendChild(this.trialBadge);

        this.proBadge = this.mkDiv("briqlab-pro-badge hidden");
        this.proBadge.textContent = "\u2713 Pro Active";
        this.root.appendChild(this.proBadge);

        this.keyErrorEl = this.mkDiv("briqlab-key-error hidden");
        this.keyErrorEl.textContent = "\u2717 Invalid key";
        this.root.appendChild(this.keyErrorEl);

        this.overlayEl = this.mkDiv("briqlab-trial-overlay hidden");
        this.buildTrialOverlay();
        this.root.appendChild(this.overlayEl);

        this.restoreProKey();
    }

    private buildTrialOverlay(): void {
        const card = this.mkDiv("briqlab-trial-card");
        const title = document.createElement("h2"); title.className = "trial-title"; title.textContent = "Free trial ended"; card.appendChild(title);
        const body  = document.createElement("p");  body.className  = "trial-body";  body.textContent  = "Activate Briqlab Pro to continue using this visual and unlock all features."; card.appendChild(body);
        const btn   = document.createElement("button"); btn.className = "trial-btn"; btn.textContent = getButtonText();
        btn.addEventListener("click", () => this.host.launchUrl(getPurchaseUrl())); card.appendChild(btn);
        const sub   = document.createElement("p"); sub.className = "trial-subtext"; sub.textContent = "Purchase on Microsoft AppSource to unlock all features instantly."; card.appendChild(sub);
        this.overlayEl.appendChild(card);
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
                        dataItems: [{ displayName: "Briqlab Violin Plot", value: "" }],
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
            this.settings = this.fmtSvc.populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews?.[0]);
            this.handleProKey();
            this.updateLicenseUI();
    
            const dv: DataView | undefined = options.dataViews?.[0];
            const categorical = dv?.categorical;
            const catColumn   = categorical?.categories?.[0];
            const valColumn   = categorical?.values?.[0];
    
            if (!catColumn || !valColumn) { this.renderEmpty("Add Category and Value fields"); return; }
    
            const settings   = this.settings;
            const bwMode     = String(settings.distributionSettings.kdeBandwidth.value?.value ?? "Auto");
    
            const groups = new Map<string, { values: number[]; firstIdx: number }>();
            catColumn.values.forEach((catVal, i) => {
                const key = catVal != null ? String(catVal) : "(Blank)";
                const num = Number(valColumn.values[i]);
                if (!isNaN(num)) {
                    const entry = groups.get(key) ?? { values: [], firstIdx: i };
                    entry.values.push(num);
                    groups.set(key, entry);
                }
            });
    
            if (groups.size === 0) { this.renderEmpty("No data to display"); return; }
    
            const violinStats: ViolinStats[] = [];
            groups.forEach(({ values, firstIdx }, cat) => {
                if (values.length > 1) {
                    const selId = this.host.createSelectionIdBuilder().withCategory(catColumn, firstIdx).createSelectionId();
                    violinStats.push(computeStats(cat, values, bwMode, selId));
                }
            });
    
            this.renderChart(violinStats);
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    private renderChart(violinStats: ViolinStats[]): void {
        while (this.chartEl.firstChild) this.chartEl.removeChild(this.chartEl.firstChild);

        const settings = this.settings;
        const width    = this.vp.width;
        const height   = this.vp.height;

        const showBoxPlot    = settings.distributionSettings.showBoxPlot.value;
        const showOutliers   = settings.distributionSettings.showOutliers.value;
        const violinWidthPct = settings.distributionSettings.violinWidth.value / 100;
        const fillOpacity    = settings.distributionSettings.fillOpacity.value / 100;
        const fontFamily     = settings.distributionSettings.fontFamily.value || "Segoe UI, sans-serif";
        const showRefLine    = String(settings.statsSettings.showRefLine.value?.value ?? "None");
        const refLineColor   = settings.statsSettings.refLineColor.value?.value || "#F97316";
        const showLabels     = settings.labelSettings.showLabels.value;
        const fontSize       = settings.labelSettings.fontSize.value;

        const padL = 50, padR = 20, padT = 30, padB = 40;
        const chartW = width - padL - padR;
        const chartH = height - padT - padB;

        const allVals = violinStats.flatMap(s => s.values);
        const yScale  = d3.scaleLinear()
            .domain([d3.min(allVals) ?? 0, d3.max(allVals) ?? 1])
            .range([chartH, 0]).nice();

        const svg = d3.select(this.chartEl)
            .append("svg").attr("width",width).attr("height",height).style("font-family",fontFamily);

        const g = svg.append("g").attr("transform",`translate(${padL},${padT})`);

        g.append("g").attr("class","gridlines")
            .selectAll("line").data(yScale.ticks(6)).join("line")
            .attr("x1",0).attr("x2",chartW)
            .attr("y1",d => yScale(d)).attr("y2",d => yScale(d))
            .attr("stroke","#E5E7EB").attr("stroke-dasharray","4,3").attr("stroke-width",1);

        g.append("g").call(d3.axisLeft(yScale).ticks(6).tickSize(3))
            .call(ax => ax.select(".domain").remove())
            .style("font-size",`${fontSize}px`);

        const slotWidth = chartW / violinStats.length;
        const self = this;

        violinStats.forEach((stats, i) => {
            const slotX    = i*slotWidth + slotWidth/2;
            const color    = CHART_COLORS[i % CHART_COLORS.length];
            const maxDensity = d3.max(stats.densityPoints, d => d[1]) ?? 1;
            const maxHalf    = slotWidth*0.5*violinWidthPct;
            const densScale  = d3.scaleLinear().domain([0,maxDensity]).range([0,maxHalf]);
            const areaGen    = d3.area<[number,number]>().y(d => yScale(d[0])).x0(d => -densScale(d[1])).x1(d => densScale(d[1])).curve(d3.curveCatmullRom);
            const catKey     = stats.category;
            const selId      = stats.selectionId;

            const vg = g.append("g").attr("transform",`translate(${slotX},0)`).attr("class","violin-group")
                .style("cursor","pointer")
                .on("click", (event: MouseEvent) => {
                    const isMulti = event.ctrlKey || event.metaKey;
                    if (self.selectedCats.has(catKey) && !isMulti) {
                        self.selectedCats.clear(); self.selMgr.clear();
                    } else {
                        if (!isMulti) self.selectedCats.clear();
                        self.selectedCats.add(catKey);
                        const ids = violinStats
                            .filter(s => self.selectedCats.has(s.category))
                            .map(s => s.selectionId);
                        self.selMgr.select(ids, isMulti);
                    }
                    g.selectAll<SVGGElement, unknown>(".violin-group").style("opacity", (_,ii) => {
                        if (self.selectedCats.size === 0) return null;
                        return self.selectedCats.has(violinStats[ii]?.category) ? "1" : "0.3";
                    });
                    event.stopPropagation();
                });

            vg.append("path").datum(stats.densityPoints).attr("d",areaGen)
                .attr("fill",color).attr("fill-opacity",fillOpacity)
                .attr("stroke",color).attr("stroke-opacity",0.8).attr("stroke-width",1.5);

            if (showBoxPlot) {
                const boxW = 8;
                vg.append("rect").attr("x",-boxW/2).attr("y",yScale(stats.q3))
                    .attr("width",boxW).attr("height",Math.max(0,yScale(stats.q1)-yScale(stats.q3)))
                    .attr("fill",color).attr("fill-opacity",0.6);
                vg.append("line").attr("x1",-boxW/2).attr("x2",boxW/2).attr("y1",yScale(stats.median)).attr("y2",yScale(stats.median)).attr("stroke","white").attr("stroke-width",2);
                vg.append("line").attr("x1",0).attr("x2",0).attr("y1",yScale(stats.whiskerLow)).attr("y2",yScale(stats.q1)).attr("stroke",color).attr("stroke-width",1);
                vg.append("line").attr("x1",0).attr("x2",0).attr("y1",yScale(stats.q3)).attr("y2",yScale(stats.whiskerHigh)).attr("stroke",color).attr("stroke-width",1);
            }

            if (showOutliers) {
                vg.selectAll("circle.outlier").data(stats.outliers).join("circle")
                    .attr("class","outlier").attr("cx",0).attr("cy",d => yScale(d)).attr("r",3)
                    .attr("fill",color).attr("fill-opacity",0.6);
            }

            if (showLabels) {
                g.append("text").attr("x",slotX).attr("y",chartH+20).attr("text-anchor","middle")
                    .attr("font-size",`${fontSize}px`).attr("fill","#374151").text(stats.category);
            }

            vg.append("rect").attr("x",-slotWidth/2).attr("y",0).attr("width",slotWidth).attr("height",chartH)
                .attr("fill","transparent")
                .on("mousemove", function(event: MouseEvent) { self.showTooltip(event, stats, color); })
                .on("mouseleave", function() { self.tooltipEl.style.display = "none"; });

            void selId;
        });

        if (showRefLine !== "None" && violinStats.length > 0) {
            const allData      = violinStats.flatMap(s => s.values);
            const globalMean   = allData.reduce((a,b) => a+b, 0) / allData.length;
            const sorted       = allData.slice().sort((a,b) => a-b);
            const m            = sorted.length;
            const globalMedian = m%2===0 ? (sorted[m/2-1]+sorted[m/2])/2 : sorted[Math.floor(m/2)];
            const drawRef = (val: number, lbl: string) => {
                const y = yScale(val);
                g.append("line").attr("x1",0).attr("x2",chartW).attr("y1",y).attr("y2",y).attr("stroke",refLineColor).attr("stroke-dasharray","6,3").attr("stroke-width",1.5);
                g.append("text").attr("x",chartW-4).attr("y",y-4).attr("text-anchor","end").attr("font-size","10px").attr("fill",refLineColor).text(lbl);
            };
            if (showRefLine === "Mean"   || showRefLine === "Both") drawRef(globalMean,   "Mean");
            if (showRefLine === "Median" || showRefLine === "Both") drawRef(globalMedian, "Median");
        }

        svg.on("click", () => { this.selectedCats.clear(); this.selMgr.clear(); });
    }

    private showTooltip(event: MouseEvent, stats: ViolinStats, color: string): void {
        const fmt = (v: number) => v.toFixed(2);
        while (this.tooltipEl.firstChild) this.tooltipEl.removeChild(this.tooltipEl.firstChild);

        const title = this.mkDiv("briq-tooltip-title");
        title.style.color = color; title.textContent = stats.category;
        this.tooltipEl.appendChild(title);

        const rows: Array<[string,string]> = [["Mean",fmt(stats.mean)],["Median",fmt(stats.median)],["Std Dev",fmt(stats.stdDev)],["Q1",fmt(stats.q1)],["Q3",fmt(stats.q3)],["Min",fmt(stats.min)],["Max",fmt(stats.max)],["N",String(stats.values.length)]];
        rows.forEach(([label,val]) => {
            const row  = this.mkDiv("briq-tooltip-row");
            const lSpan = document.createElement("span"); lSpan.className = "briq-tooltip-label"; lSpan.textContent = label; row.appendChild(lSpan);
            const vSpan = document.createElement("span"); vSpan.className = "briq-tooltip-value"; vSpan.textContent = val;   row.appendChild(vSpan);
            this.tooltipEl.appendChild(row);
        });

        const rect = this.root.getBoundingClientRect();
        this.tooltipEl.style.left    = `${event.clientX - rect.left + 12}px`;
        this.tooltipEl.style.top     = `${event.clientY - rect.top - 10}px`;
        this.tooltipEl.style.display = "block";
    }

    private renderEmpty(msg: string): void {
        while (this.chartEl.firstChild) this.chartEl.removeChild(this.chartEl.firstChild);
        const el = this.mkDiv("briq-empty"); el.textContent = msg; this.chartEl.appendChild(el);
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
            const key = (JSON.parse(stored) as { key?: string })?.key;
            if (!key) return;
            this.lastKey = key;
            this.validateKey(key).then(valid => {
                this.isPro = valid;
                if (!valid) { this.lastKey = ""; try { localStorage.removeItem(PRO_STORE_KEY); } catch { /**/ } }
                this.updateLicenseUI();
            });
        } catch { /**/ }
    }

    private updateLicenseUI(): void {
        checkMicrosoftLicence(this.host).then(p => this._msUpdateLicenceUI(p)).catch(() => this._msUpdateLicenceUI(false));
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

    private mkDiv(cls: string): HTMLDivElement {
        const el = document.createElement("div"); el.className = cls; return el;
    }
}
