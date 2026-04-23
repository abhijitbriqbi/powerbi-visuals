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
const TRIAL_KEY     = "briqlab_trial_SlopeChart_start";
const PRO_STORE_KEY = "briqlab_slopechart_prokey";

interface SlopeEntity {
    name:       string;
    period1:    number;
    period2:    number;
    category:   string;
    pctChange:  number;
    rank1:      number;
    rank2:      number;
    selectionId:ISelectionId;
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

function formatPct(val: number): string {
    return `${val >= 0 ? "+" : ""}${val.toFixed(1)}%`;
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
    private readonly trialBadge: HTMLDivElement;
    private readonly proBadge:   HTMLDivElement;
    private readonly keyErrorEl: HTMLDivElement;
    private readonly overlayEl:  HTMLDivElement;

    private settings!:   VisualFormattingSettingsModel;
    private vp:          powerbi.IViewport = { width: 400, height: 300 };
    private entities:    SlopeEntity[] = [];
    private selectedIds: Set<string>   = new Set();

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
        this.root.classList.add("briqlab-slope");
        this.root.style.position = "relative";
        this.root.style.overflow = "hidden";

        this.contentEl = this.mkDiv("briqlab-visual-content");
        this.root.appendChild(this.contentEl);

        this.chartEl = this.mkDiv("briq-slope-container");
        this.contentEl.appendChild(this.chartEl);

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
                        dataItems: [{ displayName: "Briqlab Slope Chart", value: "" }],
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
            const categorical    = dv?.categorical;
            const entityColumn   = categorical?.categories?.[0];
            const categoryColumn = categorical?.categories?.[1];
            const valColumns     = categorical?.values;
    
            // Find period columns by role (not by index)
            const period1Col = valColumns?.find(c => c.source?.roles?.["period1Value"]);
            const period2Col = valColumns?.find(c => c.source?.roles?.["period2Value"]);
    
            if (!entityColumn || !period1Col || !period2Col) {
                this.renderEmpty("Add Entity, Period 1 & Period 2 fields"); return;
            }
    
            this.entities = [];
            entityColumn.values.forEach((entVal, i) => {
                const name = entVal != null ? String(entVal) : "(Blank)";
                const p1   = Number(period1Col.values[i]);
                const p2   = Number(period2Col.values[i]);
                if (!isNaN(p1) && !isNaN(p2)) {
                    const cat = categoryColumn?.values?.[i] != null ? String(categoryColumn.values[i]) : "";
                    const pctChange = p1 !== 0 ? ((p2-p1) / Math.abs(p1)) * 100 : 0;
                    this.entities.push({
                        name, period1: p1, period2: p2, category: cat, pctChange, rank1: 0, rank2: 0,
                        selectionId: this.host.createSelectionIdBuilder().withCategory(entityColumn, i).createSelectionId()
                    });
                }
            });
    
            if (this.entities.length === 0) { this.renderEmpty("No data to display"); return; }
    
            // Compute ranks
            const sorted1 = this.entities.slice().sort((a,b) => b.period1 - a.period1);
            const sorted2 = this.entities.slice().sort((a,b) => b.period2 - a.period2);
            this.entities.forEach(e => {
                e.rank1 = sorted1.findIndex(s => s.name === e.name) + 1;
                e.rank2 = sorted2.findIndex(s => s.name === e.name) + 1;
            });
    
            this.renderChart();
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    private renderChart(): void {
        while (this.chartEl.firstChild) this.chartEl.removeChild(this.chartEl.firstChild);

        const settings  = this.settings;
        const entities  = this.entities;
        const width     = this.vp.width;
        const height    = this.vp.height;

        const period1Label  = settings.chartSettings.period1Label.value || "Period 1";
        const period2Label  = settings.chartSettings.period2Label.value || "Period 2";
        const colorBy       = String(settings.chartSettings.colorBy.value?.value ?? "Direction");
        const lineWidth     = settings.chartSettings.lineWidth.value;
        const showSlopeLabels = settings.chartSettings.showSlopeLabels.value;
        const showRankChange  = settings.chartSettings.showRankChange.value;
        const dotSize         = settings.chartSettings.dotSize.value;
        const fontFamily      = settings.chartSettings.fontFamily.value || "Segoe UI, sans-serif";
        const showSummary     = settings.summarySettings.showSummary.value;

        const marginL = 120, marginR = 120;
        const padT    = 40;
        const padB    = showSummary ? 50 : 20;
        const chartH  = height - padT - padB;

        const allVals = entities.flatMap(e => [e.period1, e.period2]);
        const yMin    = d3.min(allVals) ?? 0;
        const yMax    = d3.max(allVals) ?? 1;
        const yScale  = d3.scaleLinear()
            .domain([yMin - (yMax-yMin)*0.05, yMax + (yMax-yMin)*0.05])
            .range([chartH, 0]);

        const leftX  = marginL;
        const rightX = width - marginR;
        const catNames = Array.from(new Set(entities.map(e => e.category)));

        const getColor = (e: SlopeEntity, idx: number): string => {
            if (colorBy === "Direction")  return e.period2 > e.period1 ? "#10B981" : "#EF4444";
            if (colorBy === "Category") { const ci = catNames.indexOf(e.category); return CHART_COLORS[Math.max(0,ci) % CHART_COLORS.length]; }
            return CHART_COLORS[idx % CHART_COLORS.length];
        };

        const svg = d3.select(this.chartEl)
            .append("svg").attr("width", width).attr("height", height).style("font-family", fontFamily);

        const g = svg.append("g");

        g.append("text").attr("x",leftX).attr("y",20).attr("text-anchor","middle").attr("font-size","13px").attr("font-weight","600").attr("fill","#374151").text(period1Label);
        g.append("text").attr("x",rightX).attr("y",20).attr("text-anchor","middle").attr("font-size","13px").attr("font-weight","600").attr("fill","#374151").text(period2Label);
        g.append("line").attr("x1",leftX).attr("x2",leftX).attr("y1",padT).attr("y2",padT+chartH).attr("stroke","#E5E7EB").attr("stroke-width",1);
        g.append("line").attr("x1",rightX).attr("x2",rightX).attr("y1",padT).attr("y2",padT+chartH).attr("stroke","#E5E7EB").attr("stroke-width",1);

        // Collision detection
        const labelSpread = (positions: number[], minGap: number): number[] => {
            const adj = positions.slice();
            let changed = true;
            for (let iter = 0; iter < 50 && changed; iter++) {
                changed = false;
                for (let i = 1; i < adj.length; i++) {
                    if (adj[i] - adj[i-1] < minGap) { adj[i-1] -= 2; adj[i] += 2; changed = true; }
                }
            }
            return adj;
        };

        const leftYRaw   = entities.map(e => padT + yScale(e.period1));
        const rightYRaw  = entities.map(e => padT + yScale(e.period2));
        const leftYOrder = entities.map((_,i) => i).sort((a,b) => leftYRaw[a]-leftYRaw[b]);
        const rightYOrder= entities.map((_,i) => i).sort((a,b) => rightYRaw[a]-rightYRaw[b]);
        const leftYAdj   = labelSpread(leftYOrder.map(i => leftYRaw[i]), 14);
        const rightYAdj  = labelSpread(rightYOrder.map(i => rightYRaw[i]), 14);
        const leftYFinal : number[] = new Array(entities.length);
        const rightYFinal: number[] = new Array(entities.length);
        leftYOrder.forEach((oi,si) => { leftYFinal[oi]  = leftYAdj[si]; });
        rightYOrder.forEach((oi,si) => { rightYFinal[oi] = rightYAdj[si]; });

        const self = this;

        entities.forEach((entity, i) => {
            const color    = getColor(entity, i);
            const y1       = padT + yScale(entity.period1);
            const y2       = padT + yScale(entity.period2);
            const totalLen = Math.sqrt((rightX-leftX)**2 + (y2-y1)**2);
            const selKey   = ((entity.selectionId as unknown as Record<string,unknown>)["key"] as string) || entity.name;

            const lineEl = g.append("line")
                .attr("x1",leftX).attr("y1",y1).attr("x2",rightX).attr("y2",y2)
                .attr("stroke",color).attr("stroke-width",lineWidth).attr("stroke-opacity",0.7)
                .attr("stroke-dasharray",totalLen).attr("stroke-dashoffset",totalLen)
                .style("cursor","pointer");

            lineEl.transition().duration(600).delay(i*40).attr("stroke-dashoffset",0);

            g.append("circle").attr("cx",leftX).attr("cy",y1).attr("r",dotSize/2).attr("fill",color);
            g.append("circle").attr("cx",rightX).attr("cy",y2).attr("r",dotSize/2).attr("fill",color);
            g.append("text").attr("x",leftX-dotSize/2-4).attr("y",leftYFinal[i]+4).attr("text-anchor","end").attr("font-size","11px").attr("fill","#374151").text(entity.name);

            const rankDelta = entity.rank1 - entity.rank2;
            const rankStr   = showRankChange && rankDelta !== 0 ? ` (${rankDelta > 0 ? "\u25B2" : "\u25BC"}${Math.abs(rankDelta)})` : "";
            g.append("text").attr("x",rightX+dotSize/2+4).attr("y",rightYFinal[i]+4).attr("text-anchor","start").attr("font-size","11px").attr("fill","#374151")
                .text(`${entity.name} ${formatPct(entity.pctChange)}${rankStr}`);

            if (showSlopeLabels) {
                const mx    = (leftX+rightX)/2;
                const my    = (y1+y2)/2;
                const arrow = entity.period2 >= entity.period1 ? "\u25B2" : "\u25BC";
                g.append("text").attr("x",mx).attr("y",my-5).attr("text-anchor","middle").attr("font-size","10px").attr("fill",color)
                    .text(`${arrow}${formatPct(entity.pctChange)}`);
            }

            // Click for cross-filter
            const rowG = g.append("rect")
                .attr("x",leftX-marginL/2).attr("y",padT).attr("width",rightX-leftX+marginL)
                .attr("height",chartH).attr("fill","transparent").style("cursor","pointer");

            rowG.on("click", (event: MouseEvent) => {
                const isMulti = event.ctrlKey || event.metaKey;
                if (self.selectedIds.has(selKey) && !isMulti) {
                    self.selectedIds.clear(); self.selMgr.clear();
                } else {
                    if (!isMulti) self.selectedIds.clear();
                    self.selectedIds.add(selKey);
                    const ids = self.entities.filter(en => {
                        const k = ((en.selectionId as unknown as Record<string,unknown>)["key"] as string) || en.name;
                        return self.selectedIds.has(k);
                    }).map(en => en.selectionId);
                    self.selMgr.select(ids, isMulti);
                }
                // Visual opacity feedback
                g.selectAll<SVGLineElement, unknown>("line[stroke]").style("opacity", (_d,_i2,nodes) => {
                    const idx2 = Array.from(nodes).indexOf(nodes[_i2]);
                    if (self.selectedIds.size === 0) return null;
                    const ent = self.entities[idx2];
                    if (!ent) return null;
                    const k = ((ent.selectionId as unknown as Record<string,unknown>)["key"] as string) || ent.name;
                    return self.selectedIds.has(k) ? "1" : "0.3";
                });
                event.stopPropagation();
            });
            rowG.on("contextmenu", (event: MouseEvent) => {
                event.preventDefault();
                event.stopPropagation();
                self.selMgr.showContextMenu(
                    entity.selectionId,
                    { x: event.clientX, y: event.clientY }
                );
            });

            void rowG; void selKey;
        });

        if (showSummary) {
            const improved = entities.filter(e => e.period2 > e.period1).length;
            const sumG     = svg.append("g").attr("transform",`translate(0,${height-padB+8})`);
            sumG.append("text").attr("x",width/2).attr("y",14).attr("text-anchor","middle").attr("font-size","11px").attr("fill","#374151")
                .text(`${improved} of ${entities.length} entities improved`);
            const barW = Math.min(200,width*0.4);
            const barX = (width-barW)/2;
            sumG.append("rect").attr("x",barX).attr("y",20).attr("width",barW).attr("height",6).attr("fill","#EF4444").attr("rx",3);
            sumG.append("rect").attr("x",barX).attr("y",20).attr("width",(improved/entities.length)*barW).attr("height",6).attr("fill","#10B981").attr("rx",3);
        }

        svg.on("click", () => {
            this.selectedIds.clear();
            this.selMgr.clear();
        });
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
