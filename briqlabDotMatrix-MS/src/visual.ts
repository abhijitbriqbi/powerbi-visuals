"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import * as d3 from "d3";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;

import { VisualFormattingSettingsModel } from "./settings";
import { checkMicrosoftLicence, resetLicenceCache } from "./licenceManager";
import { getTrialDaysRemaining, isTrialExpired, getPurchaseUrl, getButtonText } from "./trialManager";

const TRIAL_MS      = 4 * 24 * 60 * 60 * 1000;
const TRIAL_KEY     = "briqlab_trial_DotMatrix_start";
const PRO_STORE_KEY = "briqlab_dotmatrix_prokey";

const CATEGORY_COLORS = ["#0D9488","#F97316","#3B82F6","#8B5CF6","#10B981","#EF4444","#F59E0B","#EC4899","#06B6D4","#84CC16"];

function getTrial(): { daysLeft: number; expired: boolean } {
    try {
        let raw = localStorage.getItem(TRIAL_KEY);
        if (!raw) { raw = String(Date.now()); localStorage.setItem(TRIAL_KEY, raw); }
        const elapsed  = Date.now() - parseInt(raw, 10);
        const daysLeft = Math.max(0, Math.ceil((TRIAL_MS - elapsed) / 86400000));
        return { daysLeft, expired: elapsed > TRIAL_MS };
    } catch { return { daysLeft: 4, expired: false }; }
}

export class Visual implements IVisual {
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private selMgr!: powerbi.extensibility.ISelectionManager;
    private _handlersAttached = false;
    private readonly host:    IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private readonly fmtSvc:  FormattingSettingsService;

    private readonly root:       HTMLElement;
    private readonly contentEl:  HTMLDivElement;
    private readonly chartEl:    HTMLDivElement;
    private readonly trialBadge: HTMLDivElement;
    private readonly proBadge:   HTMLDivElement;
    private readonly keyErrorEl: HTMLDivElement;
    private readonly overlayEl:  HTMLDivElement;

    private settings!: VisualFormattingSettingsModel;
    private vp:        powerbi.IViewport = { width: 300, height: 300 };

    private isPro   = false;
    private lastKey = "";
    private readonly keyCache: Map<string, boolean> = new Map();

    constructor(options: VisualConstructorOptions) {
        this.host   = options.host;
        this.renderingManager = options.host.eventService;
        this.selMgr       = this.host.createSelectionManager();
        this.tooltipSvc   = this.host.tooltipService;
        this.fmtSvc = new FormattingSettingsService();

        this.root = options.element;
        this.root.classList.add("briqlab-dotmatrix");

        this.contentEl = this.mkDiv("briqlab-visual-content");
        this.root.appendChild(this.contentEl);

        this.chartEl = this.mkDiv("briq-dotmatrix-root");
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
                        dataItems: [{ displayName: "Briqlab Dot Matrix", value: "" }],
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
            if (!dv?.categorical?.values || dv.categorical.values.length < 2) {
                this.renderEmpty("Add Total and Achieved fields"); return;
            }
    
            const catData = dv.categorical!;
            let totalValue = 0, achievedValue = 0;
            for (const col of catData.values!) {
                const roles = col.source.roles ?? {};
                if (roles["total"])    totalValue    = Number(col.values[0] ?? 0);
                else if (roles["achieved"]) achievedValue = Number(col.values[0] ?? 0);
            }
            if (totalValue <= 0) { this.renderEmpty("Add Total and Achieved fields"); return; }
    
            const categories   = catData.categories;
            const hasCatBreakdown = !!(categories && categories.length > 0);
            const catNames: string[]  = [];
            const catAchieved: number[] = [];
    
            if (hasCatBreakdown) {
                const catValues = categories![0].values;
                let achievedCol: powerbi.DataViewValueColumn | undefined;
                for (const col of catData.values!) { if ((col.source.roles ?? {})["achieved"]) { achievedCol = col; break; } }
                if (achievedCol) {
                    for (let i = 0; i < catValues.length; i++) {
                        catNames.push(String(catValues[i] ?? ""));
                        catAchieved.push(Number(achievedCol.values[i] ?? 0));
                    }
                }
            }
    
            this.renderChart(totalValue, achievedValue, catNames, catAchieved, hasCatBreakdown);
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    private renderChart(totalValue: number, achievedValue: number, catNames: string[], catAchieved: number[], hasCatBreakdown: boolean): void {
        while (this.chartEl.firstChild) this.chartEl.removeChild(this.chartEl.firstChild);

        const width  = this.vp.width;
        const height = this.vp.height;
        const gs     = this.settings.gridSettings;
        const ls     = this.settings.legendSettings;

        this.chartEl.style.width    = `${width}px`;
        this.chartEl.style.height   = `${height}px`;
        this.chartEl.style.position = "relative";

        const maxDots       = Math.max(10, gs.maxDots.value ?? 100);
        const dotGapRatio   = gs.dotGap.value ?? 3;
        const showCenterTxt = gs.showCenterText.value ?? true;
        const achievedColor = gs.achievedColor.value?.value ?? "#0D9488";
        const emptyColor    = gs.emptyColor.value?.value    ?? "#E2E8F0";
        const dotShape      = String(gs.dotShape.value?.value ?? "Circle");
        const fontFamily    = String(gs.fontFamily.value?.value ?? "Segoe UI");

        const normFactor = totalValue > maxDots ? totalValue / maxDots : 1;
        const totalDots  = Math.min(Math.round(totalValue / normFactor), maxDots);
        const aDots      = Math.round(achievedValue / normFactor);
        const cols       = Math.ceil(Math.sqrt(totalDots));
        const rows       = Math.ceil(totalDots / cols);

        const padding     = 24;
        const legendHeight= 40;
        const availW      = width - padding*2;
        const availH      = height - padding*2 - legendHeight;
        const dotSizeW    = availW / (cols + (cols-1)*dotGapRatio/10);
        const dotSizeH    = availH / (rows + (rows-1)*dotGapRatio/10);
        const dotSize     = Math.min(dotSizeW, dotSizeH, 20);
        const gap         = dotSize * dotGapRatio / 10;
        const gridW       = cols*dotSize + (cols-1)*gap;
        const gridH       = rows*dotSize + (rows-1)*gap;
        const svgW        = width;
        const svgH        = height - legendHeight;
        const offsetX     = (svgW - gridW) / 2;
        const offsetY     = (svgH - gridH) / 2;

        const styleId = "briq-dotmatrix-keyframes";
        if (!document.getElementById(styleId)) {
            const s = document.createElement("style"); s.id = styleId;
            s.textContent = "@keyframes fadeInScale { 0% { opacity:0; transform:scale(0.3); } 100% { opacity:1; transform:scale(1); } }";
            document.head.appendChild(s);
        }

        const catColorMap = new Map<string, string>();
        if (hasCatBreakdown) catNames.forEach((n,i) => catColorMap.set(n, CATEGORY_COLORS[i % CATEGORY_COLORS.length]));

        const dotCats: string[] = new Array(aDots).fill("");
        if (hasCatBreakdown && catAchieved.length > 0) {
            let di = 0;
            for (let ci = 0; ci < catNames.length && di < aDots; ci++) {
                const cDots = Math.round(catAchieved[ci] / normFactor);
                for (let d = 0; d < cDots && di < aDots; d++, di++) dotCats[di] = catNames[ci];
            }
        }

        const visual = this.mkDiv("briq-visual-content");
        visual.style.cssText = "position:absolute;top:0;left:0;";
        visual.style.width   = `${width}px`;
        visual.style.height  = `${height}px`;
        this.chartEl.appendChild(visual);

        const svg = d3.select(visual).append("svg").attr("width",svgW).attr("height",svgH);
        const dotGroup = svg.append("g").attr("class","dot-grid");

        for (let i = 0; i < totalDots; i++) {
            const col = i % cols;
            const row = Math.floor(i / cols);
            const cx  = offsetX + col*(dotSize+gap) + dotSize/2;
            const cy  = offsetY + row*(dotSize+gap) + dotSize/2;
            const isFilled = i < aDots;
            let fillColor = isFilled ? achievedColor : emptyColor;
            if (isFilled && hasCatBreakdown && dotCats[i]) fillColor = catColorMap.get(dotCats[i]) ?? achievedColor;
            const animDelay = `${i*(1200/totalDots)}ms`;

            if (dotShape === "Square") {
                const r = dotSize/2;
                dotGroup.append("rect")
                    .attr("x",cx-r).attr("y",cy-r).attr("width",dotSize).attr("height",dotSize)
                    .attr("rx",2).attr("fill",fillColor)
                    .style("animation",`fadeInScale 0.4s ease-out ${animDelay} both`);
            } else {
                dotGroup.append("circle")
                    .attr("cx",cx).attr("cy",cy).attr("r",dotSize/2)
                    .attr("fill",fillColor)
                    .style("animation",`fadeInScale 0.4s ease-out ${animDelay} both`);
            }
        }

        if (showCenterTxt && totalValue > 0) {
            const pct  = Math.round((achievedValue/totalValue)*100);
            const txtG = svg.append("g").attr("class","center-text").style("pointer-events","none");
            txtG.append("text")
                .attr("x",svgW/2).attr("y",svgH/2-6)
                .attr("text-anchor","middle").attr("dominant-baseline","middle")
                .attr("font-family",fontFamily).attr("font-size",Math.min(dotSize*2.5,36))
                .attr("font-weight","bold").attr("fill",achievedColor).attr("opacity","0.85")
                .text(`${pct}%`);
            txtG.append("text")
                .attr("x",svgW/2).attr("y",svgH/2+dotSize*2)
                .attr("text-anchor","middle").attr("dominant-baseline","middle")
                .attr("font-family",fontFamily).attr("font-size",Math.min(dotSize*1.2,16))
                .attr("fill","#64748B").attr("opacity","0.75")
                .text(`${achievedValue.toLocaleString()} / ${totalValue.toLocaleString()}`);
        }

        const legendDiv = this.mkDiv("briq-legend");
        legendDiv.style.cssText = "position:absolute;bottom:4px;left:0;width:100%;display:flex;flex-wrap:wrap;justify-content:center;gap:8px;";
        legendDiv.style.fontFamily = fontFamily; legendDiv.style.fontSize = "11px"; legendDiv.style.color = "#64748B";
        visual.appendChild(legendDiv);

        if (ls.showUnitLegend.value && normFactor > 1) {
            const unitSpan = document.createElement("span");
            unitSpan.textContent = `\u25CF = ${Math.round(normFactor).toLocaleString()} units`;
            legendDiv.appendChild(unitSpan);
        }
        if (ls.showCategoryLegend.value && hasCatBreakdown) {
            catNames.forEach((name,i) => {
                const item = document.createElement("span"); item.style.cssText = "display:inline-flex;align-items:center;gap:3px;";
                const dot  = document.createElement("span"); dot.style.cssText = `display:inline-block;width:8px;height:8px;border-radius:50%;background-color:${CATEGORY_COLORS[i % CATEGORY_COLORS.length]};`;
                item.appendChild(dot);
                const lbl = document.createElement("span"); lbl.textContent = name; item.appendChild(lbl);
                legendDiv.appendChild(item);
            });
        }
    }

    private renderEmpty(msg: string): void {
        while (this.chartEl.firstChild) this.chartEl.removeChild(this.chartEl.firstChild);
        const el = this.mkDiv("briq-empty-state"); el.textContent = msg; this.chartEl.appendChild(el);
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
