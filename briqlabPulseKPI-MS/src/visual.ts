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

const TRIAL_MS  = 4 * 24 * 60 * 60 * 1000;
const TRIAL_KEY = "briqlab_trial_BriqlabPulseKPI_start";
const PRO_KEY   = "briqlab_pulsekpi_prokey";

function clamp(v: number, lo: number, hi: number) { return Math.max(lo, Math.min(hi, v)); }

function fmtValue(v: number, prefix: string, format: string): string {
    const p = prefix === "None" ? "" : prefix;
    const abs = Math.abs(v);
    if (format === "exact") return `${p}${v.toLocaleString()}`;
    if (format === "K" || (format === "auto" && abs >= 1000 && abs < 1e6))
        return `${p}${(v/1000).toFixed(1).replace(/\.0$/,"")}K`;
    if (format === "M" || (format === "auto" && abs >= 1e6 && abs < 1e9))
        return `${p}${(v/1e6).toFixed(1).replace(/\.0$/,"")}M`;
    if (format === "B" || (format === "auto" && abs >= 1e9))
        return `${p}${(v/1e9).toFixed(1).replace(/\.0$/,"")}B`;
    return `${p}${v.toLocaleString()}`;
}

function getInsight(value: number, trendVals: number[], targetVal: number | null): string {
    if (trendVals.length === 0) {
        if (targetVal !== null) {
            const pct = value / targetVal;
            if (pct >= 1.1) return `${((pct-1)*100).toFixed(1)}% above target — excellent`;
            if (pct < 0.9)  return `${((1-pct)*100).toFixed(1)}% below target — needs attention`;
        }
        return "Tracking within normal range";
    }
    const allVals = trendVals.filter(v => v != null && !isNaN(v));
    if (allVals.length === 0) return "Tracking within normal range";
    const maxV = Math.max(...allVals), minV = Math.min(...allVals);
    if (value >= maxV) return "Highest value in the series";
    if (value <= minV) return "Lowest value in the series";
    const n = allVals.length;
    if (n >= 3) {
        const last3 = allVals.slice(-3);
        if (last3[2] < last3[1] && last3[1] < last3[0]) return "3 consecutive periods of decline";
        if (last3[2] > last3[1] && last3[1] > last3[0]) return "3 consecutive periods of growth";
    }
    if (targetVal !== null && targetVal !== 0) {
        const pct = value / targetVal;
        if (pct >= 1.1) return `${((pct-1)*100).toFixed(1)}% above target — excellent`;
        if (pct < 0.9)  return `${((1-pct)*100).toFixed(1)}% below target — needs attention`;
    }
    return "Tracking within normal range";
}

function getVelocity(trendVals: number[]): { text: string; color: string } {
    const vals = trendVals.filter(v => v != null && !isNaN(v));
    if (vals.length < 4) return { text: "→ stable", color: "#94A3B8" };
    const half = Math.floor(vals.length / 2);
    const first = vals.slice(0, half).reduce((a,b)=>a+b,0)/half;
    const second = vals.slice(-half).reduce((a,b)=>a+b,0)/half;
    const ratio = first === 0 ? 1 : second / first;
    if (ratio > 1.05) return { text: "▲▲ accelerating", color: "#10B981" };
    if (ratio < 0.95) return { text: "▽▽ decelerating", color: "#F59E0B" };
    return { text: "→ stable", color: "#94A3B8" };
}

export class Visual implements IVisual {
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private selMgr!:      powerbi.extensibility.ISelectionManager;
    private _handlersAttached = false;
    private readonly host:    IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private readonly fmtSvc: FormattingSettingsService;
    private settings!: VisualFormattingSettingsModel;

    private root!: HTMLElement;
    private contentEl!: HTMLElement;
    private trialBadge!: HTMLElement;
    private proBadge!: HTMLElement;
    private keyError!: HTMLElement;
    private overlay!: HTMLElement;

    private vp: powerbi.IViewport = { width: 300, height: 200 };
    private curKey   = "";
    private isPro    = false;
    private keyCache: Map<string, boolean> = new Map();
    private trialStart = 0;

    private value: number | null = null;
    private targetVal: number | null = null;
    private comparison: number | null = null;
    private trendValues: number[] = [];
    private trendDates: string[] = [];

    constructor(options: VisualConstructorOptions) {
        this.host   = options.host;
        this.renderingManager = options.host.eventService;
        this.selMgr     = this.host.createSelectionManager();
        this.tooltipSvc = this.host.tooltipService;
        this.fmtSvc = new FormattingSettingsService();
        this.root   = options.element;
        this.root.classList.add("briqlab-pulsekpi");
        this.buildDOM();
        this.initTrial();
    }

    private buildDOM(): void {
        this.contentEl = document.createElement("div");
        this.contentEl.className = "pkpi-content";
        this.root.appendChild(this.contentEl);

        this.trialBadge = document.createElement("div");
        this.trialBadge.className = "briq-trial-badge hidden";
        this.root.appendChild(this.trialBadge);

        this.proBadge = document.createElement("div");
        this.proBadge.className = "briq-pro-badge hidden";
        this.proBadge.textContent = "✓ Pro Active";
        this.root.appendChild(this.proBadge);

        this.keyError = document.createElement("div");
        this.keyError.className = "briq-key-error hidden";
        this.keyError.textContent = "✗ Invalid key";
        this.root.appendChild(this.keyError);

        this.overlay = document.createElement("div");
        this.overlay.className = "briq-trial-overlay hidden";
        this.root.appendChild(this.overlay);
        this.buildOverlay();
    }

    private buildOverlay(): void {
        const card = document.createElement("div");
        card.className = "briq-trial-card";

        const lock = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        lock.setAttribute("viewBox","0 0 24 24"); lock.setAttribute("width","40"); lock.setAttribute("height","40");
        const lp = document.createElementNS("http://www.w3.org/2000/svg","path");
        lp.setAttribute("d","M12 2a5 5 0 0 1 5 5v3h1a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2v-8a2 2 0 0 1 2-2h1V7a5 5 0 0 1 5-5zm0 2a3 3 0 0 0-3 3v3h6V7a3 3 0 0 0-3-3z");
        lp.setAttribute("fill","#F97316");
        lock.appendChild(lp);
        card.appendChild(lock);

        const t = document.createElement("p"); t.className = "trial-title"; t.textContent = "Free trial ended"; card.appendChild(t);
        const b = document.createElement("p"); b.className = "trial-body";  b.textContent = "Activate Briqlab Pro to continue using this visual and unlock all features."; card.appendChild(b);
        const btn = document.createElement("button"); btn.className = "trial-btn"; btn.textContent = getButtonText();
        btn.addEventListener("click", () => this.host.launchUrl(getPurchaseUrl()));
        card.appendChild(btn);
        const sub = document.createElement("p"); sub.className = "trial-subtext"; sub.textContent = "Purchase on Microsoft AppSource to unlock all features instantly."; card.appendChild(sub);
        this.overlay.appendChild(card);
    }

    private initTrial(): void {
        try {
            const stored = localStorage.getItem(TRIAL_KEY);
            this.trialStart = stored ? parseInt(stored, 10) : Date.now();
            if (!stored) localStorage.setItem(TRIAL_KEY, String(this.trialStart));
            const savedKey = localStorage.getItem(PRO_KEY);
            if (savedKey) { this.curKey = savedKey; this.validateKey(savedKey).then(ok => { this.isPro = ok; this.render(); }); }
        } catch { this.trialStart = Date.now(); }
    }

    private isExpired() { return Date.now() - this.trialStart >= TRIAL_MS; }
    private daysLeft()  { return Math.max(0, 4 - Math.floor((Date.now()-this.trialStart)/(24*60*60*1000))); }

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
                        dataItems: [{ displayName: "Briqlab Pulse KPI", value: "" }],
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
            this.settings = this.fmtSvc.populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews?.[0]);
    
            this.value = null; this.targetVal = null; this.comparison = null;
            this.trendValues = []; this.trendDates = [];
    
            const dv = options.dataViews?.[0];
            if (dv?.categorical) {
                const cats = dv.categorical.categories;
                const vals = dv.categorical.values;
                if (cats?.length) {
                    this.trendDates = (cats[0].values as string[]).map(v => String(v ?? ""));
                }
                if (vals?.length) {
                    for (const col of vals) {
                        const roles = col.source.roles as Record<string, unknown> ?? {};
                        const nums  = col.values.map(v => typeof v === "number" ? v : NaN);
                        if (roles["measure"])    this.value      = nums.find(n => !isNaN(n)) ?? null;
                        if (roles["target"])     this.targetVal  = nums.find(n => !isNaN(n)) ?? null;
                        if (roles["comparison"]) this.comparison = nums.find(n => !isNaN(n)) ?? null;
                        if (roles["trendValue"]) this.trendValues = nums.filter(n => !isNaN(n));
                    }
                }
            }
    
            const key = ""; // MS cert: pro key field removed
            if (key && key !== this.curKey) {
                this.curKey = key; this.isPro = false;
                this.validateKey(key).then(ok => { this.isPro = ok; this.render(); });
            } else if (!key && !this.curKey) { this.isPro = false; }
    
            this.render();
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    private render(): void {
        if (!this.settings) return;
        const s = this.settings;
        const { width, height } = this.vp;

        const expired = this.isExpired();
        if (!this.isPro && !expired) {
            this.trialBadge.textContent = `Trial: ${this.daysLeft()} days remaining`;
            this.trialBadge.classList.remove("hidden");
        } else { this.trialBadge.classList.add("hidden"); }

        if (this.isPro) { this.proBadge.classList.remove("hidden"); this.keyError.classList.add("hidden"); }
        else { this.proBadge.classList.add("hidden"); }

        if (expired && !this.isPro) {
            this.contentEl.classList.add("blurred");
            this.overlay.classList.remove("hidden");
        } else {
            this.contentEl.classList.remove("blurred");
            this.overlay.classList.add("hidden");
        }

        const bgColor    = s.colorSettings.backgroundColor.value?.value ?? "#FFFFFF";
        const primary    = s.colorSettings.primaryColor.value?.value ?? "#0D9488";
        const aboveColor = s.colorSettings.aboveTargetColor.value?.value ?? "#10B981";
        const belowColor = s.colorSettings.belowTargetColor.value?.value ?? "#EF4444";
        const fontFam    = String(s.cardSettings.fontFamily?.value?.value ?? "Segoe UI");
        const prefix     = String(s.cardSettings.currencyPrefix?.value?.value ?? "None");
        const vFormat    = String(s.cardSettings.valueFormat?.value?.value ?? "auto");
        const showRing   = s.cardSettings.showPulseRing.value;
        const showSpark  = s.cardSettings.showSparkline.value;
        const showInsight= s.cardSettings.showInsightText.value;
        const showBar    = s.cardSettings.showProgressBar.value;

        const pctOfTarget = (this.value !== null && this.targetVal !== null && this.targetVal !== 0)
            ? this.value / this.targetVal : null;

        const statusColor = pctOfTarget === null ? primary
            : pctOfTarget >= 1     ? aboveColor
            : pctOfTarget >= 0.9   ? "#F59E0B"
            : belowColor;

        this.contentEl.style.cssText =
            `width:${width}px;height:${height}px;background:${bgColor};` +
            `font-family:${fontFam};`;

        // Clear and rebuild DOM
        while (this.contentEl.firstChild) this.contentEl.removeChild(this.contentEl.firstChild);

        const card = document.createElement("div");
        card.className = "pkpi-card";
        card.style.cssText = `width:100%;height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:12px;box-sizing:border-box;gap:6px;`;

        // ── Label row ──────────────────────────────────────────────────────────
        const labelRow = document.createElement("div");
        labelRow.className = "pkpi-label-row";
        labelRow.textContent = "KPI";
        card.appendChild(labelRow);

        // ── Pulse ring + value ─────────────────────────────────────────────────
        const ringWrap = document.createElement("div");
        ringWrap.className = "pkpi-ring-wrap";

        if (showRing) {
            const svgNS = "http://www.w3.org/2000/svg";
            const svg = document.createElementNS(svgNS, "svg");
            const r = 44, cx = 50, cy = 50;
            svg.setAttribute("width", "100"); svg.setAttribute("height", "100");
            svg.setAttribute("class", "pkpi-ring-svg");

            const bgCirc = document.createElementNS(svgNS, "circle");
            bgCirc.setAttribute("cx", String(cx)); bgCirc.setAttribute("cy", String(cy));
            bgCirc.setAttribute("r", String(r));
            bgCirc.setAttribute("fill", "none");
            bgCirc.setAttribute("stroke", statusColor);
            bgCirc.setAttribute("stroke-width", "2");
            bgCirc.setAttribute("opacity", "0.15");
            svg.appendChild(bgCirc);

            const pulseCirc = document.createElementNS(svgNS, "circle");
            pulseCirc.setAttribute("cx", String(cx)); pulseCirc.setAttribute("cy", String(cy));
            pulseCirc.setAttribute("r", String(r));
            pulseCirc.setAttribute("fill", "none");
            pulseCirc.setAttribute("stroke", statusColor);
            pulseCirc.setAttribute("stroke-width", "2");
            pulseCirc.setAttribute("class", pctOfTarget !== null && pctOfTarget < 0.9 ? "pkpi-pulse-fast" : "pkpi-pulse");
            svg.appendChild(pulseCirc);
            ringWrap.appendChild(svg);
        }

        const valEl = document.createElement("div");
        valEl.className = "pkpi-value";
        valEl.style.color = statusColor;
        valEl.textContent = this.value !== null ? fmtValue(this.value, prefix, vFormat) : "—";
        ringWrap.appendChild(valEl);
        card.appendChild(ringWrap);

        // ── Change row ─────────────────────────────────────────────────────────
        if (this.comparison !== null && this.value !== null) {
            const change = this.value - this.comparison;
            const changePct = this.comparison !== 0 ? (change / Math.abs(this.comparison)) * 100 : 0;
            const up = change >= 0;
            const chRow = document.createElement("div");
            chRow.className = "pkpi-change-row";
            const arrow = document.createElement("span");
            arrow.className = up ? "pkpi-arrow-up" : "pkpi-arrow-down";
            arrow.textContent = up ? "▲" : "▼";
            arrow.style.color = up ? aboveColor : belowColor;
            const chText = document.createElement("span");
            chText.className = "pkpi-change-text";
            chText.textContent = ` ${Math.abs(changePct).toFixed(1)}%`;
            chText.style.color = up ? aboveColor : belowColor;
            chRow.appendChild(arrow); chRow.appendChild(chText);
            if (this.targetVal !== null) {
                const tgt = document.createElement("span");
                tgt.className = "pkpi-target-text";
                tgt.textContent = `  vs ${fmtValue(this.targetVal, prefix, vFormat)} target`;
                chRow.appendChild(tgt);
            }
            card.appendChild(chRow);
        }

        // ── Progress bar ───────────────────────────────────────────────────────
        if (showBar && pctOfTarget !== null) {
            const pct = clamp(pctOfTarget * 100, 0, 100);
            const barWrap = document.createElement("div");
            barWrap.className = "pkpi-bar-wrap";
            const barFill = document.createElement("div");
            barFill.className = "pkpi-bar-fill";
            barFill.style.cssText = `width:${pct}%;background:${statusColor};`;
            barWrap.appendChild(barFill);
            const barLabel = document.createElement("span");
            barLabel.className = "pkpi-bar-label";
            barLabel.textContent = `${pct.toFixed(1)}%`;
            barWrap.appendChild(barLabel);
            card.appendChild(barWrap);
        }

        // ── Sparkline ──────────────────────────────────────────────────────────
        if (showSpark && this.trendValues.length > 1) {
            const sparkH = clamp(height * 0.22, 30, 60);
            const sparkW = width - 24;
            const sparkSvg = d3.create("svg")
                .attr("width", sparkW).attr("height", sparkH)
                .attr("class", "pkpi-sparkline");

            const xScale = d3.scaleLinear().domain([0, this.trendValues.length - 1]).range([0, sparkW]);
            const ext = d3.extent(this.trendValues) as [number, number];
            const yScale = d3.scaleLinear().domain([ext[0] * 0.95, ext[1] * 1.05]).range([sparkH - 2, 2]);

            const area = d3.area<number>()
                .x((_, i) => xScale(i))
                .y0(sparkH)
                .y1(d => yScale(d))
                .curve(d3.curveCatmullRom);

            const line = d3.line<number>()
                .x((_, i) => xScale(i))
                .y(d => yScale(d))
                .curve(d3.curveCatmullRom);

            sparkSvg.append("path").datum(this.trendValues)
                .attr("d", area).attr("fill", statusColor).attr("opacity", 0.15);
            sparkSvg.append("path").datum(this.trendValues)
                .attr("d", line).attr("fill","none").attr("stroke", statusColor).attr("stroke-width", 1.5);

            const lastX = xScale(this.trendValues.length - 1);
            const lastY = yScale(this.trendValues[this.trendValues.length - 1]);
            sparkSvg.append("circle").attr("cx", lastX).attr("cy", lastY).attr("r", 4).attr("fill", statusColor);

            const sparkNode = sparkSvg.node();
            if (sparkNode) card.appendChild(sparkNode);

            // velocity
            const vel = getVelocity(this.trendValues);
            const velEl = document.createElement("div");
            velEl.className = "pkpi-velocity";
            velEl.style.color = vel.color;
            velEl.textContent = vel.text;
            card.appendChild(velEl);
        }

        // ── Insight text ───────────────────────────────────────────────────────
        if (showInsight && this.value !== null) {
            const insight = getInsight(this.value, this.trendValues, this.targetVal);
            const insightEl = document.createElement("div");
            insightEl.className = "pkpi-insight";
            insightEl.textContent = `"${insight}"`;
            card.appendChild(insightEl);
        }

        this.contentEl.appendChild(card);
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
