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

const TRIAL_MS   = 4 * 24 * 60 * 60 * 1000;
const TRIAL_KEY  = "briqlab_trial_BriqlabProgressRing_start";
const CACHED_KEY = "briqlab_progressring_prokey";

interface RingDatum {
    category:    string;
    actual:      number;
    target:      number;
    pct:         number;
}

function achievementColor(pct: number): string {
    if (pct >= 100) return "#10B981";
    if (pct >= 80)  return "#0D9488";
    if (pct >= 60)  return "#F59E0B";
    return "#EF4444";
}

export class Visual implements IVisual {
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private selMgr!:      powerbi.extensibility.ISelectionManager;
    private _handlersAttached = false;
    private readonly host: IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private readonly fmtSvc: FormattingSettingsService;
    private settings!: VisualFormattingSettingsModel;

    private root!:       HTMLElement;
    private contentEl!:  HTMLElement;
    private svgEl!:      d3.Selection<SVGSVGElement, unknown, null, undefined>;
    private trialBadge!: HTMLElement;
    private proBadge!:   HTMLElement;
    private overlay!:    HTMLElement;

    private vp: powerbi.IViewport = { width: 400, height: 300 };
    private curKey = "";
    private isPro  = false;
    private keyCache: Map<string, boolean> = new Map();
    private trialStart = 0;
    private rings: RingDatum[] = [];

    // Center display elements (update on hover)
    private centerPctText!:   d3.Selection<SVGTextElement, unknown, null, undefined>;
    private centerNameText!:  d3.Selection<SVGTextElement, unknown, null, undefined>;
    private centerMetText!:   d3.Selection<SVGTextElement, unknown, null, undefined>;

    constructor(options: VisualConstructorOptions) {
        this.host   = options.host;
        this.renderingManager = options.host.eventService;
        this.selMgr     = this.host.createSelectionManager();
        this.tooltipSvc = this.host.tooltipService;
        this.fmtSvc = new FormattingSettingsService();
        this.root   = options.element;
        this.root.classList.add("briqlab-ring");
        this.buildDOM();
        this.initTrial();
    }

    private buildDOM(): void {
        this.contentEl = document.createElement("div");
        this.contentEl.className = "ring-content";
        this.root.appendChild(this.contentEl);

        this.svgEl = d3.select(this.contentEl)
            .append<SVGSVGElement>("svg")
            .attr("class", "ring-svg");

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
                        dataItems: [{ displayName: "Briqlab Progress Ring", value: "" }],
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
    
            this.rings = [];
            const dv = options.dataViews?.[0];
            if (dv?.categorical) {
                const cats = dv.categorical.categories ?? [];
                const vals = dv.categorical.values    ?? [];
                const catCol    = cats.find(c => (c.source.roles as Record<string, unknown>)["category"]);
                const actualCol = vals.find(c => (c.source.roles as Record<string, unknown>)["actualValue"]);
                const targetCol = vals.find(c => (c.source.roles as Record<string, unknown>)["targetValue"]);
    
                if (catCol) {
                    const maxR = Math.min(6, Math.max(1, this.settings.ringSettings.maxRings.value ?? 6));
                    const n    = Math.min(catCol.values.length, maxR);
                    for (let i = 0; i < n; i++) {
                        const actual = actualCol ? Number(actualCol.values[i]) : 0;
                        const target = targetCol ? Number(targetCol.values[i]) : 0;
                        const safeTgt = isNaN(target) || target === 0 ? 1 : target;
                        const pct = Math.min(100, Math.max(0, (isNaN(actual) ? 0 : actual) / safeTgt * 100));
                        this.rings.push({
                            category: String(catCol.values[i] ?? ""),
                            actual:   isNaN(actual) ? 0 : actual,
                            target:   isNaN(target) ? 0 : target,
                            pct
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

        const rs = this.settings.ringSettings;
        const ls = this.settings.labelSettings;
        const cs = this.settings.centerSettings;

        const { width, height } = this.vp;
        const trackWidth  = Math.min(24, Math.max(8,  rs.trackWidth.value  ?? 18));
        const ringGap     = Math.min(12, Math.max(2,  rs.ringGap.value     ?? 6));
        const autoColor   = rs.autoColor.value;
        const fontFamily  = String(rs.fontFamily?.value?.value ?? "Segoe UI");
        const showLabels  = ls.showLabels.value;
        const labelFS     = Math.min(13, Math.max(9, ls.labelFontSize.value ?? 11));
        const showCenter  = cs.showCenter.value;
        const summaryMetric = String(cs.summaryMetric?.value?.value ?? "Average %");

        this.svgEl.attr("width", width).attr("height", height);
        this.svgEl.selectAll("*").remove();

        if (this.rings.length === 0) {
            this.svgEl.append("text")
                .attr("x", width / 2).attr("y", height / 2)
                .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                .attr("class", "ring-empty").attr("font-family", fontFamily)
                .text("Add Category, Actual & Target fields");
            return;
        }

        const n          = this.rings.length;
        const labelAreaW = showLabels ? 140 : 0;
        const ringAreaW  = width - labelAreaW;
        const outerR     = Math.min(ringAreaW / 2, height / 2) - 8;
        const innerR     = outerR - n * (trackWidth + ringGap) + ringGap;
        const cx         = ringAreaW / 2;
        const cy         = height / 2;

        const avgPct  = this.rings.reduce((a, r) => a + r.pct, 0) / n;
        const metCnt  = this.rings.filter(r => r.pct >= 100).length;

        // Default center text
        const defaultCenterPct   = summaryMetric === "Count Met" ? `${metCnt}/${n}` : `${avgPct.toFixed(0)}%`;
        const defaultCenterName  = summaryMetric === "Count Met" ? "Targets Met" : "Overall";
        const defaultCenterMet   = summaryMetric === "Count Met" ? "" : `${metCnt} of ${n} met`;

        const updateCenter = (pctStr: string, name: string, met: string) => {
            if (!showCenter) return;
            if (this.centerPctText)  this.centerPctText.text(pctStr);
            if (this.centerNameText) this.centerNameText.text(name);
            if (this.centerMetText)  this.centerMetText.text(met);
        };

        const g = this.svgEl.append("g").attr("class", "ring-g");

        // Center background circle
        if (showCenter && innerR > 10) {
            const centerR = Math.max(0, innerR - ringGap - 2);
            g.append("circle")
                .attr("cx", cx).attr("cy", cy).attr("r", centerR)
                .attr("fill", "#fff")
                .attr("filter", "drop-shadow(0 1px 3px rgba(0,0,0,0.08))");

            // Center text
            this.centerPctText = g.append("text")
                .attr("x", cx).attr("y", cy - 10)
                .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                .attr("class", "ring-center-pct").attr("font-family", fontFamily)
                .text(defaultCenterPct);

            this.centerNameText = g.append("text")
                .attr("x", cx).attr("y", cy + 8)
                .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                .attr("class", "ring-center-name").attr("font-family", fontFamily)
                .text(defaultCenterName);

            this.centerMetText = g.append("text")
                .attr("x", cx).attr("y", cy + 22)
                .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                .attr("class", "ring-center-met").attr("font-family", fontFamily)
                .text(defaultCenterMet);
        }

        // Draw rings (outermost = index 0)
        this.rings.forEach((ring, i) => {
            const ro = outerR - i * (trackWidth + ringGap);
            const ri = ro - trackWidth;
            const color = autoColor ? achievementColor(ring.pct) : "#0D9488";

            const arcFn = d3.arc<number>()
                .innerRadius(ri)
                .outerRadius(ro)
                .startAngle(-Math.PI / 2)
                .endAngle((d: number) => -Math.PI / 2 + (d / 100) * 2 * Math.PI);

            const trackFn = d3.arc<number>()
                .innerRadius(ri)
                .outerRadius(ro)
                .startAngle(-Math.PI / 2)
                .endAngle(3 * Math.PI / 2);

            const ringG = g.append("g")
                .attr("transform", `translate(${cx},${cy})`)
                .attr("class", "ring-group")
                .style("cursor", "pointer");

            // Track background
            ringG.append("path")
                .datum(100)
                .attr("d", trackFn)
                .attr("fill", "rgba(0,0,0,0.06)")
                .attr("class", "ring-track");

            // Progress arc
            ringG.append("path")
                .datum(ring.pct)
                .attr("d", arcFn)
                .attr("fill", color)
                .attr("stroke", color)
                .attr("stroke-width", 0)
                .attr("stroke-linecap", "round")
                .attr("class", "ring-arc");

            // Hover
            ringG.on("mouseover", () => {
                updateCenter(`${ring.pct.toFixed(0)}%`, ring.category, `${ring.actual} / ${ring.target}`);
            }).on("mouseout", () => {
                updateCenter(defaultCenterPct, defaultCenterName, defaultCenterMet);
            });

            // tooltip
            ringG.append("title")
                .text(`${ring.category}: ${ring.pct.toFixed(1)}% (${ring.actual} / ${ring.target})`);

            // Right-side labels
            if (showLabels) {
                const labelX = ringAreaW + 8;
                const labelY = cy - (n / 2 - i - 0.5) * (trackWidth + ringGap);
                const rowG = this.svgEl.append("g").attr("class", "ring-label-group");

                // Color dot
                rowG.append("circle")
                    .attr("cx", labelX + 5).attr("cy", labelY)
                    .attr("r", 4).attr("fill", color);

                rowG.append("text")
                    .attr("x", labelX + 14).attr("y", labelY - 5)
                    .attr("dominant-baseline", "middle")
                    .attr("class", "ring-label-name").attr("font-size", labelFS)
                    .attr("font-family", fontFamily).attr("fill", "#0F172A")
                    .text(ring.category);

                rowG.append("text")
                    .attr("x", labelX + 14).attr("y", labelY + 8)
                    .attr("dominant-baseline", "middle")
                    .attr("class", "ring-label-pct").attr("font-size", labelFS - 1)
                    .attr("font-family", fontFamily).attr("fill", color)
                    .text(`${ring.pct.toFixed(0)}%  ${ring.actual}/${ring.target}`);
            }
        });
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
