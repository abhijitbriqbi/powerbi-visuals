"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import * as d3 from "d3";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

import { VisualFormattingSettingsModel } from "./settings";
import { checkMicrosoftLicence, resetLicenceCache } from "./licenceManager";
import { getTrialDaysRemaining, isTrialExpired, getPurchaseUrl, getButtonText } from "./trialManager";

// ── Constants ─────────────────────────────────────────────────────────────────
const START_ANGLE = -Math.PI / 2; // 9 o'clock (left)
const END_ANGLE   =  Math.PI / 2; // 3 o'clock (right)
const ARC_BG_COLOR = "#F3F4F6";
const TRIAL_MS = 4 * 24 * 60 * 60 * 1000;
const TRIAL_KEY = "briqlab_trial_gauge_start";
const PRO_STORAGE_KEY = "briqlab_gauge_prokey";

// ── Helpers ───────────────────────────────────────────────────────────────────
function clamp(v: number, lo: number, hi: number): number {
    return Math.max(lo, Math.min(hi, v));
}

function formatNumber(value: number): string {
    const abs = Math.abs(value);
    if (abs >= 1_000_000) {
        const r = (value / 1_000_000).toFixed(1);
        return (r.endsWith(".0") ? r.slice(0, -2) : r) + "M";
    }
    if (abs >= 1_000) {
        const r = (value / 1_000).toFixed(1);
        return (r.endsWith(".0") ? r.slice(0, -2) : r) + "K";
    }
    return value.toLocaleString("en-US");
}

/** Returns an angle (radians, D3 convention) for a value within [min, max]. */
function valueToAngle(value: number, min: number, max: number): number {
    const range = max - min;
    if (range === 0) return START_ANGLE;
    return START_ANGLE + clamp((value - min) / range, 0, 1) * Math.PI;
}

/**
 * Arc fill color:
 *   - If target is known: colour by (value / target)
 *   - Otherwise: colour by (value - min) / (max - min)
 *   < 50 % → red  |  50–80 % → amber  |  ≥ 80 % → green
 */
function autoColor(value: number, min: number, max: number, target: number | null): string {
    let pct: number;
    if (target !== null && target !== 0) {
        pct = value / target;
    } else {
        const range = max - min;
        pct = range === 0 ? 0 : (value - min) / range;
    }
    if (pct < 0.5) return "#EF4444";
    if (pct < 0.8) return "#F59E0B";
    return "#10B981";
}

// ── DOM helper (no innerHTML allowed) ─────────────────────────────────────────
function el<K extends keyof HTMLElementTagNameMap>(
    tag: K,
    classes?: string[],
    text?: string
): HTMLElementTagNameMap[K] {
    const e = document.createElement(tag);
    if (classes) classes.forEach(c => e.classList.add(c));
    if (text !== undefined) e.textContent = text;
    return e;
}

// ── Visual ────────────────────────────────────────────────────────────────────
export class Visual implements IVisual {
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private selMgr!:      powerbi.extensibility.ISelectionManager;
    private _handlersAttached = false;
    private readonly host: IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private readonly target: HTMLElement;
    private readonly fmtService: FormattingSettingsService;

    private settings!: VisualFormattingSettingsModel;

    // DOM nodes
    private root!: HTMLDivElement;
    private content!: HTMLDivElement;
    private trialBadge!: HTMLDivElement;
    private proBadge!: HTMLDivElement;
    private keyError!: HTMLDivElement;
    private overlay!: HTMLDivElement;
    private svg!: d3.Selection<SVGSVGElement, unknown, null, undefined>;

    // Extracted data
    private value: number | null = null;
    private minVal: number = 0;
    private maxVal: number = 100;
    private targetVal: number | null = null;

    // Viewport
    private viewport: powerbi.IViewport = { width: 200, height: 200 };

    // Pro key
    private currentKey: string = "";
    private isPro: boolean = false;
    private keyCache: Map<string, boolean> = new Map();

    // Trial
    private trialStart: number = 0;

    // ── Constructor ──────────────────────────────────────────────────────────
    constructor(options: VisualConstructorOptions) {
        this.fmtService = new FormattingSettingsService();
        this.host = options.host;
        this.renderingManager = options.host.eventService;
        this.selMgr     = this.host.createSelectionManager();
        this.tooltipSvc = this.host.tooltipService;
        this.target = options.element;

        this.buildDOM();
        this.initTrial();
        this.restoreProKey();
    }

    // ── DOM Build ────────────────────────────────────────────────────────────
    private buildDOM(): void {
        this.root = el("div", ["briqlab-gauge"]);
        this.target.appendChild(this.root);

        // Visual content (holds SVG, gets blurred when expired)
        this.content = el("div", ["briqlab-visual-content"]);
        this.root.appendChild(this.content);

        // SVG inside content
        this.svg = d3.select(this.content)
            .append<SVGSVGElement>("svg")
            .attr("class", "gauge-svg");

        // Trial badge (bottom-left, outside content)
        this.trialBadge = el("div", ["briqlab-trial-badge", "hidden"]);
        this.root.appendChild(this.trialBadge);

        // Pro badge (bottom-right, outside content)
        this.proBadge = el("div", ["briqlab-pro-badge", "hidden"]);
        this.root.appendChild(this.proBadge);

        // Key error (near bottom-right, outside content)
        this.keyError = el("div", ["briqlab-key-error", "hidden"]);
        this.root.appendChild(this.keyError);

        // Trial overlay (inset:0, z-index:100)
        this.overlay = el("div", ["briqlab-trial-overlay", "hidden"]);
        this.root.appendChild(this.overlay);
        this.buildOverlay();
    }

    private buildOverlay(): void {
        const card = el("div", ["briqlab-trial-card"]);

        const title = el("p", ["trial-title"], "Free trial ended");
        card.appendChild(title);

        const body = el("p", ["trial-body"], "Activate Briqlab Pro to continue using this visual and unlock all features.");
        card.appendChild(body);

        const btn = el("button", ["trial-btn"], "Get Pro at www.briqlab.io/pricing");
        btn.addEventListener("click", () => {
            this.host.launchUrl(getPurchaseUrl());
        });
        card.appendChild(btn);

        const sub = el("p", ["trial-subtext"], "Purchase on Microsoft AppSource to unlock all features instantly.");
        card.appendChild(sub);

        this.overlay.appendChild(card);
    }

    // ── Trial management ─────────────────────────────────────────────────────
    private initTrial(): void {
        try {
            const stored = localStorage.getItem(TRIAL_KEY);
            if (stored) {
                this.trialStart = parseInt(stored, 10);
            } else {
                this.trialStart = Date.now();
                localStorage.setItem(TRIAL_KEY, String(this.trialStart));
            }
        } catch {
            this.trialStart = Date.now();
        }
    }

    private getTrialDaysRemaining(): number {
        const elapsed = Date.now() - this.trialStart;
        const daysUsed = Math.floor(elapsed / (24 * 60 * 60 * 1000));
        return Math.max(0, 4 - daysUsed);
    }

    private isTrialExpired(): boolean {
        return Date.now() - this.trialStart >= TRIAL_MS;
    }

    // ── Pro key handling ─────────────────────────────────────────────────────
    private restoreProKey(): void {
        try {
            const saved = localStorage.getItem(PRO_STORAGE_KEY);
            if (saved) {
                this.currentKey = saved;
                this.validateKey(saved).then((ok) => {
                    this.isPro = ok;
                    this.updateTrialUI();
                    this.renderGauge();
                });
            }
        } catch {
            // ignore
        }
    }

    private async validateKey(key: string): Promise<boolean> {
        return checkMicrosoftLicence(this.host);
    }

    // ── Trial / Pro UI update ────────────────────────────────────────────────
    private updateTrialUI(): void {
        checkMicrosoftLicence(this.host).then(p => this._msUpdateLicenceUI(p)).catch(() => this._msUpdateLicenceUI(false));
    }

    // ── update ───────────────────────────────────────────────────────────────
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
                        dataItems: [{ displayName: "Briqlab Gauge", value: "" }],
                        identities: [],
                        coordinates: [e.clientX, e.clientY],
                        isTouchEvent: false
                    });
                });
                this.root.addEventListener("mouseleave", () => {
                    this.tooltipSvc.hide({ isTouchEvent: false, immediately: false });
                });
            }
            this.viewport = options.viewport;
    
            this.settings = this.fmtService.populateFormattingSettingsModel(
                VisualFormattingSettingsModel,
                options.dataViews?.[0]
            );
    
            // Extract measures from categorical dataView (no categories, values only)
            this.value     = null;
            this.minVal    = 0;
            this.maxVal    = 100;
            this.targetVal = null;
    
            const dv = options.dataViews?.[0];
            if (dv?.categorical?.values) {
                for (const col of dv.categorical.values) {
                    const raw = col.values?.[0];
                    const num = typeof raw === "number" ? raw : null;
                    const roles = col.source.roles as Record<string, unknown> ?? {};
                    if (roles["measure"]  && num !== null) this.value     = num;
                    if (roles["minValue"] && num !== null) this.minVal    = num;
                    if (roles["maxValue"] && num !== null) this.maxVal    = num;
                    if (roles["target"]   && num !== null) this.targetVal = num;
                }
            }
    
            // Pro key check
            const key = ""; // MS cert: pro key field removed
            if (key !== this.currentKey) {
                this.currentKey = key;
                this.isPro = false;
                this.keyError.classList.add("hidden");
    
                if (key) {
                    try {
                        localStorage.setItem(PRO_STORAGE_KEY, key);
                    } catch {
                        // ignore
                    }
                    this.validateKey(key).then((ok) => {
                        this.isPro = ok;
                        if (!ok) {
                            this.keyError.classList.remove("hidden");
                            this.keyError.textContent = "\u2717 Invalid key";
                        } else {
                            this.keyError.classList.add("hidden");
                        }
                        this.updateTrialUI();
                        this.renderGauge();
                    });
                }
            }
    
            this.updateTrialUI();
            this.renderGauge();
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    // ── renderGauge ──────────────────────────────────────────────────────────
    private renderGauge(): void {
        if (!this.settings) return;

        const { width, height } = this.viewport;
        const s = this.settings;

        // Gauge settings
        const gaugeType     = String(s.gaugeSettings.gaugeType?.value?.value ?? "semi");
        const arcThickPct   = clamp(s.gaugeSettings.arcThickness?.value ?? 18, 5, 40) / 100;
        const valueFontSize = clamp(s.gaugeSettings.valueFontSize.value, 8, 60);
        const showTarget    = s.gaugeSettings.showTarget.value;
        const targetColor   = s.gaugeSettings.targetColor?.value?.value ?? "#F97316";
        const isManual      = s.gaugeSettings.manualColorMode.value;
        const manualColor   = s.gaugeSettings.manualColor.value?.value ?? "#0D9488";
        const showNeedle    = s.gaugeSettings.showNeedle?.value ?? true;
        const showTicks     = s.gaugeSettings.showTicks?.value ?? true;
        const tickCount     = clamp(s.gaugeSettings.tickCount?.value ?? 5, 2, 20);
        const showPct       = s.gaugeSettings.showPctComplete?.value ?? true;
        const showMinMax    = s.gaugeSettings.showMinMax?.value ?? true;

        // Zone settings
        const showZones  = s.zoneSettings?.showZones?.value ?? true;
        const z1Max      = clamp(s.zoneSettings?.zone1MaxPct?.value ?? 50, 0, 100) / 100;
        const z2Max      = clamp(s.zoneSettings?.zone2MaxPct?.value ?? 80, 0, 100) / 100;
        const zone1Color = s.zoneSettings?.zone1Color?.value?.value ?? "#EF4444";
        const zone2Color = s.zoneSettings?.zone2Color?.value?.value ?? "#F59E0B";
        const zone3Color = s.zoneSettings?.zone3Color?.value?.value ?? "#10B981";
        const trackColor = s.zoneSettings?.trackColor?.value?.value ?? "#E5E7EB";

        // Font settings
        const fontFam    = String(s.fontSettings?.fontFamily?.value?.value ?? "Segoe UI");
        const boldValue  = s.fontSettings?.boldValue?.value ?? true;
        const valPrefix  = s.fontSettings?.valuePrefix?.value ?? "";
        const valSuffix  = s.fontSettings?.valueSuffix?.value ?? "";
        const valueColor = s.fontSettings?.valueColor?.value?.value ?? "#111827";

        // Gauge arc angles based on type
        let arcStartAngle = -Math.PI / 2; // default: semi
        let arcEndAngle   =  Math.PI / 2;
        if (gaugeType === "threequarter") {
            arcStartAngle = -Math.PI * 0.75;
            arcEndAngle   =  Math.PI * 0.75;
        } else if (gaugeType === "full") {
            arcStartAngle = -Math.PI;
            arcEndAngle   =  Math.PI;
        }
        const totalSweep = arcEndAngle - arcStartAngle;

        // ── Geometry ───────────────────────────────────────────────────────
        const pad    = 16;
        const isFull = gaugeType === "full";
        const cy     = isFull
            ? clamp(height / 2, pad + 20, height - pad)
            : clamp(height * 0.60, pad + 20, height - 50);
        const outerR = clamp(Math.min((width - 2 * pad) / 2, isFull ? (height - 2 * pad) / 2 : cy - pad), 20, 600);
        const innerR = outerR * (1 - arcThickPct);
        const cx     = width / 2;

        const range = this.maxVal - this.minVal;
        const hasValue = this.value !== null;

        const valueToAngle2 = (v: number) => {
            const pct = range === 0 ? 0 : clamp((v - this.minVal) / range, 0, 1);
            return arcStartAngle + pct * totalSweep;
        };

        const valueAngle = hasValue ? valueToAngle2(this.value!) : arcStartAngle;

        // Auto fill color based on zones
        const getZoneColor = (v: number): string => {
            const pct = range === 0 ? 0 : (v - this.minVal) / range;
            if (pct < z1Max) return zone1Color;
            if (pct < z2Max) return zone2Color;
            return zone3Color;
        };

        const fillColor = isManual
            ? manualColor
            : (hasValue ? getZoneColor(this.value!) : zone3Color);

        // ── Clear and resize SVG ──────────────────────────────────────────
        this.svg.attr("width", width).attr("height", height);
        this.svg.selectAll("*").remove();

        const arcGen = d3.arc<d3.DefaultArcObject>();
        const makeArc = (sa: number, ea: number): d3.DefaultArcObject =>
            ({ innerRadius: innerR, outerRadius: outerR, startAngle: sa, endAngle: ea, padAngle: 0 });

        const gMain = this.svg.append("g").attr("transform", `translate(${cx},${cy})`);

        // ── Zone arcs (background) ─────────────────────────────────────────
        if (showZones) {
            const zones = [
                { start: arcStartAngle, end: arcStartAngle + totalSweep * z1Max, color: zone1Color },
                { start: arcStartAngle + totalSweep * z1Max, end: arcStartAngle + totalSweep * z2Max, color: zone2Color },
                { start: arcStartAngle + totalSweep * z2Max, end: arcEndAngle, color: zone3Color }
            ];
            for (const z of zones) {
                if (z.end > z.start) {
                    gMain.append("path")
                        .attr("d", arcGen(makeArc(z.start, z.end)) ?? "")
                        .attr("fill", z.color)
                        .attr("opacity", 0.2);
                }
            }
        }

        // Track arc (full background)
        gMain.append("path")
            .attr("d", arcGen(makeArc(arcStartAngle, arcEndAngle)) ?? "")
            .attr("fill", "none")
            .attr("stroke", trackColor)
            .attr("stroke-width", outerR - innerR)
            .attr("fill-rule", "evenodd");

        // ── Value fill arc ─────────────────────────────────────────────────
        if (hasValue && valueAngle > arcStartAngle) {
            gMain.append("path")
                .attr("d", arcGen(makeArc(arcStartAngle, valueAngle)) ?? "")
                .attr("fill", "none")
                .attr("stroke", fillColor)
                .attr("stroke-width", outerR - innerR);
        }

        // ── Tick marks ────────────────────────────────────────────────────
        if (showTicks) {
            for (let i = 0; i <= tickCount; i++) {
                const pct = i / tickCount;
                const angle = arcStartAngle + pct * totalSweep;
                const sin = Math.sin(angle), cos = Math.cos(angle);
                const r1 = outerR + 4, r2 = outerR + 10;
                gMain.append("line")
                    .attr("x1", sin * r1).attr("y1", -cos * r1)
                    .attr("x2", sin * r2).attr("y2", -cos * r2)
                    .attr("stroke", "#6B7280")
                    .attr("stroke-width", i === 0 || i === tickCount ? 1.5 : 1);
            }
        }

        // ── Target marker ─────────────────────────────────────────────────
        if (showTarget && this.targetVal !== null && range !== 0) {
            const tAngle = valueToAngle2(this.targetVal);
            const sin = Math.sin(tAngle), cos = Math.cos(tAngle);
            const r1 = innerR - 4, r2 = outerR + 8;
            gMain.append("line")
                .attr("x1", sin * r1).attr("y1", -cos * r1)
                .attr("x2", sin * r2).attr("y2", -cos * r2)
                .attr("stroke", targetColor).attr("stroke-width", 2.5)
                .attr("stroke-linecap", "round");
        }

        // ── Needle ────────────────────────────────────────────────────────
        if (showNeedle && hasValue) {
            const sin = Math.sin(valueAngle), cos = Math.cos(valueAngle);
            const needleLen = outerR - 4;
            const needleW   = Math.max(2, outerR * 0.03);
            const perpSin   = Math.cos(valueAngle), perpCos = Math.sin(valueAngle);
            const points = [
                `${sin * needleLen},${-cos * needleLen}`,
                `${-perpSin * needleW},${perpCos * needleW}`,
                `${perpSin * needleW},${-perpCos * needleW}`
            ].join(" ");
            gMain.append("polygon")
                .attr("points", points)
                .attr("fill", "#374151")
                .attr("opacity", 0.9);
            gMain.append("circle").attr("r", Math.max(4, outerR * 0.05)).attr("fill", "#374151");
        }

        // ── Min / Max labels ──────────────────────────────────────────────
        if (showMinMax) {
            const labelFontSize = clamp(outerR * 0.12, 9, 14);
            const minAngle = arcStartAngle, maxAngle = arcEndAngle;
            const minSin = Math.sin(minAngle), minCos = Math.cos(minAngle);
            const maxSin = Math.sin(maxAngle), maxCos = Math.cos(maxAngle);
            const lr = outerR + 14;

            gMain.append("text")
                .attr("x", minSin * lr).attr("y", -minCos * lr + 4)
                .attr("text-anchor", minSin < 0 ? "end" : "middle")
                .attr("class", "gauge-minmax").style("font-size", `${labelFontSize}px`)
                .style("font-family", fontFam).text(formatNumber(this.minVal));

            gMain.append("text")
                .attr("x", maxSin * lr).attr("y", -maxCos * lr + 4)
                .attr("text-anchor", maxSin > 0 ? "start" : "middle")
                .attr("class", "gauge-minmax").style("font-size", `${labelFontSize}px`)
                .style("font-family", fontFam).text(formatNumber(this.maxVal));
        }

        // ── Value text ────────────────────────────────────────────────────
        const textZone  = isFull ? outerR * 0.4 : height - cy;
        const valueY    = isFull ? outerR * 0.15 : cy + textZone * 0.38;
        const valueText = hasValue ? `${valPrefix}${formatNumber(this.value!)}${valSuffix}` : "\u2014";

        this.svg.append("text")
            .attr("x", cx).attr("y", valueY)
            .attr("text-anchor", "middle").attr("class", "gauge-value")
            .style("font-size", `${valueFontSize}px`)
            .style("font-family", fontFam)
            .style("font-weight", boldValue ? "bold" : "normal")
            .style("fill", hasValue ? valueColor : "#9CA3AF")
            .text(valueText);

        // ── % complete ────────────────────────────────────────────────────
        if (showPct && hasValue && range !== 0) {
            const pct = clamp((this.value! - this.minVal) / range, 0, 1) * 100;
            this.svg.append("text")
                .attr("x", cx).attr("y", valueY + valueFontSize * 0.85 + 6)
                .attr("text-anchor", "middle").attr("class", "gauge-pct")
                .style("font-family", fontFam)
                .text(`${pct.toFixed(1)}% of max`);
        }
    }

    // ── Formatting model ─────────────────────────────────────────────────────

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
