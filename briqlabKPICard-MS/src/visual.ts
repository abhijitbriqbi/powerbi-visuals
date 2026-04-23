"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;

import { VisualFormattingSettingsModel } from "./settings";
import { checkMicrosoftLicence, resetLicenceCache } from "./licenceManager";
import { getTrialDaysRemaining, isTrialExpired, getPurchaseUrl, getButtonText } from "./trialManager";

const TRIAL_MS = 4 * 24 * 60 * 60 * 1000;
const LS_TRIAL_KEY = "briqlab_trial_kpicard_start";
const LS_PRO_KEY = "briqlab_kpicard_prokey";

export class Visual implements IVisual {
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private selMgr!:      powerbi.extensibility.ISelectionManager;
    private _handlersAttached = false;
    private target: HTMLElement;
    private host: IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private formattingSettings!: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    // DOM elements
    private rootEl!: HTMLElement;
    private contentEl!: HTMLElement;
    private valueEl!: HTMLElement;
    private labelEl!: HTMLElement;
    private trendEl!: HTMLElement;
    private trialOverlayEl!: HTMLElement;
    private trialBadgeEl!: HTMLElement;
    private proBadgeEl!: HTMLElement;
    private keyErrorEl!: HTMLElement;

    // Data state
    private lastKpiValue: number | null = null;
    private lastLabelText: string = "";
    private lastComparisonValue: number | null = null;
    private previousKpiValue: number | null = null;

    // Trial / pro state
    private trialStart: number = 0;
    private currentProKey: string = "";
    private isPro: boolean = false;
    private proKeyValidating: boolean = false;
    private proKeyCache: Map<string, boolean> = new Map();

    constructor(options: VisualConstructorOptions) {
        this.formattingSettingsService = new FormattingSettingsService();
        this.target     = options.element;
        this.host       = options.host;
        this.renderingManager = options.host.eventService;
        this.selMgr     = this.host.createSelectionManager();
        this.tooltipSvc = this.host.tooltipService;

        this.initDom();
        this.initTrial();

        // Restore saved pro key from localStorage
        try {
            const saved = localStorage.getItem(LS_PRO_KEY) ?? "";
            if (saved) {
                this.currentProKey = saved;
                this.validateAndApplyProKey(saved);
            }
        } catch {
            // localStorage may be unavailable in some environments
        }
    }

    // ── DOM construction ─────────────────────────────────────────────────────

    private initDom(): void {
        // Root
        const root = document.createElement("div");
        root.className = "briqlab-kpi-card";
        this.rootEl = root;

        // Visual content wrapper (blurred when trial expired)
        const content = document.createElement("div");
        content.className = "briqlab-visual-content";
        this.contentEl = content;

        // Inner card
        const inner = document.createElement("div");
        inner.className = "kpi-card-inner";

        this.labelEl = document.createElement("div");
        this.labelEl.className = "kpi-label";

        this.valueEl = document.createElement("div");
        this.valueEl.className = "kpi-value";

        this.trendEl = document.createElement("div");
        this.trendEl.className = "kpi-trend";

        inner.appendChild(this.labelEl);
        inner.appendChild(this.valueEl);
        inner.appendChild(this.trendEl);
        content.appendChild(inner);
        root.appendChild(content);

        // Trial badge (bottom-left)
        const trialBadge = document.createElement("div");
        trialBadge.className = "briqlab-trial-badge hidden";
        this.trialBadgeEl = trialBadge;
        root.appendChild(trialBadge);

        // Pro badge (bottom-right)
        const proBadge = document.createElement("div");
        proBadge.className = "briqlab-pro-badge hidden";
        this.proBadgeEl = proBadge;
        root.appendChild(proBadge);

        // Key error label
        const keyError = document.createElement("div");
        keyError.className = "briqlab-key-error hidden";
        this.keyErrorEl = keyError;
        root.appendChild(keyError);

        // Trial expired overlay
        const overlay = document.createElement("div");
        overlay.className = "briqlab-trial-overlay hidden";
        this.trialOverlayEl = overlay;

        const card = document.createElement("div");
        card.className = "briqlab-trial-card";

        const title = document.createElement("div");
        title.className = "trial-title";
        title.textContent = "Free trial ended";

        const body = document.createElement("div");
        body.className = "trial-body";
        body.textContent = "Activate Briqlab Pro to continue using this visual and unlock all features.";

        const btn = document.createElement("button");
        btn.className = "trial-btn";
        btn.textContent = getButtonText();
        btn.addEventListener("click", () => {
            this.host.launchUrl(getPurchaseUrl());
        });

        const sub = document.createElement("div");
        sub.className = "trial-subtext";
        sub.textContent = "Purchase on Microsoft AppSource to unlock all features instantly.";

        card.appendChild(title);
        card.appendChild(body);
        card.appendChild(btn);
        card.appendChild(sub);
        overlay.appendChild(card);
        root.appendChild(overlay);

        while (this.target.firstChild) {
            this.target.removeChild(this.target.firstChild);
        }
        this.target.appendChild(root);
    }

    // ── Trial management ─────────────────────────────────────────────────────

    private initTrial(): void {
        try {
            const stored = localStorage.getItem(LS_TRIAL_KEY);
            if (stored) {
                this.trialStart = parseInt(stored, 10);
            } else {
                this.trialStart = Date.now();
                localStorage.setItem(LS_TRIAL_KEY, String(this.trialStart));
            }
        } catch {
            this.trialStart = Date.now();
        }
    }

    private getTrialElapsedMs(): number {
        return Date.now() - this.trialStart;
    }

    private isTrialExpired(): boolean {
        return this.getTrialElapsedMs() > TRIAL_MS;
    }

    private trialDaysRemaining(): number {
        const remaining = TRIAL_MS - this.getTrialElapsedMs();
        return Math.max(0, Math.ceil(remaining / (24 * 60 * 60 * 1000)));
    }

    // ── Pro key validation ───────────────────────────────────────────────────

    private validateAndApplyProKey(key: string): void {
        checkMicrosoftLicence(this.host).then(p => { this._msUpdateLicenceUI(p); this.render(); }).catch(() => this._msUpdateLicenceUI(false));
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private formatNumber(value: number): string {
        const abs = Math.abs(value);
        if (abs >= 1_000_000) {
            const result = (value / 1_000_000).toFixed(1);
            return (result.endsWith(".0") ? result.slice(0, -2) : result) + "M";
        }
        if (abs >= 1_000) {
            const result = (value / 1_000).toFixed(1);
            return (result.endsWith(".0") ? result.slice(0, -2) : result) + "K";
        }
        return value.toLocaleString("en-US");
    }

    // ── Update ───────────────────────────────────────────────────────────────

    public update(options: VisualUpdateOptions): void {
        this.renderingManager.renderingStarted(options);
        try {
            // AppSource: attach context-menu + tooltip once
            if (!this._handlersAttached) {
                this._handlersAttached = true;
                this.target.addEventListener("contextmenu", (e: MouseEvent) => {
                    e.preventDefault();
                    this.selMgr.showContextMenu(
                        null as unknown as powerbi.visuals.ISelectionId,
                        { x: e.clientX, y: e.clientY }
                    );
                });
                this.target.addEventListener("mousemove", (e: MouseEvent) => {
                    this.tooltipSvc.show({
                        dataItems: [{ displayName: "Briqlab KPI Card", value: "" }],
                        identities: [],
                        coordinates: [e.clientX, e.clientY],
                        isTouchEvent: false
                    });
                });
                this.target.addEventListener("mouseleave", () => {
                    this.tooltipSvc.hide({ isTouchEvent: false, immediately: false });
                });
            }
            this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
                VisualFormattingSettingsModel,
                options.dataViews?.[0]
            );
    
            // Extract data
            this.previousKpiValue = this.lastKpiValue;
            this.lastKpiValue = null;
            this.lastLabelText = "";
            this.lastComparisonValue = null;
    
            const dataView = options.dataViews?.[0];
            if (dataView?.categorical) {
                const cat = dataView.categorical;
    
                if (cat.categories && cat.categories.length > 0) {
                    this.lastLabelText = cat.categories[0].values?.[0]?.toString() ?? "";
                }
    
                if (cat.values) {
                    for (const col of cat.values) {
                        const raw = col.values?.[0];
                        const num = typeof raw === "number" ? raw : null;
                        if (col.source.roles?.["measure"] && num !== null) {
                            this.lastKpiValue = num;
                        }
                        if (col.source.roles?.["comparisonMeasure"] && num !== null) {
                            this.lastComparisonValue = num;
                        }
                    }
                }
            }
    
            // Label override from format pane if no data bound
            if (!this.lastLabelText) {
                // label stays empty; visual.ts does not expose a label text setting
            }
    
            // Handle pro key changes
            const proKey = ""; // MS cert: pro key field removed
            if (proKey !== this.currentProKey) {
                this.currentProKey = proKey;
                this.isPro = false;
    
                try {
                    if (proKey) {
                        localStorage.setItem(LS_PRO_KEY, proKey);
                    } else {
                        localStorage.removeItem(LS_PRO_KEY);
                    }
                } catch {
                    // ignore
                }
    
                if (proKey) {
                    this.validateAndApplyProKey(proKey);
                    // render will be called after validation completes
                }
            }
    
            this.render();
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    // ── Render ───────────────────────────────────────────────────────────────

    private render(): void {
        if (!this.formattingSettings) return;

        const s = this.formattingSettings;
        const trialExpired = !this.isPro && this.isTrialExpired();

        // ── Display settings (always applied, no isPro gate) ─────────────────
        const valueFontSize = Math.min(72, Math.max(8, s.displaySettings.valueFontSize.value));
        const labelFontSize = Math.min(24, Math.max(8, s.displaySettings.labelFontSize.value));
        const primaryColor = s.displaySettings.primaryColor.value?.value ?? "#0D9488";
        const bgColor = s.displaySettings.backgroundColor.value?.value ?? "#FFFFFF";
        const showBorder = s.displaySettings.showBorder.value;
        const borderColor = s.displaySettings.borderColor.value?.value ?? "#E2E8F0";
        const borderRadius = Math.min(20, Math.max(0, s.displaySettings.borderRadius.value));

        // Apply container styles
        this.rootEl.style.backgroundColor = bgColor;
        this.rootEl.style.borderRadius = `${borderRadius}px`;
        this.rootEl.style.border = showBorder ? `1px solid ${borderColor}` : "none";

        // ── Content blur ─────────────────────────────────────────────────────
        if (trialExpired) {
            this.contentEl.classList.add("blurred");
        } else {
            this.contentEl.classList.remove("blurred");
        }

        // ── KPI value ────────────────────────────────────────────────────────
        const newValueText = this.lastKpiValue !== null
            ? this.formatNumber(this.lastKpiValue)
            : "\u2014";

        const valueChanged = this.lastKpiValue !== this.previousKpiValue;
        this.valueEl.style.fontSize = `${valueFontSize}px`;
        this.valueEl.style.color = primaryColor;
        this.valueEl.textContent = newValueText;

        // Entry animation: pop class triggers scale-up then remove
        if (valueChanged && this.lastKpiValue !== null) {
            this.valueEl.classList.remove("pop");
            // Force reflow to restart animation
            void this.valueEl.offsetWidth;
            this.valueEl.classList.add("pop");
            setTimeout(() => {
                this.valueEl.classList.remove("pop");
            }, 300);
        }

        // ── Label ────────────────────────────────────────────────────────────
        this.labelEl.style.fontSize = `${labelFontSize}px`;
        if (this.lastLabelText) {
            this.labelEl.textContent = this.lastLabelText;
            this.labelEl.style.display = "";
        } else {
            this.labelEl.textContent = "";
            this.labelEl.style.display = "none";
        }

        // ── Trend ────────────────────────────────────────────────────────────
        if (
            this.lastKpiValue !== null &&
            this.lastComparisonValue !== null &&
            this.lastComparisonValue !== 0
        ) {
            const pct =
                ((this.lastKpiValue - this.lastComparisonValue) /
                    Math.abs(this.lastComparisonValue)) *
                100;
            const positive = pct >= 0;
            this.trendEl.textContent = `${positive ? "\u25b2" : "\u25bc"} ${Math.abs(pct).toFixed(1)}%`;
            this.trendEl.className = `kpi-trend ${positive ? "positive" : "negative"}`;
            this.trendEl.style.display = "";
        } else {
            this.trendEl.textContent = "";
            this.trendEl.style.display = "none";
        }

        // ── Trial overlay ────────────────────────────────────────────────────
        if (trialExpired) {
            this.trialOverlayEl.classList.remove("hidden");
        } else {
            this.trialOverlayEl.classList.add("hidden");
        }

        // ── Trial badge ──────────────────────────────────────────────────────
        if (!this.isPro && !trialExpired) {
            const days = this.trialDaysRemaining();
            this.trialBadgeEl.textContent = `Trial: ${days} day${days === 1 ? "" : "s"} remaining`;
            this.trialBadgeEl.classList.remove("hidden");
        } else {
            this.trialBadgeEl.classList.add("hidden");
        }

        // ── Pro badge ────────────────────────────────────────────────────────
        if (this.isPro) {
            this.proBadgeEl.textContent = "\u2713 Pro Active";
            this.proBadgeEl.classList.remove("hidden");
        } else {
            this.proBadgeEl.classList.add("hidden");
        }

        // ── Key error ────────────────────────────────────────────────────────
        const keyEntered = this.currentProKey.length > 0;
        const keyInvalid = keyEntered && !this.isPro && !this.proKeyValidating &&
            this.proKeyCache.has(this.currentProKey) &&
            this.proKeyCache.get(this.currentProKey) === false;

        if (keyInvalid) {
            this.keyErrorEl.textContent = "\u2717 Invalid key";
            this.keyErrorEl.classList.remove("hidden");
        } else {
            this.keyErrorEl.classList.add("hidden");
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
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
