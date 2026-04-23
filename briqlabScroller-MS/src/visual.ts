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

const TRIAL_DAYS = 4;
const LS_TRIAL_KEY = "briqlab_trial_BriqlabScroller_start";

interface TickerItem {
    label: string;
    value: number | null;
    change: number | null;
    hasChange: boolean;
}

export class Visual implements IVisual {
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private selMgr!:      powerbi.extensibility.ISelectionManager;
    private _handlersAttached = false;
    private target: HTMLElement;
    private host: IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private formattingSettings!: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    // DOM
    private rootEl!: HTMLElement;
    private viewportEl!: HTMLElement;
    private trackEl!: HTMLElement;
    private overlayEl!: HTMLElement;
    private trialBadgeEl!: HTMLElement;
    private proBadgeEl!: HTMLElement;

    // Animation state
    private animationId: number | null = null;
    private position: number = 0;
    private lastTimestamp: number | null = null;
    private isPaused: boolean = false;
    private contentWidth: number = 0;
    private items: TickerItem[] = [];

    // Trial / pro
    private trialStart: number = 0;
    private isPro: boolean = false;
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
    }

    // ── DOM ──────────────────────────────────────────────────────────────────

    private initDom(): void {
        this.rootEl = document.createElement("div");
        this.rootEl.className = "briqlab-scroller-root";

        this.viewportEl = document.createElement("div");
        this.viewportEl.className = "briqlab-scroller-viewport";

        this.trackEl = document.createElement("div");
        this.trackEl.className = "briqlab-scroller-track";

        this.viewportEl.appendChild(this.trackEl);
        this.rootEl.appendChild(this.viewportEl);

        // Trial badge (bottom-right)
        this.trialBadgeEl = document.createElement("div");
        this.trialBadgeEl.className = "briqlab-trial-badge hidden";
        this.rootEl.appendChild(this.trialBadgeEl);

        // Pro badge (bottom-right)
        this.proBadgeEl = document.createElement("div");
        this.proBadgeEl.className = "briqlab-pro-badge hidden";
        this.rootEl.appendChild(this.proBadgeEl);

        // Trial expired overlay
        this.overlayEl = document.createElement("div");
        this.overlayEl.className = "briqlab-trial-overlay hidden";

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
        this.overlayEl.appendChild(card);
        this.rootEl.appendChild(this.overlayEl);

        // Hover pause
        this.viewportEl.addEventListener("mouseenter", () => { this.isPaused = true; });
        this.viewportEl.addEventListener("mouseleave", () => { this.isPaused = false; });

        this.target.appendChild(this.rootEl);
    }

    // ── Trial / Pro ──────────────────────────────────────────────────────────

    private initTrial(): void {
        try {
            let stored = localStorage.getItem(LS_TRIAL_KEY);
            if (!stored) {
                stored = new Date().toISOString();
                localStorage.setItem(LS_TRIAL_KEY, stored);
            }
            this.trialStart = new Date(stored).getTime();
        } catch {
            this.trialStart = Date.now();
        }
    }

    private getTrialDaysElapsed(): number {
        return Math.floor((Date.now() - this.trialStart) / (1000 * 60 * 60 * 24));
    }

    private async validateProKey(key: string): Promise<boolean> {
        return checkMicrosoftLicence(this.host);
    }

    private updateTrialUI(daysElapsed: number): void {
        checkMicrosoftLicence(this.host).then(p => this._msUpdateLicenceUI(p)).catch(() => this._msUpdateLicenceUI(false));
    }

    // ── Number formatting ────────────────────────────────────────────────────

    private formatNumber(val: number, decimals: number, useKM: boolean, prefix: string, suffix: string): string {
        let formatted: string;
        if (useKM) {
            if (Math.abs(val) >= 1_000_000_000) {
                formatted = (val / 1_000_000_000).toFixed(decimals) + "B";
            } else if (Math.abs(val) >= 1_000_000) {
                formatted = (val / 1_000_000).toFixed(decimals) + "M";
            } else if (Math.abs(val) >= 1_000) {
                formatted = (val / 1_000).toFixed(decimals) + "K";
            } else {
                formatted = val.toFixed(decimals);
            }
        } else {
            formatted = val.toLocaleString(undefined, {
                minimumFractionDigits: decimals,
                maximumFractionDigits: decimals
            });
        }
        return `${prefix}${formatted}${suffix}`;
    }

    // ── Render ticker items ──────────────────────────────────────────────────

    private buildTrackContent(items: TickerItem[], settings: VisualFormattingSettingsModel): DocumentFragment {
        const frag = document.createDocumentFragment();
        const sep = settings.separatorSettings.separatorChar.value || "  |  ";
        const sepColor = settings.separatorSettings.separatorColor.value?.value || "#4B5563";
        const showArrows = settings.indicatorSettings.showArrows.value;
        const showValue = settings.indicatorSettings.showValue.value;
        const showChange = settings.indicatorSettings.showChange.value;
        const invert = settings.indicatorSettings.invertColors.value;
        const posColor = settings.indicatorSettings.positiveColor.value?.value || "#10B981";
        const negColor = settings.indicatorSettings.negativeColor.value?.value || "#EF4444";
        const decimals = settings.numberSettings.decimalPlaces.value ?? 2;
        const useKM = settings.numberSettings.useKM.value;
        const prefix = settings.numberSettings.prefix.value || "";
        const suffix = settings.numberSettings.suffix.value || "";
        const changePrefix = settings.numberSettings.changePrefix.value || "";
        const changeSuffix = settings.numberSettings.changeSuffix.value || "%";
        const gap = settings.scrollSettings.gap.value ?? 40;

        items.forEach((item, idx) => {
            const itemEl = document.createElement("span");
            itemEl.className = "ticker-item";
            itemEl.style.marginRight = `${gap}px`;

            // Label
            const labelEl = document.createElement("span");
            labelEl.className = "ticker-label";
            labelEl.textContent = item.label;
            itemEl.appendChild(labelEl);

            // Value
            if (showValue && item.value !== null) {
                const valEl = document.createElement("span");
                valEl.className = "ticker-value";
                valEl.textContent = " " + this.formatNumber(item.value, decimals, useKM, prefix, suffix);
                itemEl.appendChild(valEl);
            }

            // Change + arrow
            if (showChange && item.hasChange && item.change !== null) {
                const changeEl = document.createElement("span");
                changeEl.className = "ticker-change";

                const isPositive = item.change >= 0;
                const isGood = invert ? !isPositive : isPositive;
                const color = isGood ? posColor : negColor;
                changeEl.style.color = color;

                let arrow = "";
                if (showArrows) {
                    arrow = isPositive ? " ▲" : " ▼";
                }

                const sign = isPositive ? "+" : "";
                const changeFormatted = this.formatNumber(
                    Math.abs(item.change),
                    decimals,
                    useKM,
                    changePrefix,
                    changeSuffix
                );
                changeEl.textContent = `${arrow} ${sign}${changeFormatted}`;
                itemEl.appendChild(changeEl);
            }

            frag.appendChild(itemEl);

            // Separator (not after last)
            if (idx < items.length - 1) {
                const sepEl = document.createElement("span");
                sepEl.className = "ticker-separator";
                sepEl.style.color = sepColor;
                sepEl.textContent = sep;
                frag.appendChild(sepEl);
            }
        });

        return frag;
    }

    private buildCustomTextContent(text: string, settings: VisualFormattingSettingsModel): DocumentFragment {
        const frag = document.createDocumentFragment();
        const span = document.createElement("span");
        span.className = "ticker-custom-text";
        span.textContent = text;
        frag.appendChild(span);
        return frag;
    }

    private clearElement(el: HTMLElement): void {
        while (el.firstChild) el.removeChild(el.firstChild);
    }

    private rebuildTrack(settings: VisualFormattingSettingsModel): void {
        this.clearElement(this.trackEl);

        const useCustom = settings.customTextSettings.useCustomText.value;
        const customText = settings.customTextSettings.customText.value || "";

        if (useCustom) {
            if (!customText.trim()) {
                this.showEmptyState();
                return;
            }
            // Duplicate for seamless loop
            this.trackEl.appendChild(this.buildCustomTextContent(customText, settings));
            this.trackEl.appendChild(this.buildCustomTextContent(customText, settings));
        } else {
            if (this.items.length === 0) {
                this.showEmptyState();
                return;
            }
            this.trackEl.appendChild(this.buildTrackContent(this.items, settings));
            this.trackEl.appendChild(this.buildTrackContent(this.items, settings));
        }

        // Measure content width (half — one copy)
        requestAnimationFrame(() => {
            this.contentWidth = this.trackEl.scrollWidth / 2;
        });
    }

    private showEmptyState(): void {
        this.clearElement(this.trackEl);
        const empty = document.createElement("div");
        empty.className = "briqlab-empty-state";

        const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        svg.setAttribute("width", "24");
        svg.setAttribute("height", "20");
        svg.setAttribute("viewBox", "0 0 24 20");
        svg.setAttribute("fill", "none");
        const rects: [number, number, number, number, number][] = [
            [0, 0, 24, 3, 1.5],
            [4, 6, 20, 3, 1.5],
            [2, 12, 22, 3, 1.5],
            [6, 18, 18, 2, 1]
        ];
        for (const [x, y, w, h, r] of rects) {
            const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
            rect.setAttribute("x", String(x));
            rect.setAttribute("y", String(y));
            rect.setAttribute("width", String(w));
            rect.setAttribute("height", String(h));
            rect.setAttribute("rx", String(r));
            rect.setAttribute("fill", "#9CA3AF");
            svg.appendChild(rect);
        }

        const label = document.createElement("span");
        label.textContent = "Connect a Category or enable Custom Text";

        empty.appendChild(svg);
        empty.appendChild(label);
        this.trackEl.appendChild(empty);
        this.contentWidth = 0;
    }

    // ── Animation loop ───────────────────────────────────────────────────────

    private startAnimation(): void {
        if (this.animationId !== null) {
            cancelAnimationFrame(this.animationId);
        }
        this.lastTimestamp = null;
        this.animationId = requestAnimationFrame(this.animate.bind(this));
    }

    private animate(timestamp: number): void {
        if (this.lastTimestamp === null) this.lastTimestamp = timestamp;
        const delta = timestamp - this.lastTimestamp;
        this.lastTimestamp = timestamp;

        const speed = this.formattingSettings?.scrollSettings?.speed?.value ?? 60;
        const direction = String(this.formattingSettings?.scrollSettings?.direction?.value?.value ?? "left");
        const pauseOnHover = this.formattingSettings?.scrollSettings?.pauseOnHover?.value ?? true;

        if (speed > 0 && this.contentWidth > 0 && !(pauseOnHover && this.isPaused)) {
            const delta_px = (speed * delta) / 1000;

            if (direction === "left") {
                this.position -= delta_px;
                if (this.position <= -this.contentWidth) {
                    this.position += this.contentWidth;
                }
            } else {
                this.position += delta_px;
                if (this.position >= this.contentWidth) {
                    this.position -= this.contentWidth;
                }
            }

            this.trackEl.style.transform = `translateX(${this.position}px)`;
        }

        this.animationId = requestAnimationFrame(this.animate.bind(this));
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
                        dataItems: [{ displayName: "Briqlab Scroller", value: "" }],
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
    
            // Pro key validation
            const proKey = ""; // MS cert: pro key field removed
            if (proKey && !this.proKeyCache.has(proKey)) {
                this.validateProKey(proKey).then(valid => {
                    this.isPro = valid;
                    this.updateTrialUI(this.getTrialDaysElapsed());
                });
            } else if (proKey && this.proKeyCache.has(proKey)) {
                this.isPro = this.proKeyCache.get(proKey)!;
            } else {
                this.isPro = false;
            }
    
            // Parse data
            this.items = this.parseData(options);
    
            // Apply appearance
            this.applyAppearance();
    
            // Rebuild ticker content
            this.rebuildTrack(this.formattingSettings);
    
            // Trial gating
            this.updateTrialUI(this.getTrialDaysElapsed());
    
            // Start animation loop if not running
            if (this.animationId === null) {
                this.startAnimation();
            }
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    private parseData(options: VisualUpdateOptions): TickerItem[] {
        const dv = options.dataViews?.[0];
        if (!dv?.categorical) return [];

        const cats = dv.categorical.categories?.[0];
        if (!cats) return [];

        const allValues = dv.categorical.values ?? [];
        let valueCol: powerbi.DataViewValueColumn | null = null;
        let changeCol: powerbi.DataViewValueColumn | null = null;

        for (const col of allValues) {
            const source = col.source;
            if (source.roles?.["value"]) valueCol = col;
            if (source.roles?.["changeValue"]) changeCol = col;
        }

        const results: TickerItem[] = [];
        for (let i = 0; i < cats.values.length; i++) {
            const label = cats.values[i] != null ? String(cats.values[i]) : "(blank)";
            const value = valueCol ? (valueCol.values[i] as number | null) : null;
            const change = changeCol ? (changeCol.values[i] as number | null) : null;
            results.push({
                label,
                value,
                change,
                hasChange: changeCol !== null
            });
        }
        return results;
    }

    private applyAppearance(): void {
        const s = this.formattingSettings.appearanceSettings;
        const transparent = s.transparentBackground.value;
        const bg = s.backgroundColor.value?.value || "#111827";
        const textColor = s.textColor.value?.value || "#FFFFFF";
        const fontSize = s.fontSize.value ?? 14;
        const fontFamily = String(s.fontFamily.value?.value || "Segoe UI");
        const padding = s.padding.value ?? 8;

        this.rootEl.style.backgroundColor = transparent ? "transparent" : bg;
        this.rootEl.style.color = textColor;
        this.rootEl.style.fontSize = `${fontSize}px`;
        this.rootEl.style.fontFamily = fontFamily;
        this.viewportEl.style.padding = `${padding}px 0`;
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
