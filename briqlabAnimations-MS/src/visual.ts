"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions      = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual                  = powerbi.extensibility.visual.IVisual;
import IVisualHost              = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager        = powerbi.extensibility.ISelectionManager;
import ISelectionId             = powerbi.visuals.ISelectionId;

import { VisualFormattingSettingsModel } from "./settings";
import { checkMicrosoftLicence, resetLicenceCache } from "./licenceManager";
import { getTrialDaysRemaining, isTrialExpired, getPurchaseUrl, getButtonText } from "./trialManager";

const TRIAL_MS      = 4 * 24 * 60 * 60 * 1000;
const TRIAL_KEY     = "briqlab_trial_animations_start";
const PRO_STORE_KEY = "briqlab_animations_prokey";

function getTrial(): { daysLeft: number; expired: boolean } {
    try {
        let raw = localStorage.getItem(TRIAL_KEY);
        if (!raw) { raw = String(Date.now()); localStorage.setItem(TRIAL_KEY, raw); }
        const elapsed  = Date.now() - parseInt(raw, 10);
        const daysLeft = Math.max(0, Math.ceil((TRIAL_MS - elapsed) / 86400000));
        return { daysLeft, expired: elapsed > TRIAL_MS };
    } catch { return { daysLeft: 4, expired: false }; }
}

interface KpiItem {
    label:       string;
    value:       number;
    target:      number | null;
    selectionId: ISelectionId;
}

export class Visual implements IVisual {
    private readonly host:       IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private readonly selMgr:     ISelectionManager;
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private _handlersAttached = false;
    private readonly fmtSvc:     FormattingSettingsService;

    private readonly root:        HTMLElement;
    private readonly contentEl:   HTMLDivElement;
    private readonly cardsEl:     HTMLDivElement;
    private readonly trialBadge:  HTMLDivElement;
    private readonly proBadge:    HTMLDivElement;
    private readonly keyErrorEl:  HTMLDivElement;
    private readonly overlayEl:   HTMLDivElement;

    private settings!:  VisualFormattingSettingsModel;
    private vp:         powerbi.IViewport = { width: 300, height: 200 };
    private items:      KpiItem[] = [];
    private selected:   Set<string> = new Set();

    private isPro   = false;
    private lastKey = "";
    private readonly keyCache: Map<string, boolean> = new Map();

    // Tracks active animation frame handles so we can cancel on re-render
    private animHandles: number[] = [];

    constructor(options: VisualConstructorOptions) {
        this.host   = options.host;
        this.renderingManager = options.host.eventService;
        this.selMgr = options.host.createSelectionManager();
        this.tooltipSvc = this.host.tooltipService;
        this.fmtSvc = new FormattingSettingsService();

        this.root = options.element;
        this.root.classList.add("briqlab-animations");

        this.contentEl = this.mkDiv("briqlab-anim-content");
        this.root.appendChild(this.contentEl);

        this.cardsEl = this.mkDiv("briqlab-anim-cards");
        this.contentEl.appendChild(this.cardsEl);

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
        const card  = this.mkDiv("briqlab-trial-card");
        const title = document.createElement("h2");
        title.className   = "trial-title";
        title.textContent = "Free trial ended";
        card.appendChild(title);

        const body = document.createElement("p");
        body.className   = "trial-body";
        body.textContent = "Activate Briqlab Pro to continue using this visual and unlock all features.";
        card.appendChild(body);

        const btn = document.createElement("button");
        btn.className   = "trial-btn";
        btn.textContent = getButtonText();
        btn.addEventListener("click", () => this.host.launchUrl(getPurchaseUrl()));
        card.appendChild(btn);

        const sub = document.createElement("p");
        sub.className   = "trial-subtext";
        sub.textContent = "Purchase on Microsoft AppSource to unlock all features instantly.";
        card.appendChild(sub);

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
                        dataItems: [{ displayName: "Briqlab Animated Counter", value: "" }],
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
            this.parseData(options);
            this.render();
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    private parseData(options: VisualUpdateOptions): void {
        this.items = [];
        const dv = options.dataViews?.[0];
        if (!dv?.categorical?.categories?.[0]) return;

        const cats    = dv.categorical.categories[0];
        let valCol: powerbi.DataViewValueColumn | undefined;
        let tgtCol: powerbi.DataViewValueColumn | undefined;

        for (const col of dv.categorical.values ?? []) {
            const roles = col.source.roles ?? {};
            if (roles["value"])  valCol = col;
            if (roles["target"]) tgtCol = col;
        }
        if (!valCol) return;

        for (let i = 0; i < cats.values.length; i++) {
            const label  = String(cats.values[i] ?? "");
            const value  = Number(valCol.values[i] ?? 0);
            const target = tgtCol ? (tgtCol.values[i] != null ? Number(tgtCol.values[i]) : null) : null;
            const selId  = this.host.createSelectionIdBuilder()
                .withCategory(cats, i)
                .createSelectionId();
            this.items.push({ label, value, target, selectionId: selId });
        }
    }

    private render(): void {
        this.cancelAnimations();

        const as = this.settings.animationSettings;
        const ss = this.settings.styleSettings;
        const duration     = Math.max(200, as.duration.value ?? 1500);
        const easing       = String(as.easing.value?.value ?? "easeOut");
        const showProgress = as.showProgressBar.value ?? true;
        const autoPlay     = as.autoPlay.value ?? true;

        const primaryColor = ss.primaryColor.value?.value ?? "#0D9488";
        const secondaryColor = ss.secondaryColor.value?.value ?? "#E2E8F0";
        const bgColor      = ss.backgroundColor.value?.value ?? "#FFFFFF";
        const fontFamily   = String(ss.fontFamily.value?.value ?? "Segoe UI");
        const valueFontSize = ss.valueFontSize.value ?? 36;
        const labelFontSize = ss.labelFontSize.value ?? 13;
        const layout       = String(ss.cardLayout.value?.value ?? "grid");

        this.root.style.setProperty("--briq-anim-primary",    primaryColor);
        this.root.style.setProperty("--briq-anim-secondary",  secondaryColor);
        this.root.style.setProperty("--briq-anim-bg",         bgColor);
        this.root.style.setProperty("--briq-anim-font",       fontFamily);
        this.root.style.setProperty("--briq-anim-value-size", `${valueFontSize}px`);
        this.root.style.setProperty("--briq-anim-label-size", `${labelFontSize}px`);

        while (this.cardsEl.firstChild) this.cardsEl.removeChild(this.cardsEl.firstChild);

        if (this.items.length === 0) {
            const empty = this.mkDiv("briq-anim-empty");
            empty.textContent = "Add Label and Value fields to get started";
            this.cardsEl.appendChild(empty);
            return;
        }

        // Determine grid columns based on layout and viewport
        const cols = this.calcColumns(layout, this.items.length, this.vp.width);
        this.cardsEl.style.gridTemplateColumns = `repeat(${cols}, 1fr)`;

        this.items.forEach((item, idx) => {
            const card = this.mkDiv("briq-anim-card");
            card.style.backgroundColor = bgColor;
            if (this.selected.has(item.label)) card.classList.add("selected");

            const valueEl = document.createElement("div");
            valueEl.className   = "briq-anim-value";
            valueEl.textContent = this.fmtNum(0);
            card.appendChild(valueEl);

            if (item.target !== null) {
                const targetEl = document.createElement("div");
                targetEl.className   = "briq-anim-target";
                targetEl.textContent = `/ ${this.fmtNum(item.target)}`;
                card.appendChild(targetEl);
            }

            const labelEl = document.createElement("div");
            labelEl.className   = "briq-anim-label";
            labelEl.textContent = item.label;
            card.appendChild(labelEl);

            if (showProgress && item.target !== null && item.target > 0) {
                const barTrack = this.mkDiv("briq-anim-bar-track");
                const barFill  = this.mkDiv("briq-anim-bar-fill");
                barFill.style.backgroundColor = primaryColor;
                barTrack.style.backgroundColor = secondaryColor;
                barTrack.appendChild(barFill);
                card.appendChild(barTrack);

                if (autoPlay) {
                    const pct = Math.min(100, (item.value / item.target) * 100);
                    barFill.style.width = "0%";
                    const handle = window.setTimeout(() => {
                        barFill.style.transition = `width ${duration}ms ${this.cssEasing(easing)}`;
                        barFill.style.width      = `${pct}%`;
                    }, idx * 80);
                    this.animHandles.push(handle);
                } else {
                    const pct = Math.min(100, (item.value / item.target) * 100);
                    barFill.style.width = `${pct}%`;
                }
            }

            card.addEventListener("click", (e) => {
                e.stopPropagation();
                const isMulti = e.ctrlKey || e.metaKey;
                if (this.selected.has(item.label) && !isMulti) {
                    this.selected.clear();
                    this.selMgr.clear();
                } else {
                    if (!isMulti) this.selected.clear();
                    this.selected.add(item.label);
                    const ids = this.items
                        .filter(it => this.selected.has(it.label))
                        .map(it => it.selectionId);
                    this.selMgr.select(ids, isMulti);
                }
                // Refresh selected state without re-animating
                this.cardsEl.querySelectorAll<HTMLElement>(".briq-anim-card").forEach((el, i) => {
                    el.classList.toggle("selected", this.selected.has(this.items[i]?.label ?? ""));
                });
            });

            this.cardsEl.appendChild(card);

            if (autoPlay) {
                this.animateCount(valueEl, 0, item.value, duration, easing, idx * 80);
            } else {
                valueEl.textContent = this.fmtNum(item.value);
            }
        });
    }

    private animateCount(
        el: HTMLElement,
        from: number,
        to: number,
        duration: number,
        easing: string,
        delay: number
    ): void {
        const startTime = performance.now() + delay;
        const range     = to - from;

        const tick = (now: number) => {
            if (now < startTime) {
                this.animHandles.push(requestAnimationFrame(tick));
                return;
            }
            const elapsed  = now - startTime;
            const progress = Math.min(elapsed / duration, 1);
            const eased    = this.applyEasing(progress, easing);
            el.textContent = this.fmtNum(from + range * eased);
            if (progress < 1) {
                this.animHandles.push(requestAnimationFrame(tick));
            } else {
                el.textContent = this.fmtNum(to);
            }
        };
        this.animHandles.push(requestAnimationFrame(tick));
    }

    private cancelAnimations(): void {
        this.animHandles.forEach(h => {
            if (h > 1000) cancelAnimationFrame(h); // rAF handles are > 0
            else clearTimeout(h);
        });
        this.animHandles = [];
    }

    private applyEasing(t: number, easing: string): number {
        switch (easing) {
            case "linear":    return t;
            case "easeInOut": return t < 0.5 ? 2*t*t : -1+(4-2*t)*t;
            case "bounce": {
                if (t < 1/2.75)      return 7.5625*t*t;
                else if (t < 2/2.75) { t -= 1.5/2.75;   return 7.5625*t*t+0.75;   }
                else if (t < 2.5/2.75){ t -= 2.25/2.75; return 7.5625*t*t+0.9375; }
                else                 { t -= 2.625/2.75;  return 7.5625*t*t+0.984375; }
            }
            default: return t * (2 - t); // easeOut quad
        }
    }

    private cssEasing(easing: string): string {
        switch (easing) {
            case "linear":    return "linear";
            case "easeInOut": return "ease-in-out";
            case "bounce":    return "cubic-bezier(0.34,1.56,0.64,1)";
            default:          return "ease-out";
        }
    }

    private calcColumns(layout: string, count: number, width: number): number {
        if (layout === "column") return 1;
        if (layout === "row")    return count;
        const minCardWidth = 140;
        return Math.max(1, Math.min(count, Math.floor(width / minCardWidth)));
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
                if (!valid) {
                    this.lastKey = "";
                    try { localStorage.removeItem(PRO_STORE_KEY); } catch { /**/ }
                }
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

    private fmtNum(v: number): string {
        const a = Math.abs(v);
        if (a >= 1e9) return (v / 1e9).toFixed(1).replace(/\.0$/, "") + "B";
        if (a >= 1e6) return (v / 1e6).toFixed(1).replace(/\.0$/, "") + "M";
        if (a >= 1e3) return (v / 1e3).toFixed(1).replace(/\.0$/, "") + "K";
        return Math.round(v).toLocaleString();
    }

    private mkDiv(cls: string): HTMLDivElement {
        const el = document.createElement("div");
        el.className = cls;
        return el;
    }
}
