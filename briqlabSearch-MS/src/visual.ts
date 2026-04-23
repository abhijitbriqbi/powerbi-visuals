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

const TRIAL_MS   = 4 * 24 * 60 * 60 * 1000;
const TRIAL_KEY  = "briqlab_trial_search_start";
const PRO_KEY_LS = "briqlab_search_prokey";
const KEY_CACHE: Map<string, boolean> = new Map();

function trialDaysLeft(): number {
    let raw = localStorage.getItem(TRIAL_KEY);
    if (!raw) { localStorage.setItem(TRIAL_KEY, String(Date.now())); raw = String(Date.now()); }
    const elapsed = Date.now() - parseInt(raw, 10);
    return Math.max(0, Math.ceil((TRIAL_MS - elapsed) / 86400000));
}

async function validateKey(key: string): Promise<boolean> {
    // MS certification: no external API calls
    return false;
}

interface SearchItem {
    label:       string;
    value:       number | null;
    selectionId: ISelectionId;
}

export class Visual implements IVisual {
    private host:       IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private container:  HTMLElement;
    private fmtSvc:     FormattingSettingsService;
    private settings!:  VisualFormattingSettingsModel;
    private selMgr:     ISelectionManager;
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private _handlersAttached = false;
    private isPro:      boolean = false;
    private items:      SearchItem[] = [];
    private selected:   Set<string>  = new Set();

    // DOM refs
    private root!:        HTMLElement;
    private inputEl!:     HTMLInputElement;
    private listEl!:      HTMLElement;
    private countEl!:     HTMLElement;
    private trialBadge!:  HTMLElement;
    private proBadge!:    HTMLElement;
    private overlay!:     HTMLElement;

    constructor(options: VisualConstructorOptions) {
        this.host      = options.host as unknown as IVisualHost;
        this.renderingManager = options.host.eventService;
        this.fmtSvc    = new FormattingSettingsService();
        this.selMgr    = this.host.createSelectionManager();
        this.tooltipSvc = this.host.tooltipService;
        this.container = options.element as HTMLElement;
        this.buildDOM();
    }

    private buildDOM(): void {
        this.root = document.createElement("div");
        this.root.className = "briq-search-root";
        this.container.appendChild(this.root);

        // Content wrapper
        const content = document.createElement("div");
        content.className = "briq-search-content";
        this.root.appendChild(content);

        // Search input row
        const inputRow = document.createElement("div");
        inputRow.className = "briq-search-input-row";
        content.appendChild(inputRow);

        // Search icon
        const icon = document.createElement("span");
        icon.className = "briq-search-icon";
        const svgNS = "http://www.w3.org/2000/svg";
        const svgEl = document.createElementNS(svgNS, "svg");
        svgEl.setAttribute("width", "14"); svgEl.setAttribute("height", "14");
        svgEl.setAttribute("viewBox", "0 0 16 16"); svgEl.setAttribute("fill", "none");
        const circ = document.createElementNS(svgNS, "circle");
        circ.setAttribute("cx", "6.5"); circ.setAttribute("cy", "6.5"); circ.setAttribute("r", "5");
        circ.setAttribute("stroke", "currentColor"); circ.setAttribute("stroke-width", "1.5");
        svgEl.appendChild(circ);
        const pth = document.createElementNS(svgNS, "path");
        pth.setAttribute("d", "M10.5 10.5L14 14"); pth.setAttribute("stroke", "currentColor");
        pth.setAttribute("stroke-width", "1.5"); pth.setAttribute("stroke-linecap", "round");
        svgEl.appendChild(pth);
        icon.appendChild(svgEl);
        inputRow.appendChild(icon);

        this.inputEl = document.createElement("input");
        this.inputEl.type = "text";
        this.inputEl.className = "briq-search-input";
        this.inputEl.placeholder = "Search\u2026";
        inputRow.appendChild(this.inputEl);

        // Clear button
        const clearBtn = document.createElement("button");
        clearBtn.className = "briq-search-clear hidden";
        clearBtn.textContent = "\u2715";
        clearBtn.title = "Clear search";
        inputRow.appendChild(clearBtn);

        // Result count
        this.countEl = document.createElement("div");
        this.countEl.className = "briq-search-count hidden";
        content.appendChild(this.countEl);

        // Results list
        this.listEl = document.createElement("div");
        this.listEl.className = "briq-search-list";
        content.appendChild(this.listEl);

        // Trial badge (standard briqlab class so overlay CSS applies)
        this.trialBadge = document.createElement("div");
        this.trialBadge.className = "briqlab-trial-badge hidden";
        this.root.appendChild(this.trialBadge);

        // Pro badge
        this.proBadge = document.createElement("div");
        this.proBadge.className = "briqlab-pro-badge hidden";
        this.proBadge.textContent = "\u2713 Pro Active";
        this.root.appendChild(this.proBadge);

        // Trial overlay (standard briqlab classes — matches DrillPie card style)
        this.overlay = document.createElement("div");
        this.overlay.className = "briqlab-trial-overlay hidden";
        const ovCard  = document.createElement("div");    ovCard.className  = "briqlab-trial-card";
        const ovTitle = document.createElement("h2");     ovTitle.className = "trial-title";   ovTitle.textContent = "Free trial ended";                                               ovCard.appendChild(ovTitle);
        const ovBody  = document.createElement("p");      ovBody.className  = "trial-body";    ovBody.textContent  = "Activate Briqlab Pro to continue using this visual and unlock all features."; ovCard.appendChild(ovBody);
        const ovBtn   = document.createElement("button"); ovBtn.className   = "trial-btn";     ovBtn.textContent = getButtonText();
        ovBtn.addEventListener("click", () => this.host.launchUrl(getPurchaseUrl())); ovCard.appendChild(ovBtn);
        const ovSub   = document.createElement("p");      ovSub.className   = "trial-subtext"; ovSub.textContent   = "Purchase on Microsoft AppSource to unlock all features instantly."; ovCard.appendChild(ovSub);
        this.overlay.appendChild(ovCard);
        this.root.appendChild(this.overlay);

        // Wire input events
        this.inputEl.addEventListener("input", () => {
            if (!this.settings) return;   // guard: update() not yet called
            const q = this.inputEl.value;
            clearBtn.classList.toggle("hidden", q.length === 0);
            this.filterAndRender(q);
        });
        clearBtn.addEventListener("click", () => {
            if (!this.settings) return;   // guard
            this.inputEl.value = "";
            clearBtn.classList.add("hidden");
            this.selected.clear();
            this.selMgr.clear();
            this.filterAndRender("");
        });
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
                        dataItems: [{ displayName: "Briqlab Search", value: "" }],
                        identities: [],
                        coordinates: [e.clientX, e.clientY],
                        isTouchEvent: false
                    });
                });
                this.root.addEventListener("mouseleave", () => {
                    this.tooltipSvc.hide({ isTouchEvent: false, immediately: false });
                });
            }
            this.settings = this.fmtSvc.populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews[0]);
            this.applyStyle();
    
            // ── Parse data and render SYNCHRONOUSLY so typing always works ──────────
            // Do NOT wait for key validation — data must be available immediately.
            this.parseData(options);
            this.filterAndRender(this.inputEl.value);
    
            // ── Key validation only affects the license badge/overlay UI ────────────
            const rawKey = ""; // MS cert: pro key field removed
            const cachedRaw = (() => { try { const s = localStorage.getItem(PRO_KEY_LS); if (s) { const p = JSON.parse(s); return p.key ?? ""; } } catch { } return ""; })();
            const effectiveKey = rawKey || cachedRaw;
    
            validateKey(effectiveKey).then(valid => {
                this.isPro = valid;
                this.updateLicenseUI();
            }).catch(() => {
                this.updateLicenseUI();
            });
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    private applyStyle(): void {
        const s = this.settings;
        const accent = s.colorSettings.accentColor.value?.value ?? "#0D9488";
        const bg     = s.colorSettings.backgroundColor.value?.value ?? "#FFFFFF";
        const text   = s.colorSettings.textColor.value?.value ?? "#111827";
        const border = s.colorSettings.borderColor.value?.value ?? "#E2E8F0";
        const font   = s.searchSettings.fontFamily.value?.value ?? "Segoe UI";
        const size   = s.searchSettings.fontSize.value ?? 12;
        const ph     = s.searchSettings.placeholder.value ?? "Search\u2026";

        this.root.style.setProperty("--briq-accent", accent);
        this.root.style.setProperty("--briq-bg", bg);
        this.root.style.setProperty("--briq-text", text);
        this.root.style.setProperty("--briq-border", border);
        this.root.style.setProperty("--briq-font", String(font));
        this.root.style.setProperty("--briq-size", size + "px");
        this.inputEl.placeholder = ph;
    }

    private updateLicenseUI(): void {
        checkMicrosoftLicence(this.host).then(p => this._msUpdateLicenceUI(p)).catch(() => this._msUpdateLicenceUI(false));
    }

    private parseData(options: VisualUpdateOptions): void {
        this.items = [];
        const dv = options.dataViews?.[0];
        if (!dv?.categorical?.categories?.[0]) return;

        const cats   = dv.categorical.categories[0];
        const vals   = dv.categorical.values?.[0]?.values ?? [];

        for (let i = 0; i < cats.values.length; i++) {
            const label = String(cats.values[i] ?? "");
            const value = vals[i] != null ? Number(vals[i]) : null;
            const selId = this.host.createSelectionIdBuilder()
                .withCategory(cats, i)
                .createSelectionId();
            this.items.push({ label, value, selectionId: selId });
        }
    }

    private filterAndRender(query: string): void {
        const s = this.settings;
        const cs        = s.searchSettings.caseSensitive.value ?? false;
        const maxRes    = Math.max(1, s.searchSettings.maxResults.value ?? 10);
        const showCount = s.searchSettings.showResultCount.value ?? true;

        const q = cs ? query : query.toLowerCase();
        const filtered = q.length === 0
            ? this.items.slice(0, maxRes)
            : this.items.filter(it => {
                const lbl = cs ? it.label : it.label.toLowerCase();
                return lbl.includes(q);
              }).slice(0, maxRes);

        // Count
        if (showCount && q.length > 0) {
            const total = this.items.filter(it => {
                const lbl = cs ? it.label : it.label.toLowerCase();
                return lbl.includes(q);
            }).length;
            this.countEl.textContent = `${total} result${total !== 1 ? "s" : ""}`;
            this.countEl.classList.remove("hidden");
        } else {
            this.countEl.classList.add("hidden");
        }

        // Render list
        while (this.listEl.firstChild) this.listEl.removeChild(this.listEl.firstChild);
        if (filtered.length === 0 && q.length > 0) {
            const empty = document.createElement("div");
            empty.className = "briq-search-empty";
            empty.textContent = "No results found";
            this.listEl.appendChild(empty);
            return;
        }

        filtered.forEach(item => {
            const row = document.createElement("div");
            row.className = "briq-search-item";
            if (this.selected.has(item.label)) row.classList.add("selected");

            const lbl = document.createElement("span");
            lbl.className = "briq-search-item-label";
            // Highlight matching text
            if (q.length > 0) {
                const idx = (cs ? item.label : item.label.toLowerCase()).indexOf(q);
                if (idx >= 0) {
                    if (idx > 0) lbl.appendChild(document.createTextNode(item.label.slice(0, idx)));
                    const mark = document.createElement("mark");
                    mark.textContent = item.label.slice(idx, idx + q.length);
                    lbl.appendChild(mark);
                    if (idx + q.length < item.label.length) lbl.appendChild(document.createTextNode(item.label.slice(idx + q.length)));
                } else {
                    lbl.textContent = item.label;
                }
            } else {
                lbl.textContent = item.label;
            }
            row.appendChild(lbl);

            if (item.value != null) {
                const val = document.createElement("span");
                val.className = "briq-search-item-value";
                val.textContent = this.fmtNum(item.value);
                row.appendChild(val);
            }

            row.addEventListener("click", (e) => {
                e.stopPropagation();
                if (this.selected.has(item.label)) {
                    this.selected.delete(item.label);
                    row.classList.remove("selected");
                } else {
                    this.selected.add(item.label);
                    row.classList.add("selected");
                }
                const selectedItems = filtered.filter(it => this.selected.has(it.label));
                if (selectedItems.length === 0) {
                    this.selMgr.clear();
                } else {
                    this.selMgr.select(selectedItems.map(it => it.selectionId), e.ctrlKey || e.metaKey);
                }
            });

            this.listEl.appendChild(row);
        });
    }

    private escHtml(s: string): string {
        return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }

    private fmtNum(v: number): string {
        const a = Math.abs(v);
        if (a >= 1e9) return (v / 1e9).toFixed(1).replace(/\.0$/, "") + "B";
        if (a >= 1e6) return (v / 1e6).toFixed(1).replace(/\.0$/, "") + "M";
        if (a >= 1e3) return (v / 1e3).toFixed(1).replace(/\.0$/, "") + "K";
        return v.toLocaleString();
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
