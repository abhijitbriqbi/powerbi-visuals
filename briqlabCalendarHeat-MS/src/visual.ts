"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import * as d3 from "d3";
import "./../style/visual.less";

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

const TRIAL_MS      = 4 * 24 * 60 * 60 * 1000;
const TRIAL_KEY     = "briqlab_trial_CalendarHeat_start";
const PRO_STORE_KEY = "briqlab_calendarheat_prokey";

const SCALE_PRESETS: Record<string, { low: string; high: string }> = {
    Teal:   { low: "#CCFBF1", high: "#0D9488" },
    Blue:   { low: "#DBEAFE", high: "#1D4ED8" },
    Green:  { low: "#DCFCE7", high: "#15803D" },
    Purple: { low: "#EDE9FE", high: "#7C3AED" },
    Orange: { low: "#FFF7ED", high: "#EA580C" },
    Custom: { low: "#CCFBF1", high: "#0D9488" }
};

const MONTH_NAMES   = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const DAY_LABELS_MON = ["M","","W","","F","",""];
const DAY_LABELS_SUN = ["","M","","W","","F",""];

function getTrial(): { daysLeft: number; expired: boolean } {
    try {
        let raw = localStorage.getItem(TRIAL_KEY);
        if (!raw) { raw = String(Date.now()); localStorage.setItem(TRIAL_KEY, raw); }
        const elapsed  = Date.now() - parseInt(raw, 10);
        const daysLeft = Math.max(0, Math.ceil((TRIAL_MS - elapsed) / 86400000));
        return { daysLeft, expired: elapsed > TRIAL_MS };
    } catch { return { daysLeft: 4, expired: false }; }
}

function toISODate(d: Date): string {
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}

interface CalCell {
    iso:         string;
    selectionId: ISelectionId;
    value:       number;
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

    private settings!:    VisualFormattingSettingsModel;
    private vp:           powerbi.IViewport = { width: 400, height: 300 };
    private cells:        CalCell[] = [];
    private selectedIsos: Set<string> = new Set();

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
        this.root.classList.add("briqlab-calendarheat");

        this.contentEl = this.mkDiv("briqlab-visual-content");
        this.root.appendChild(this.contentEl);

        this.chartEl = this.mkDiv("briq-calendarheat-root");
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
                        dataItems: [{ displayName: "Briqlab Calendar Heatmap", value: "" }],
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
            if (!dv?.categorical?.categories?.length) { this.renderEmpty("Add Date and Value fields"); return; }
    
            const catData     = dv.categorical!;
            const dateCategory= catData.categories![0];
            const valueCol    = catData.values?.[0] ?? null;
    
            if (!valueCol) { this.renderEmpty("Add Date and Value fields"); return; }
    
            const dataMap = new Map<string, { value: number; selId: ISelectionId }>();
            for (let i = 0; i < dateCategory.values.length; i++) {
                const rawDate = dateCategory.values[i];
                if (rawDate == null) continue;
                const d = rawDate instanceof Date ? rawDate : new Date(String(rawDate));
                if (isNaN(d.getTime())) continue;
                const iso = toISODate(d);
                dataMap.set(iso, {
                    value: Number(valueCol.values[i] ?? 0),
                    selId: this.host.createSelectionIdBuilder().withCategory(dateCategory, i).createSelectionId()
                });
            }
    
            if (dataMap.size === 0) { this.renderEmpty("No date data to display"); return; }
    
            this.cells = Array.from(dataMap.entries()).map(([iso, { value, selId }]) => ({
                iso, value, selectionId: selId
            }));
    
            this.renderChart(dataMap);
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    private renderChart(dataMap: Map<string, { value: number; selId: ISelectionId }>): void {
        while (this.chartEl.firstChild) this.chartEl.removeChild(this.chartEl.firstChild);

        const width  = this.vp.width;
        const height = this.vp.height;
        const layout = this.settings.layoutSettings;
        const colors = this.settings.colorSettings;

        this.chartEl.style.width    = `${width}px`;
        this.chartEl.style.height   = `${height}px`;
        this.chartEl.style.position = "relative";

        const allValues = Array.from(dataMap.values()).map(v => v.value);
        const minVal    = d3.min(allValues) ?? 0;
        const maxVal    = d3.max(allValues) ?? 1;
        const avgVal    = d3.mean(allValues) ?? 0;

        const allDates = Array.from(dataMap.keys()).map(s => new Date(s)).filter(d => !isNaN(d.getTime()));
        allDates.sort((a,b) => a.getTime() - b.getTime());
        const startYear = allDates[0].getFullYear();
        const endYear   = allDates[allDates.length-1].getFullYear();

        const startOfWeek  = String(layout.startOfWeek.value?.value ?? "Mon");
        const mondayFirst  = startOfWeek !== "Sun";
        const showMonthLbl = layout.showMonthLabels.value ?? true;
        const showDayLbl   = layout.showDayLabels.value ?? true;
        const showSummary  = layout.showSummaryStats.value ?? true;
        const highlightWE  = layout.highlightWeekends.value ?? true;
        const fontFamily   = String(layout.fontFamily.value?.value ?? "Segoe UI");

        const colorScaleName = String(colors.colorScale.value?.value ?? "Teal");
        const steps     = Math.max(2, colors.colorSteps.value ?? 7);
        const nullColor = colors.nullColor.value?.value ?? "#F1F5F9";
        let lowColor: string, highColor: string;
        if (colorScaleName === "Custom") {
            lowColor  = colors.lowColor.value?.value  ?? "#CCFBF1";
            highColor = colors.highColor.value?.value ?? "#0D9488";
        } else {
            const preset = SCALE_PRESETS[colorScaleName] ?? SCALE_PRESETS["Teal"];
            lowColor = preset.low; highColor = preset.high;
        }

        const summaryH   = showSummary ? 36 : 0;
        const monthLabelH= showMonthLbl ? 18 : 0;
        const dayLabelW  = showDayLbl ? 18 : 0;
        const years      = endYear - startYear + 1;
        const cellSize   = Math.max(6, Math.floor((width - dayLabelW - 20) / 54));
        const gap        = 2;
        const yearHeight = monthLabelH + 7*(cellSize+gap) + 16;
        const totalSvgH  = years * yearHeight + summaryH + 8;

        const visual = this.mkDiv("briq-visual-content");
        visual.style.cssText = "position:absolute;top:0;left:0;overflow-y:auto;";
        visual.style.width   = `${width}px`;
        visual.style.height  = `${height}px`;
        this.chartEl.appendChild(visual);

        const svg = d3.select(visual).append("svg")
            .attr("width", width).attr("height", Math.max(totalSvgH, height));

        const colorInterp = d3.interpolate(lowColor, highColor);
        const colorScale  = (v: number): string => {
            if (maxVal === minVal) return highColor;
            return colorInterp((v - minVal) / (maxVal - minVal));
        };

        const tooltip = this.mkDiv("briq-tooltip");
        tooltip.style.display = "none";
        visual.appendChild(tooltip);

        // Inject keyframe animation
        const styleId = "briq-calendarheat-keyframes";
        if (!document.getElementById(styleId)) {
            const s = document.createElement("style");
            s.id = styleId;
            s.textContent = "@keyframes calFadeIn { 0% { opacity:0; } 100% { opacity:1; } }";
            document.head.appendChild(s);
        }

        for (let yr = startYear; yr <= endYear; yr++) {
            const yearOffY = (yr - startYear) * yearHeight;

            svg.append("text")
                .attr("x", dayLabelW+2).attr("y", yearOffY+monthLabelH-4)
                .attr("font-family", fontFamily).attr("font-size", 11)
                .attr("fill", "#64748B").attr("font-weight", "600").text(String(yr));

            const jan1    = new Date(yr, 0, 1);
            const jan1Day = jan1.getDay();
            let gridStart = new Date(jan1);
            if (mondayFirst) {
                const shift = jan1Day === 0 ? 6 : jan1Day - 1;
                gridStart   = new Date(jan1.getTime() - shift*86400000);
            } else {
                gridStart   = new Date(jan1.getTime() - jan1Day*86400000);
            }

            if (showDayLbl) {
                const labels = mondayFirst ? DAY_LABELS_MON : DAY_LABELS_SUN;
                for (let di = 0; di < 7; di++) {
                    if (labels[di]) {
                        svg.append("text")
                            .attr("x", dayLabelW-2).attr("y", yearOffY+monthLabelH+di*(cellSize+gap)+cellSize/2+4)
                            .attr("text-anchor","end").attr("font-family",fontFamily)
                            .attr("font-size",9).attr("fill","#94A3B8").text(labels[di]);
                    }
                }
            }

            let weekCount = 0;
            const monthLblPlaced = new Set<number>();
            let weekDate = new Date(gridStart);

            while (weekDate.getFullYear() <= yr && weekCount < 54) {
                const weekX = dayLabelW + weekCount*(cellSize+gap);

                for (let di = 0; di < 7; di++) {
                    const cellDate = new Date(weekDate.getTime() + di*86400000);
                    if (cellDate.getFullYear() === yr && cellDate.getDate() === 1 && showMonthLbl) {
                        const mon = cellDate.getMonth();
                        if (!monthLblPlaced.has(mon)) {
                            monthLblPlaced.add(mon);
                            svg.append("text")
                                .attr("x",weekX).attr("y",yearOffY+monthLabelH-4)
                                .attr("font-family",fontFamily).attr("font-size",10).attr("fill","#64748B")
                                .text(MONTH_NAMES[mon]);
                        }
                    }
                }

                if (highlightWE) {
                    const colDay = new Date(weekDate).getDay();
                    if (colDay === 0 || colDay === 6) {
                        svg.append("rect")
                            .attr("x",weekX-1).attr("y",yearOffY+monthLabelH)
                            .attr("width",cellSize+2).attr("height",7*(cellSize+gap))
                            .attr("fill","#F8FAFC").attr("rx",2);
                    }
                }

                for (let di = 0; di < 7; di++) {
                    const cellDate = new Date(weekDate.getTime() + di*86400000);
                    const cellY    = yearOffY + monthLabelH + di*(cellSize+gap);
                    const iso      = toISODate(cellDate);
                    const inYear   = cellDate.getFullYear() === yr;
                    const entry    = dataMap.get(iso);
                    const fill     = (inYear && entry !== undefined) ? colorScale(entry.value) : nullColor;

                    const rect = svg.append("rect")
                        .attr("x",weekX).attr("y",cellY)
                        .attr("width",cellSize).attr("height",cellSize)
                        .attr("rx",2).attr("fill",fill)
                        .style("cursor","pointer")
                        .style("animation",`calFadeIn 0.3s ease-out ${weekCount*15}ms both`);

                    if (inYear && entry !== undefined) {
                        const dateLabel = cellDate.toLocaleDateString("en-GB",{day:"numeric",month:"short",year:"numeric"});
                        const valLabel  = entry.value.toLocaleString();
                        const avgPct    = avgVal > 0 ? ` (${entry.value >= avgVal ? "+" : ""}${Math.round(((entry.value-avgVal)/avgVal)*100)}% vs avg)` : "";
                        const capturedSelId = entry.selId;
                        const capturedIso   = iso;

                        rect.on("mouseover", function(event: MouseEvent) {
                            d3.select(this).attr("stroke","#0D9488").attr("stroke-width",1.5);
                            tooltip.style.display = "block";
                            while (tooltip.firstChild) tooltip.removeChild(tooltip.firstChild);
                            const l1 = document.createElement("div"); l1.style.fontWeight = "600"; l1.textContent = dateLabel; tooltip.appendChild(l1);
                            const l2 = document.createElement("div"); l2.textContent = `Value: ${valLabel}${avgPct}`; tooltip.appendChild(l2);
                            tooltip.style.left = `${(event as MouseEvent).offsetX+10}px`;
                            tooltip.style.top  = `${(event as MouseEvent).offsetY-10}px`;
                        })
                        .on("mousemove", function(event: MouseEvent) {
                            tooltip.style.left = `${(event as MouseEvent).offsetX+10}px`;
                            tooltip.style.top  = `${(event as MouseEvent).offsetY-10}px`;
                        })
                        .on("mouseout", function() {
                            d3.select(this).attr("stroke","none");
                            tooltip.style.display = "none";
                        })
                        .on("click", (event: MouseEvent) => {
                            const isMulti = event.ctrlKey || event.metaKey;
                            if (this.selectedIsos.has(capturedIso) && !isMulti) {
                                this.selectedIsos.clear();
                                this.selMgr.clear();
                            } else {
                                if (!isMulti) this.selectedIsos.clear();
                                this.selectedIsos.add(capturedIso);
                                const selIds = this.cells
                                    .filter(c => this.selectedIsos.has(c.iso))
                                    .map(c => c.selectionId);
                                this.selMgr.select(selIds, isMulti);
                            }
                            event.stopPropagation();
                        });

                        void capturedSelId;
                    }
                }

                weekDate = new Date(weekDate.getTime() + 7*86400000);
                weekCount++;
                if (weekDate.getFullYear() > yr+1) break;
            }
        }

        if (showSummary && allDates.length > 0) {
            const allEntries = Array.from(dataMap.entries());
            const bestEntry  = allEntries.reduce((b,c) => c[1].value > b[1].value ? c : b);
            const worstEntry = allEntries.reduce((w,c) => c[1].value < w[1].value ? c : w);
            const totalVal   = allValues.reduce((a,b) => a+b, 0);

            const statsDiv = this.mkDiv("briq-summary-stats");
            statsDiv.style.fontFamily = fontFamily;
            const items = [
                { label:"Best Day",  val:`${bestEntry[0]}: ${bestEntry[1].value.toLocaleString()}` },
                { label:"Worst Day", val:`${worstEntry[0]}: ${worstEntry[1].value.toLocaleString()}` },
                { label:"Daily Avg", val:avgVal.toFixed(1) },
                { label:"Total",     val:totalVal.toLocaleString() }
            ];
            items.forEach(item => {
                const span = document.createElement("span"); span.className = "briq-stat-item";
                const lbl  = document.createElement("strong"); lbl.textContent = item.label+": "; span.appendChild(lbl);
                const val  = document.createElement("span"); val.textContent = item.val; span.appendChild(val);
                statsDiv.appendChild(span);
            });
            visual.appendChild(statsDiv);
        }

        const legendDiv = this.mkDiv("briq-color-legend");
        legendDiv.style.fontFamily = fontFamily;
        const lessLbl = document.createElement("span"); lessLbl.textContent = "Less"; legendDiv.appendChild(lessLbl);
        for (let s = 0; s < steps; s++) {
            const sq = document.createElement("span"); sq.className = "briq-legend-sq";
            sq.style.backgroundColor = colorInterp(s / (steps-1));
            legendDiv.appendChild(sq);
        }
        const moreLbl = document.createElement("span"); moreLbl.textContent = "More"; legendDiv.appendChild(moreLbl);
        visual.appendChild(legendDiv);
    }

    private renderEmpty(msg: string): void {
        while (this.chartEl.firstChild) this.chartEl.removeChild(this.chartEl.firstChild);
        const el = this.mkDiv("briq-empty-state");
        el.textContent = msg;
        this.chartEl.appendChild(el);
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
