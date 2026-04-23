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

const TRIAL_MS      = 4 * 24 * 60 * 60 * 1000;
const TRIAL_KEY     = "briqlab_trial_BulletPro_start";
const PRO_STORE_KEY = "briqlab_bulletpro_prokey";

interface BulletRow {
    label:          string;
    value:          number;
    target:         number | null;
    comparative:    number | null;
    poorThreshold:  number | null;
    satisfThreshold:number | null;
    maximum:        number | null;
    selectionId:    ISelectionId;
}

function getTrial(): { daysLeft: number; expired: boolean } {
    try {
        let raw = localStorage.getItem(TRIAL_KEY);
        if (!raw) { raw = String(Date.now()); localStorage.setItem(TRIAL_KEY, raw); }
        const elapsed  = Date.now() - parseInt(raw, 10);
        const daysLeft = Math.max(0, Math.ceil((TRIAL_MS - elapsed) / 86400000));
        return { daysLeft, expired: elapsed > TRIAL_MS };
    } catch {
        return { daysLeft: 4, expired: false };
    }
}

function numOrNull(val: powerbi.PrimitiveValue | undefined): number | null {
    if (val == null) return null;
    const n = Number(val);
    return isNaN(n) ? null : n;
}

export class Visual implements IVisual {
    private readonly host:    IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private readonly selMgr:  ISelectionManager;
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private _handlersAttached = false;
    private readonly fmtSvc:  FormattingSettingsService;

    private readonly root:        HTMLElement;
    private readonly contentEl:   HTMLDivElement;
    private readonly chartEl:     HTMLDivElement;
    private readonly trialBadge:  HTMLDivElement;
    private readonly proBadge:    HTMLDivElement;
    private readonly keyErrorEl:  HTMLDivElement;
    private readonly overlayEl:   HTMLDivElement;

    private settings!:  VisualFormattingSettingsModel;
    private vp:         powerbi.IViewport = { width: 300, height: 300 };
    private rows:       BulletRow[] = [];
    private selectedIds: Set<string> = new Set();

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
        this.root.classList.add("briqlab-bullet");
        this.root.style.position = "relative";
        this.root.style.overflow = "hidden";

        this.contentEl = this.mkDiv("briqlab-visual-content");
        this.root.appendChild(this.contentEl);

        this.chartEl = this.mkDiv("briq-bullet-container");
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

        const title = document.createElement("h2");
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
                        dataItems: [{ displayName: "Briqlab Bullet Chart", value: "" }],
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
            this.settings = this.fmtSvc.populateFormattingSettingsModel(
                VisualFormattingSettingsModel, options.dataViews?.[0]
            );
    
            this.handleProKey();
            this.updateLicenseUI();
    
            const dv: DataView | undefined = options.dataViews?.[0];
            const categorical = dv?.categorical;
            const labelColumn = categorical?.categories?.[0];
            const valCols     = categorical?.values;
    
            if (!labelColumn || !valCols || valCols.length === 0) {
                this.renderEmpty("Add Label and Value fields");
                return;
            }
    
            const findCol = (role: string) => valCols.find(c => c.source?.roles?.[role]);
            const valueCol      = findCol("value");
            const targetCol     = findCol("target");
            const comparativeCol= findCol("comparative");
            const poorCol       = findCol("poorThreshold");
            const satisfCol     = findCol("satisfThreshold");
            const maxCol        = findCol("maximum");
    
            if (!valueCol) { this.renderEmpty("Add Label and Value fields"); return; }
    
            this.rows = [];
            labelColumn.values.forEach((lv, i) => {
                const label = lv != null ? String(lv) : "(Blank)";
                const value = numOrNull(valueCol.values[i]);
                if (value === null) return;
                this.rows.push({
                    label,
                    value,
                    target:          targetCol      ? numOrNull(targetCol.values[i])      : null,
                    comparative:     comparativeCol ? numOrNull(comparativeCol.values[i]) : null,
                    poorThreshold:   poorCol        ? numOrNull(poorCol.values[i])        : null,
                    satisfThreshold: satisfCol      ? numOrNull(satisfCol.values[i])      : null,
                    maximum:         maxCol         ? numOrNull(maxCol.values[i])         : null,
                    selectionId:     this.host.createSelectionIdBuilder()
                        .withCategory(labelColumn, i).createSelectionId()
                });
            });
    
            if (this.rows.length === 0) { this.renderEmpty("No data to display"); return; }
            this.render();
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    private render(): void {
        while (this.chartEl.firstChild) this.chartEl.removeChild(this.chartEl.firstChild);

        const settings  = this.settings;
        const rows      = this.rows;
        const width     = this.vp.width;
        const height    = this.vp.height;

        const rowHeight      = settings.chartSettings.rowHeight.value;
        const bgBarHeightPct = settings.chartSettings.bgBarHeight.value / 100;
        const perfBarHPct    = settings.chartSettings.perfBarHeight.value / 100;
        const showComparative= settings.chartSettings.showComparative.value;
        const showSummary    = settings.chartSettings.showSummary.value;
        const fontFamily     = settings.chartSettings.fontFamily.value || "Segoe UI, sans-serif";
        const redZone    = settings.colorSettings.redZoneColor.value?.value    || "#FEE2E2";
        const amberZone  = settings.colorSettings.amberZoneColor.value?.value  || "#FEF3C7";
        const greenZone  = settings.colorSettings.greenZoneColor.value?.value  || "#DCFCE7";
        const targetColor= settings.colorSettings.targetColor.value?.value     || "#0F172A";
        const compColor  = settings.colorSettings.comparativeColor.value?.value|| "#64748B";

        const allMax    = rows.map(r => r.maximum ?? r.value);
        const globalMax = d3.max(allMax) ?? 1;
        const labelWidth= 120;
        const valueWidth= 50;
        const padT = showSummary ? 36 : 8;
        const padB = 30;
        const axisH = 20;
        const barAreaW = width - labelWidth - valueWidth - 8;
        const svgHeight = Math.max(height, padT + rows.length * rowHeight + axisH + padB);

        const svg = d3.select(this.chartEl)
            .append("svg")
            .attr("width", width)
            .attr("height", svgHeight)
            .style("font-family", fontFamily);

        const xScale = d3.scaleLinear().domain([0, globalMax]).range([0, barAreaW]);

        if (showSummary) {
            const onTarget = rows.filter(r => r.target !== null && r.value >= r.target).length;
            svg.append("text")
                .attr("x", width / 2).attr("y", 22)
                .attr("text-anchor", "middle")
                .attr("font-size", "12px").attr("font-weight", "600").attr("fill", "#374151")
                .text(`${onTarget} of ${rows.length} KPIs on target`);
        }

        const barsG = svg.append("g").attr("transform", `translate(${labelWidth},${padT})`);

        rows.forEach((row, i) => {
            const rowY     = i * rowHeight;
            const bgBarH   = rowHeight * bgBarHeightPct;
            const bgBarY   = (rowHeight - bgBarH) / 2;
            const perfBarH = bgBarH * perfBarHPct;
            const perfBarY = bgBarY + (bgBarH - perfBarH) / 2;
            const rowMax   = row.maximum ?? globalMax;
            const xRow     = d3.scaleLinear().domain([0, rowMax]).range([0, barAreaW]);
            const rowG     = barsG.append("g").attr("transform", `translate(0,${rowY})`);

            const poor   = row.poorThreshold   != null ? xRow(Math.min(row.poorThreshold, rowMax))   : 0;
            const satisf = row.satisfThreshold != null ? xRow(Math.min(row.satisfThreshold, rowMax)) : 0;

            if (poor > 0)         rowG.append("rect").attr("x",0).attr("y",bgBarY).attr("width",poor).attr("height",bgBarH).attr("fill",redZone).attr("rx",2);
            if (satisf > poor)    rowG.append("rect").attr("x",poor).attr("y",bgBarY).attr("width",satisf-poor).attr("height",bgBarH).attr("fill",amberZone);
            if (barAreaW > satisf)rowG.append("rect").attr("x",satisf).attr("y",bgBarY).attr("width",barAreaW-satisf).attr("height",bgBarH).attr("fill",greenZone).attr("rx",2);

            let perfColor = "#0D9488";
            if (row.poorThreshold   != null && row.value < row.poorThreshold)   perfColor = "#EF4444";
            else if (row.satisfThreshold != null && row.value < row.satisfThreshold) perfColor = "#F59E0B";

            const perfW  = xRow(Math.min(row.value, rowMax));
            const perfBar = rowG.append("rect")
                .attr("x",0).attr("y",perfBarY).attr("width",0).attr("height",perfBarH)
                .attr("fill",perfColor).attr("rx",2)
                .style("cursor","pointer");

            perfBar.transition().duration(500).delay(i * 50).attr("width", perfW);

            // Click to select
            rowG.on("click", (event: MouseEvent) => {
                const key = ((row.selectionId as unknown as Record<string,unknown>)["key"] as string) || row.label;
                const isMulti = event.ctrlKey || event.metaKey;
                if (isMulti) {
                    this.selectedIds.has(key) ? this.selectedIds.delete(key) : this.selectedIds.add(key);
                } else {
                    if (this.selectedIds.has(key) && this.selectedIds.size === 1) {
                        this.selectedIds.clear();
                    } else {
                        this.selectedIds.clear(); this.selectedIds.add(key);
                    }
                }
                const ids = this.rows.filter(r => {
                    const k = ((r.selectionId as unknown as Record<string,unknown>)["key"] as string) || r.label;
                    return this.selectedIds.has(k);
                }).map(r => r.selectionId);
                if (ids.length > 0) {
                    this.selMgr.select(ids, isMulti).then(() => this.applySelection(barsG));
                } else {
                    this.selMgr.clear().then(() => this.applySelection(barsG));
                }
                this.applySelection(barsG);
                event.stopPropagation();
            });

            if (row.target !== null) {
                const tx = xRow(Math.min(row.target, rowMax));
                rowG.append("rect").attr("x",tx-1).attr("y",bgBarY).attr("width",2).attr("height",bgBarH).attr("fill",targetColor);
            }

            if (showComparative && row.comparative !== null) {
                const cx = xRow(Math.min(row.comparative, rowMax));
                const cy = rowHeight / 2;
                const ds = 5;
                rowG.append("path")
                    .attr("d", `M${cx},${cy-ds} L${cx+ds},${cy} L${cx},${cy+ds} L${cx-ds},${cy} Z`)
                    .attr("fill","none").attr("stroke",compColor).attr("stroke-width",1.5);
            }

            if (i < rows.length - 1) {
                rowG.append("line")
                    .attr("x1",-labelWidth).attr("x2",barAreaW+valueWidth)
                    .attr("y1",rowHeight).attr("y2",rowHeight)
                    .attr("stroke","#F1F5F9").attr("stroke-width",1);
            }

            svg.append("text")
                .attr("x",labelWidth-8).attr("y",padT+rowY+rowHeight/2+4)
                .attr("text-anchor","end").attr("font-size","11px").attr("fill","#374151")
                .text(row.label);

            svg.append("text")
                .attr("x",labelWidth+barAreaW+6).attr("y",padT+rowY+rowHeight/2+4)
                .attr("text-anchor","start").attr("font-size","11px")
                .attr("font-weight","600").attr("fill",perfColor)
                .text(row.value % 1 === 0 ? String(row.value) : row.value.toFixed(1));
        });

        svg.append("g")
            .attr("transform",`translate(${labelWidth},${padT+rows.length*rowHeight})`)
            .call(d3.axisBottom(xScale).ticks(5).tickSize(4))
            .call(ax => ax.select(".domain").attr("stroke","#E5E7EB"))
            .call(ax => ax.selectAll("text").attr("font-size","10px").attr("fill","#6B7280"))
            .call(ax => ax.selectAll("line").attr("stroke","#E5E7EB"));

        svg.on("click", () => {
            this.selectedIds.clear();
            this.selMgr.clear();
            this.applySelection(barsG);
        });
    }

    private applySelection(barsG: d3.Selection<SVGGElement, unknown, null, undefined>): void {
        const hasSel = this.selectedIds.size > 0;
        barsG.selectAll<SVGGElement, unknown>("g").style("opacity", (_d, i) => {
            if (!hasSel) return null;
            const row = this.rows[i];
            if (!row) return null;
            const key = ((row.selectionId as unknown as Record<string,unknown>)["key"] as string) || row.label;
            return this.selectedIds.has(key) ? "1" : "0.4";
        });
    }

    private renderEmpty(msg: string): void {
        while (this.chartEl.firstChild) this.chartEl.removeChild(this.chartEl.firstChild);
        const el = this.mkDiv("briq-empty");
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
            const parsed = JSON.parse(stored) as { key?: string };
            const key    = parsed?.key;
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
        const el = document.createElement("div");
        el.className = cls;
        return el;
    }
}
