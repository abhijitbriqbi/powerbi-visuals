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

const CHART_COLORS = [
    "#0D9488","#F97316","#3B82F6","#8B5CF6","#10B981",
    "#EF4444","#F59E0B","#EC4899","#06B6D4","#84CC16"
];

const TRIAL_MS  = 4 * 24 * 60 * 60 * 1000;
const TRIAL_KEY = "briqlab_trial_BriqlabRadarPro_start";
const CACHED_KEY = "briqlab_radarpro_prokey";

interface RadarDataPoint {
    entity: string;
    dimension: string;
    value: number;
}

export class Visual implements IVisual {
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private selMgr!:      powerbi.extensibility.ISelectionManager;
    private _handlersAttached = false;
    private readonly host: IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private readonly fmtSvc: FormattingSettingsService;
    private settings!: VisualFormattingSettingsModel;

    private root!:      HTMLElement;
    private contentEl!: HTMLElement;
    private svgEl!:     d3.Selection<SVGSVGElement, unknown, null, undefined>;
    private trialBadge!: HTMLElement;
    private proBadge!:   HTMLElement;
    private overlay!:    HTMLElement;

    private vp: powerbi.IViewport = { width: 400, height: 300 };
    private curKey = "";
    private isPro  = false;
    private keyCache: Map<string, boolean> = new Map();
    private trialStart = 0;
    private dataPoints: RadarDataPoint[] = [];
    private benchmarkMap: Map<string, number> = new Map();

    constructor(options: VisualConstructorOptions) {
        this.host   = options.host;
        this.renderingManager = options.host.eventService;
        this.selMgr     = this.host.createSelectionManager();
        this.tooltipSvc = this.host.tooltipService;
        this.fmtSvc = new FormattingSettingsService();
        this.root   = options.element;
        this.root.classList.add("briqlab-radar");
        this.buildDOM();
        this.initTrial();
    }

    private buildDOM(): void {
        this.contentEl = document.createElement("div");
        this.contentEl.className = "radar-content";
        this.root.appendChild(this.contentEl);

        this.svgEl = d3.select(this.contentEl)
            .append<SVGSVGElement>("svg")
            .attr("class", "radar-svg");

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
                        dataItems: [{ displayName: "Briqlab Radar Chart", value: "" }],
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
    
            this.dataPoints = [];
            this.benchmarkMap = new Map();
    
            const dv = options.dataViews?.[0];
            if (dv?.categorical) {
                const cats = dv.categorical.categories ?? [];
                const vals = dv.categorical.values    ?? [];
                const entityCol    = cats.find(c => (c.source.roles as Record<string, unknown>)["entity"]);
                const dimensionCol = cats.find(c => (c.source.roles as Record<string, unknown>)["dimension"]);
                const valueCol     = vals.find(c => (c.source.roles as Record<string, unknown>)["value"]);
                const benchmarkCol = vals.find(c => (c.source.roles as Record<string, unknown>)["benchmark"]);
    
                if (entityCol && dimensionCol && valueCol) {
                    for (let i = 0; i < entityCol.values.length; i++) {
                        const v = Number(valueCol.values[i]);
                        if (!isNaN(v)) {
                            this.dataPoints.push({
                                entity:    String(entityCol.values[i]    ?? ""),
                                dimension: String(dimensionCol.values[i] ?? ""),
                                value: v
                            });
                        }
                        if (benchmarkCol) {
                            const dim = String(dimensionCol.values[i] ?? "");
                            const bv  = Number(benchmarkCol.values[i]);
                            if (!isNaN(bv)) this.benchmarkMap.set(dim, bv);
                        }
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

        const rs = this.settings.radarSettings;
        const ls = this.settings.labelSettings;
        const { width, height } = this.vp;

        const gridRings    = Math.min(8, Math.max(3, rs.gridRings.value ?? 5));
        const maxEntities  = Math.min(5, Math.max(2, rs.maxEntities.value ?? 5));
        const fillOpacity  = Math.min(100, Math.max(0, rs.fillOpacity.value ?? 12)) / 100;
        const showDots     = rs.showDots.value;
        const showBenchmark = rs.showBenchmark.value;
        const axisFontSize = ls.axisFontSize.value ?? 11;
        const fontFamily   = String(ls.fontFamily?.value?.value ?? "Segoe UI");

        this.svgEl.attr("width", width).attr("height", height);
        this.svgEl.selectAll("*").remove();

        if (this.dataPoints.length === 0) {
            this.svgEl.append("text")
                .attr("x", width / 2).attr("y", height / 2)
                .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                .attr("class", "radar-empty").attr("font-family", fontFamily)
                .text("Add Entity, Dimension & Value fields");
            return;
        }

        const entities   = Array.from(new Set(this.dataPoints.map(d => d.entity))).slice(0, maxEntities);
        const dimensions = Array.from(new Set(this.dataPoints.map(d => d.dimension)));
        const N          = dimensions.length;
        if (N < 3) {
            this.svgEl.append("text")
                .attr("x", width / 2).attr("y", height / 2)
                .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                .attr("class", "radar-empty").attr("font-family", fontFamily)
                .text("Add at least 3 dimensions");
            return;
        }

        // Legend space
        const legendH = entities.length > 0 ? 24 : 0;
        const padH    = 16 + axisFontSize + 12;
        const cx      = width / 2;
        const cy      = (height - legendH) / 2;
        const radius  = Math.min(cx, cy) - padH;

        // Gather max value per dimension for normalizing
        const dimMax = new Map<string, number>();
        for (const dp of this.dataPoints) {
            const cur = dimMax.get(dp.dimension) ?? 0;
            if (dp.value > cur) dimMax.set(dp.dimension, dp.value);
        }
        const globalMax = Math.max(...Array.from(dimMax.values()), 1);

        const angle = (i: number) => (Math.PI * 2 * i) / N - Math.PI / 2;
        const px    = (i: number, r: number) => cx + r * Math.cos(angle(i));
        const py    = (i: number, r: number) => cy + r * Math.sin(angle(i));

        const g = this.svgEl.append("g").attr("class", "radar-g");

        // Grid rings
        for (let ring = 1; ring <= gridRings; ring++) {
            const r = (ring / gridRings) * radius;
            const points = dimensions.map((_, i) => `${px(i, r)},${py(i, r)}`).join(" ");
            g.append("polygon")
                .attr("points", points)
                .attr("class", "radar-grid-ring");
        }

        // Axis spokes
        for (let i = 0; i < N; i++) {
            g.append("line")
                .attr("x1", cx).attr("y1", cy)
                .attr("x2", px(i, radius)).attr("y2", py(i, radius))
                .attr("class", "radar-spoke");

            // Axis labels
            const lx = px(i, radius + 12);
            const ly = py(i, radius + 12);
            g.append("text")
                .attr("x", lx).attr("y", ly)
                .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                .attr("class", "radar-axis-label")
                .attr("font-size", axisFontSize)
                .attr("font-family", fontFamily)
                .text(dimensions[i]);
        }

        // Build entity data map
        const entityMap = new Map<string, Map<string, number>>();
        for (const dp of this.dataPoints) {
            if (!entities.includes(dp.entity)) continue;
            let m = entityMap.get(dp.entity);
            if (!m) { m = new Map(); entityMap.set(dp.entity, m); }
            m.set(dp.dimension, dp.value);
        }

        // Draw entity polygons
        entities.forEach((entity, ei) => {
            const color   = CHART_COLORS[ei % CHART_COLORS.length];
            const dimData = entityMap.get(entity) ?? new Map<string, number>();
            const pts     = dimensions.map((dim, i) => {
                const v = dimData.get(dim) ?? 0;
                const r = (v / globalMax) * radius;
                return [px(i, r), py(i, r)] as [number, number];
            });

            const pointsStr = pts.map(p => `${p[0]},${p[1]}`).join(" ");

            g.append("polygon")
                .attr("points", pointsStr)
                .attr("fill", color)
                .attr("fill-opacity", fillOpacity)
                .attr("stroke", color)
                .attr("stroke-opacity", 0.85)
                .attr("stroke-width", 2)
                .attr("class", "radar-entity");

            if (showDots) {
                pts.forEach(p => {
                    g.append("circle")
                        .attr("cx", p[0]).attr("cy", p[1])
                        .attr("r", 3.5)
                        .attr("fill", color)
                        .attr("stroke", "#fff")
                        .attr("stroke-width", 1.5)
                        .attr("class", "radar-dot");
                });
            }
        });

        // Benchmark polygon
        if (showBenchmark && this.benchmarkMap.size > 0) {
            const bmPts = dimensions.map((dim, i) => {
                const bv = this.benchmarkMap.get(dim) ?? 0;
                const r  = (bv / globalMax) * radius;
                return `${px(i, r)},${py(i, r)}`;
            }).join(" ");

            g.append("polygon")
                .attr("points", bmPts)
                .attr("fill", "none")
                .attr("stroke", "#ffffff")
                .attr("stroke-width", 1.5)
                .attr("stroke-dasharray", "5,3")
                .attr("class", "radar-benchmark");
        }

        // Legend
        if (entities.length > 0) {
            const legendY  = height - legendH + 8;
            const totalW   = entities.length * 100;
            let legendX    = Math.max(8, (width - totalW) / 2);
            entities.forEach((entity, ei) => {
                const color = CHART_COLORS[ei % CHART_COLORS.length];
                this.svgEl.append("circle")
                    .attr("cx", legendX + 5).attr("cy", legendY)
                    .attr("r", 5).attr("fill", color);
                this.svgEl.append("text")
                    .attr("x", legendX + 14).attr("y", legendY)
                    .attr("dominant-baseline", "middle")
                    .attr("class", "radar-legend-label")
                    .attr("font-family", fontFamily)
                    .text(entity);
                legendX += 100;
            });
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
        return this.fmtSvc.buildFormattingModel(this.settings);
    }
}
