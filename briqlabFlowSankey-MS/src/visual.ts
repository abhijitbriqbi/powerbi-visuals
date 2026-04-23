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
const TRIAL_KEY = "briqlab_trial_BriqlabFlowSankey_start";
const PRO_KEY   = "briqlab_flowsankey_prokey";

interface SankeyNode { id: string; color: string; x: number; y: number; h: number; totalIn: number; totalOut: number; }
interface SankeyLink { source: string; target: string; value: number; pct: number; srcColor: string; dstColor: string; srcY: number; dstY: number; }

function fmtNum(v: number): string {
    if (Math.abs(v) >= 1e6) return `${(v/1e6).toFixed(1)}M`;
    if (Math.abs(v) >= 1e3) return `${(v/1e3).toFixed(1)}K`;
    return v.toLocaleString();
}

export class Visual implements IVisual {
    private tooltipSvc!:  powerbi.extensibility.ITooltipService;
    private selMgr!: powerbi.extensibility.ISelectionManager;
    private _handlersAttached = false;
    private readonly host:    IVisualHost;
    private renderingManager!: powerbi.extensibility.IVisualEventService;
    private readonly fmtSvc: FormattingSettingsService;
    private settings!: VisualFormattingSettingsModel;

    private root!:     HTMLElement;
    private contentEl!: HTMLElement;
    private svgEl!: d3.Selection<SVGSVGElement, unknown, null, undefined>;
    private trialBadge!: HTMLElement;
    private proBadge!:   HTMLElement;
    private keyError!:   HTMLElement;
    private overlay!:    HTMLElement;

    private vp: powerbi.IViewport = { width: 400, height: 300 };
    private curKey = ""; private isPro = false;
    private keyCache: Map<string, boolean> = new Map();
    private trialStart = 0;

    private flows: { source: string; destination: string; value: number }[] = [];

    constructor(options: VisualConstructorOptions) {
        this.host   = options.host;
        this.renderingManager = options.host.eventService;
        this.selMgr       = this.host.createSelectionManager();
        this.tooltipSvc   = this.host.tooltipService;
        this.fmtSvc = new FormattingSettingsService();
        this.root   = options.element;
        this.root.classList.add("briqlab-sankey");
        this.buildDOM();
        this.initTrial();
    }

    private buildDOM(): void {
        this.contentEl = document.createElement("div");
        this.contentEl.className = "sankey-content";
        this.root.appendChild(this.contentEl);

        this.svgEl = d3.select(this.contentEl).append<SVGSVGElement>("svg").attr("class","sankey-svg");

        this.trialBadge = document.createElement("div"); this.trialBadge.className = "briq-trial-badge hidden"; this.root.appendChild(this.trialBadge);
        this.proBadge   = document.createElement("div"); this.proBadge.className   = "briq-pro-badge hidden";   this.proBadge.textContent = "✓ Pro Active"; this.root.appendChild(this.proBadge);
        this.keyError   = document.createElement("div"); this.keyError.className   = "briq-key-error hidden";   this.keyError.textContent = "✗ Invalid key"; this.root.appendChild(this.keyError);
        this.overlay    = document.createElement("div"); this.overlay.className    = "briq-trial-overlay hidden"; this.root.appendChild(this.overlay);
        this.buildOverlay();
    }

    private buildOverlay(): void {
        const card = document.createElement("div"); card.className = "briq-trial-card";
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
            const s = localStorage.getItem(TRIAL_KEY);
            this.trialStart = s ? parseInt(s,10) : Date.now();
            if (!s) localStorage.setItem(TRIAL_KEY, String(this.trialStart));
            const k = localStorage.getItem(PRO_KEY);
            if (k) { this.curKey = k; this.validateKey(k).then(ok => { this.isPro = ok; this.render(); }); }
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
                        dataItems: [{ displayName: "Briqlab Flow Sankey", value: "" }],
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
    
            this.flows = [];
            const dv = options.dataViews?.[0];
            if (dv?.categorical) {
                const cats = dv.categorical.categories ?? [];
                const vals = dv.categorical.values ?? [];
                const srcCol  = cats.find(c => (c.source.roles as Record<string,unknown>)["source"]);
                const dstCol  = cats.find(c => (c.source.roles as Record<string,unknown>)["destination"]);
                const valCol  = vals.find(c => (c.source.roles as Record<string,unknown>)["value"]);
                if (srcCol && dstCol && valCol) {
                    for (let i = 0; i < srcCol.values.length; i++) {
                        const v = Number(valCol.values[i]);
                        if (!isNaN(v) && v > 0) {
                            this.flows.push({
                                source:      String(srcCol.values[i] ?? ""),
                                destination: String(dstCol.values[i] ?? ""),
                                value: v
                            });
                        }
                    }
                }
            }
    
            const key = ""; // MS cert: pro key field removed
            if (key && key !== this.curKey) {
                this.curKey = key; this.isPro = false;
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
        const s = this.settings;
        const { width, height } = this.vp;

        const expired = this.isExpired();
        if (!this.isPro && !expired) {
            this.trialBadge.textContent = `Trial: ${this.daysLeft()} days remaining`;
            this.trialBadge.classList.remove("hidden");
        } else { this.trialBadge.classList.add("hidden"); }
        if (this.isPro) { this.proBadge.classList.remove("hidden"); }
        else            { this.proBadge.classList.add("hidden");    }
        if (expired && !this.isPro) {
            this.contentEl.classList.add("blurred"); this.overlay.classList.remove("hidden");
        } else {
            this.contentEl.classList.remove("blurred"); this.overlay.classList.add("hidden");
        }

        const nodeW      = Math.max(8, Math.min(40, s.layoutSettings.nodeWidth.value ?? 20));
        const nodeGap    = Math.max(2, s.layoutSettings.nodeGap.value ?? 12);
        const opacity    = Math.min(0.9, Math.max(0.1, (s.layoutSettings.flowOpacity.value ?? 60) / 100));
        const showLabels = s.layoutSettings.showFlowLabels.value;
        const minPct     = s.layoutSettings.minLabelPct.value ?? 5;
        const fontFam    = String(s.layoutSettings.fontFamily?.value?.value ?? "Segoe UI");

        this.svgEl.attr("width", width).attr("height", height);
        this.svgEl.selectAll("*").remove();

        // ── Empty state ────────────────────────────────────────────────────────
        if (this.flows.length === 0) {
            this.svgEl.append("text")
                .attr("x", width/2).attr("y", height/2)
                .attr("text-anchor","middle").attr("dominant-baseline","middle")
                .attr("class","sankey-empty").attr("font-family", fontFam)
                .text("Add Source, Destination & Value fields");
            return;
        }

        const pad = { t: 24, r: 120, b: 24, l: 120 };
        const innerW = width  - pad.l - pad.r;
        const innerH = height - pad.t - pad.b;

        // ── Build node maps ────────────────────────────────────────────────────
        const srcNames = Array.from(new Set(this.flows.map(f => f.source)));
        const dstNames = Array.from(new Set(this.flows.map(f => f.destination)));
        const allNames = Array.from(new Set([...srcNames, ...dstNames]));
        const colorMap = new Map<string, string>();
        allNames.forEach((n, i) => colorMap.set(n, CHART_COLORS[i % CHART_COLORS.length]));

        const total = this.flows.reduce((a, f) => a + f.value, 0);

        // Compute node totals
        const nodeTotal = new Map<string, number>();
        for (const f of this.flows) {
            nodeTotal.set(f.source,      (nodeTotal.get(f.source)      ?? 0) + f.value);
            nodeTotal.set(f.destination, (nodeTotal.get(f.destination) ?? 0) + f.value);
        }

        // Layout nodes
        const layoutNodes = (names: string[], xPos: number): SankeyNode[] => {
            const totalH = innerH - (names.length - 1) * nodeGap;
            let y = pad.t;
            return names.map(name => {
                const tot = nodeTotal.get(name) ?? 0;
                const h   = names.length === 1 ? totalH : Math.max(8, (tot / total) * totalH);
                const node: SankeyNode = { id: name, color: colorMap.get(name) ?? "#0D9488", x: xPos, y, h, totalIn: 0, totalOut: 0 };
                y += h + nodeGap;
                return node;
            });
        };

        const srcNodes = layoutNodes(srcNames, pad.l);
        const dstNodes = layoutNodes(dstNames, pad.l + innerW - nodeW);
        const nodeMap  = new Map<string, SankeyNode>();
        [...srcNodes, ...dstNodes].forEach(n => nodeMap.set(n.id, n));

        // Track vertical offsets for flow stacking
        const srcUsed = new Map<string, number>();
        const dstUsed = new Map<string, number>();
        srcNodes.forEach(n => srcUsed.set(n.id, n.y));
        dstNodes.forEach(n => dstUsed.set(n.id, n.y));

        // Build links
        const links: SankeyLink[] = this.flows.map(f => {
            const src = nodeMap.get(f.source)!;
            const dst = nodeMap.get(f.destination)!;
            const srcH = (f.value / (nodeTotal.get(f.source) ?? f.value)) * (src?.h ?? 0);
            const dstH = (f.value / (nodeTotal.get(f.destination) ?? f.value)) * (dst?.h ?? 0);
            const sy = srcUsed.get(f.source) ?? 0;
            const dy = dstUsed.get(f.destination) ?? 0;
            srcUsed.set(f.source,      sy + srcH);
            dstUsed.set(f.destination, dy + dstH);
            return {
                source: f.source, target: f.destination, value: f.value,
                pct: total ? (f.value / total) * 100 : 0,
                srcColor: colorMap.get(f.source)      ?? "#0D9488",
                dstColor: colorMap.get(f.destination) ?? "#F97316",
                srcY: sy + srcH / 2,
                dstY: dy + dstH / 2,
            };
        });

        const defs = this.svgEl.append("defs");

        // ── Draw links ─────────────────────────────────────────────────────────
        const linksG = this.svgEl.append("g").attr("class","sankey-links");

        links.forEach((lk, i) => {
            const src = nodeMap.get(lk.source);
            const dst = nodeMap.get(lk.target);
            if (!src || !dst) return;

            const gradId = `sg-${i}`;
            const grad = defs.append("linearGradient").attr("id", gradId)
                .attr("x1","0%").attr("y1","0%").attr("x2","100%").attr("y2","0%");
            grad.append("stop").attr("offset","0%").attr("stop-color", lk.srcColor).attr("stop-opacity", opacity);
            grad.append("stop").attr("offset","100%").attr("stop-color", lk.dstColor).attr("stop-opacity", opacity);

            const flowVal = (lk.value / (nodeTotal.get(lk.source) ?? lk.value));
            const strokeW = Math.max(2, flowVal * (src.h ?? 0));
            const x1 = src.x + nodeW;
            const x2 = dst.x;
            const cpX = (x1 + x2) / 2;

            const path = linksG.append("path")
                .attr("d", `M${x1},${lk.srcY} C${cpX},${lk.srcY} ${cpX},${lk.dstY} ${x2},${lk.dstY}`)
                .attr("fill", "none")
                .attr("stroke", `url(#${gradId})`)
                .attr("stroke-width", strokeW)
                .attr("class","sankey-link");

            // Flow label
            if (showLabels && lk.pct >= minPct) {
                const midX = (x1 + x2) / 2;
                const midY = (lk.srcY + lk.dstY) / 2;
                this.svgEl.append("text")
                    .attr("x", midX).attr("y", midY)
                    .attr("text-anchor","middle").attr("dominant-baseline","middle")
                    .attr("class","sankey-flow-label")
                    .attr("font-family", fontFam)
                    .text(`${lk.pct.toFixed(1)}%`);
            }

            // Hover tooltip
            path.append("title").text(`${lk.source} → ${lk.target}: ${fmtNum(lk.value)} (${lk.pct.toFixed(1)}%)`);

            path.on("mouseover", function() {
                linksG.selectAll(".sankey-link").style("opacity", 0.15);
                d3.select(this).style("opacity", 1);
            }).on("mouseout", function() {
                linksG.selectAll(".sankey-link").style("opacity", null);
            });
        });

        // ── Draw nodes ─────────────────────────────────────────────────────────
        const drawNodes = (nodes: SankeyNode[], labelRight: boolean) => {
            nodes.forEach(n => {
                this.svgEl.append("rect")
                    .attr("x", n.x).attr("y", n.y)
                    .attr("width", nodeW).attr("height", Math.max(4, n.h))
                    .attr("rx", 3).attr("ry", 3)
                    .attr("fill", n.color)
                    .attr("class","sankey-node");

                const lx = labelRight ? n.x + nodeW + 6 : n.x - 6;
                const anchor = labelRight ? "start" : "end";
                const ly = n.y + n.h / 2;

                this.svgEl.append("text")
                    .attr("x", lx).attr("y", ly - 6)
                    .attr("text-anchor", anchor).attr("dominant-baseline","middle")
                    .attr("class","sankey-node-label")
                    .attr("font-family", fontFam)
                    .text(n.id);

                this.svgEl.append("text")
                    .attr("x", lx).attr("y", ly + 8)
                    .attr("text-anchor", anchor).attr("dominant-baseline","middle")
                    .attr("class","sankey-node-value")
                    .attr("font-family", fontFam)
                    .text(fmtNum(nodeTotal.get(n.id) ?? 0));
            });
        };

        drawNodes(srcNodes, false);
        drawNodes(dstNodes, true);
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
