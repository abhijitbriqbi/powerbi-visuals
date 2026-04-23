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
const TRIAL_KEY     = "briqlab_trial_WordCloud_start";
const PRO_STORE_KEY = "briqlab_wordcloud_prokey";

interface WordEntry {
    text:        string;
    weight:      number;
    sentiment:   number;
    fontSize:    number;
    x:           number;
    y:           number;
    rotation:    number;
    width:       number;
    height:      number;
    placed:      boolean;
    selectionId: ISelectionId;
}

interface BBox { x: number; y: number; w: number; h: number; }

function getTrial(): { daysLeft: number; expired: boolean } {
    try {
        let raw = localStorage.getItem(TRIAL_KEY);
        if (!raw) { raw = String(Date.now()); localStorage.setItem(TRIAL_KEY, raw); }
        const elapsed  = Date.now() - parseInt(raw, 10);
        const daysLeft = Math.max(0, Math.ceil((TRIAL_MS - elapsed) / 86400000));
        return { daysLeft, expired: elapsed > TRIAL_MS };
    } catch { return { daysLeft: 4, expired: false }; }
}

function measureText(text: string, fontSize: number, fontFamily: string): { w: number; h: number } {
    const canvas = document.createElement("canvas");
    const ctx    = canvas.getContext("2d");
    if (!ctx) return { w: text.length * fontSize * 0.6, h: fontSize };
    ctx.font = `${fontSize}px ${fontFamily}`;
    return { w: ctx.measureText(text).width, h: fontSize * 1.2 };
}

function boxesOverlap(a: BBox, b: BBox, padding: number): boolean {
    return !(a.x+a.w+padding < b.x-padding || b.x+b.w+padding < a.x-padding || a.y+a.h+padding < b.y-padding || b.y+b.h+padding < a.y-padding);
}

function spiralPlace(word: WordEntry, placed: WordEntry[], cx: number, cy: number, padding: number): boolean {
    const a = 2; let t = 0;
    for (let attempt = 0; attempt < 500; attempt++) {
        word.x = cx + a*t*Math.cos(t) - word.width/2;
        word.y = cy + a*t*Math.sin(t) - word.height/2;
        const bbox: BBox = { x: word.x, y: word.y, w: word.width, h: word.height };
        const collision  = placed.some(p => p.placed && boxesOverlap(bbox, { x:p.x, y:p.y, w:p.width, h:p.height }, padding));
        if (!collision) { word.placed = true; return true; }
        t += 0.2;
    }
    return false;
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

    private settings!:     VisualFormattingSettingsModel;
    private vp:            powerbi.IViewport = { width: 400, height: 300 };
    private selectedWords: Set<string> = new Set();

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
        this.root.classList.add("briqlab-wordcloud");

        this.contentEl = this.mkDiv("briqlab-visual-content");
        this.root.appendChild(this.contentEl);

        this.chartEl = this.mkDiv("briq-wordcloud-root");
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
                        dataItems: [{ displayName: "Briqlab Word Cloud", value: "" }],
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
            if (!dv?.categorical?.categories?.length) { this.renderEmpty("Add Word and Weight fields"); return; }
    
            const catData     = dv.categorical!;
            const wordCategory= catData.categories![0];
            const values      = catData.values;
    
            if (!values || values.length === 0) { this.renderEmpty("Add Word and Weight fields"); return; }
    
            let weightCol: powerbi.DataViewValueColumn | undefined;
            let sentimentCol: powerbi.DataViewValueColumn | undefined;
            for (const col of values) {
                const roles = col.source.roles ?? {};
                if (roles["weight"])   weightCol    = col;
                else if (roles["sentiment"]) sentimentCol = col;
            }
            if (!weightCol) { this.renderEmpty("Add Word and Weight fields"); return; }
    
            const ws    = this.settings.wordSettings;
            const cs    = this.settings.colorSettings;
            const es    = this.settings.exclusionSettings;
    
            const minFontSize   = Math.max(8,  ws.minFontSize.value ?? 10);
            const maxFontSize   = Math.min(80, ws.maxFontSize.value ?? 60);
            const maxWords      = Math.min(200, Math.max(20, ws.maxWords.value ?? 100));
            const rotationMode  = String(ws.rotations.value?.value ?? "None");
            const wordPadding   = ws.wordPadding.value ?? 4;
            const fontFamily    = String(ws.fontFamily.value?.value ?? "Segoe UI");
            const colorMode     = String(cs.colorMode.value?.value ?? "Rank");
            const uniformColor  = cs.uniformColor.value?.value  ?? "#0D9488";
            const positiveColor = cs.positiveColor.value?.value ?? "#0D9488";
            const negativeColor = cs.negativeColor.value?.value ?? "#EF4444";
            const neutralColor  = cs.neutralColor.value?.value  ?? "#94A3B8";
    
            const stopWordsRaw = es.stopWords.value ?? "";
            const caseSensitive= es.caseSensitive.value ?? false;
            const stopSet      = new Set<string>(stopWordsRaw.split(/[,\s\n]+/).map(w => caseSensitive ? w.trim() : w.trim().toLowerCase()).filter(w => w.length > 0));
    
            interface RawWord { text: string; weight: number; sentiment: number; idx: number; }
            const rawWords: RawWord[] = [];
            for (let i = 0; i < wordCategory.values.length; i++) {
                const rawText   = String(wordCategory.values[i] ?? "").trim();
                if (!rawText) continue;
                const compareText = caseSensitive ? rawText : rawText.toLowerCase();
                if (stopSet.has(compareText)) continue;
                const w = Number(weightCol.values[i] ?? 0);
                if (w <= 0) continue;
                const s = sentimentCol ? Number(sentimentCol.values[i] ?? 0) : 0;
                rawWords.push({ text: rawText, weight: w, sentiment: s, idx: i });
            }
    
            rawWords.sort((a,b) => b.weight - a.weight);
            const topWords = rawWords.slice(0, maxWords);
            if (topWords.length === 0) { this.renderEmpty("No words to display"); return; }
    
            const minW = topWords[topWords.length-1].weight;
            const maxW = topWords[0].weight;
            const sizeScale = d3.scaleLinear().domain([minW, maxW]).range([minFontSize, maxFontSize]);
    
            const rotAngles: number[] = [0];
            if (rotationMode === "Slight") rotAngles.push(15, -15);
            else if (rotationMode === "Full") rotAngles.push(90, -90);
    
            const wordEntries: WordEntry[] = topWords.map((rw, idx) => {
                const fontSize  = maxW === minW ? maxFontSize : sizeScale(rw.weight);
                const rotation  = rotAngles[idx % rotAngles.length];
                const measured  = measureText(rw.text, fontSize, fontFamily);
                return {
                    text: rw.text, weight: rw.weight, sentiment: rw.sentiment,
                    fontSize, x: 0, y: 0, rotation,
                    width:  rotation !== 0 ? measured.h : measured.w,
                    height: rotation !== 0 ? measured.w : measured.h,
                    placed: false,
                    selectionId: this.host.createSelectionIdBuilder().withCategory(wordCategory, rw.idx).createSelectionId()
                };
            });
    
            const cx = this.vp.width / 2, cy = this.vp.height / 2;
            const placed: WordEntry[] = [];
            for (const entry of wordEntries) { spiralPlace(entry, placed, cx, cy, wordPadding); placed.push(entry); }
    
            const rankColorScale = d3.scaleLinear<string>().domain([0, topWords.length-1]).range(["#0D9488","#CCFBF1"]);
            const getColor = (entry: WordEntry, rank: number): string => {
                if (colorMode === "Uniform")   return uniformColor;
                if (colorMode === "Sentiment") { const s = entry.sentiment; return s > 0.2 ? positiveColor : s < -0.2 ? negativeColor : neutralColor; }
                return rankColorScale(rank);
            };
    
            this.renderChart(wordEntries, fontFamily, getColor, sentimentCol !== undefined);
            this.renderingManager.renderingFinished(options);
        } catch (e: unknown) {
            this.renderingManager.renderingFailed(options, String(e));
        }
    }

    private renderChart(wordEntries: WordEntry[], fontFamily: string, getColor: (e: WordEntry, r: number) => string, hasSentiment: boolean): void {
        while (this.chartEl.firstChild) this.chartEl.removeChild(this.chartEl.firstChild);

        const width  = this.vp.width;
        const height = this.vp.height;
        const maxFS  = this.settings.wordSettings.maxFontSize.value ?? 60;

        this.chartEl.style.width    = `${width}px`;
        this.chartEl.style.height   = `${height}px`;
        this.chartEl.style.position = "relative";

        const visual = this.mkDiv("briq-visual-content");
        visual.style.cssText = "position:absolute;top:0;left:0;";
        visual.style.width   = `${width}px`;
        visual.style.height  = `${height}px`;
        this.chartEl.appendChild(visual);

        const svg = d3.select(visual).append("svg").attr("width",width).attr("height",height);

        const tooltip = this.mkDiv("briq-tooltip");
        tooltip.style.display = "none";
        visual.appendChild(tooltip);

        const self = this;

        const wordTexts = svg.selectAll<SVGTextElement, WordEntry>("text.briq-word")
            .data(wordEntries.filter(w => w.placed))
            .enter().append("text")
            .attr("class","briq-word")
            .attr("text-anchor","middle").attr("dominant-baseline","middle")
            .attr("font-family",fontFamily)
            .attr("font-size",d => d.fontSize)
            .attr("font-weight",d => d.fontSize > maxFS*0.7 ? "700" : "400")
            .attr("fill",(d,i) => getColor(d,i))
            .attr("transform",d => {
                const tx = d.x + d.width/2, ty = d.y + d.height/2;
                return d.rotation !== 0 ? `translate(${tx},${ty}) rotate(${d.rotation})` : `translate(${tx},${ty})`;
            })
            .style("cursor","pointer")
            .style("transition","opacity 0.2s, transform 0.2s")
            .text(d => d.text);

        wordTexts
            .on("mouseover", function(event: MouseEvent, d: WordEntry) {
                wordTexts.style("opacity", (other: WordEntry) => other === d ? "1" : "0.4");
                d3.select<SVGTextElement, WordEntry>(this).attr("transform", (wd: WordEntry) => {
                    const tx = wd.x+wd.width/2, ty = wd.y+wd.height/2;
                    const rot = wd.rotation !== 0 ? ` rotate(${wd.rotation})` : "";
                    return `translate(${tx},${ty})${rot} scale(1.1)`;
                });
                while (tooltip.firstChild) tooltip.removeChild(tooltip.firstChild);
                const l1 = document.createElement("div"); l1.style.fontWeight = "600"; l1.textContent = d.text; tooltip.appendChild(l1);
                const l2 = document.createElement("div"); l2.textContent = `Weight: ${d.weight.toLocaleString()}`; tooltip.appendChild(l2);
                if (hasSentiment) { const l3 = document.createElement("div"); l3.textContent = `Sentiment: ${d.sentiment.toFixed(2)}`; tooltip.appendChild(l3); }
                tooltip.style.display = "block";
                tooltip.style.left = `${event.offsetX+10}px`;
                tooltip.style.top  = `${event.offsetY-10}px`;
            })
            .on("mousemove", function(event: MouseEvent) {
                tooltip.style.left = `${event.offsetX+10}px`;
                tooltip.style.top  = `${event.offsetY-10}px`;
            })
            .on("mouseout", function(_event: MouseEvent, d: WordEntry) {
                wordTexts.style("opacity","1");
                d3.select<SVGTextElement, WordEntry>(this).attr("transform", (wd: WordEntry) => {
                    const tx = wd.x+wd.width/2, ty = wd.y+wd.height/2;
                    return wd.rotation !== 0 ? `translate(${tx},${ty}) rotate(${wd.rotation})` : `translate(${tx},${ty})`;
                });
                tooltip.style.display = "none";
                void d;
            })
            .on("click", function(event: MouseEvent, d: WordEntry) {
                const isMulti = event.ctrlKey || event.metaKey;
                if (self.selectedWords.has(d.text) && !isMulti) {
                    self.selectedWords.clear(); self.selMgr.clear();
                    wordTexts.style("opacity","1");
                } else {
                    if (!isMulti) self.selectedWords.clear();
                    self.selectedWords.add(d.text);
                    const filteredEntries = wordEntries.filter(e => e.placed);
                    const ids = filteredEntries.filter(e => self.selectedWords.has(e.text)).map(e => e.selectionId);
                    self.selMgr.select(ids, isMulti);
                    wordTexts.style("opacity", (other: WordEntry) => self.selectedWords.size === 0 ? "1" : self.selectedWords.has(other.text) ? "1" : "0.3");
                }
                event.stopPropagation();
            });

        wordTexts.on("contextmenu", function(event: MouseEvent, d: WordEntry) {
            event.preventDefault();
            event.stopPropagation();
            self.selMgr.showContextMenu(
                d.selectionId,
                { x: event.clientX, y: event.clientY }
            );
        });

        svg.on("click", () => {
            self.selectedWords.clear(); self.selMgr.clear(); wordTexts.style("opacity","1");
        });
    }

    private renderEmpty(msg: string): void {
        while (this.chartEl.firstChild) this.chartEl.removeChild(this.chartEl.firstChild);
        const el = this.mkDiv("briq-empty-state"); el.textContent = msg; this.chartEl.appendChild(el);
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
