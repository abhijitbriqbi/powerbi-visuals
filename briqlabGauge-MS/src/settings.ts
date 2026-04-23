"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

// ── Gauge card ─────────────────────────────────────────────────────────────────
class GaugeSettingsCard extends FormattingSettingsCard {
    gaugeType = new formattingSettings.ItemDropdown({
        name: "gaugeType",
        displayName: "Gauge Type",
        items: [
            { value: "semi",         displayName: "Semi-circle (180°)"  },
            { value: "threequarter", displayName: "Three-quarter (270°)" },
            { value: "full",         displayName: "Full circle (360°)"  }
        ],
        value: { value: "semi", displayName: "Semi-circle (180°)" }
    });

    arcThickness = new formattingSettings.NumUpDown({
        name: "arcThickness",
        displayName: "Arc Thickness %",
        value: 18
    });

    showTarget = new formattingSettings.ToggleSwitch({
        name: "showTarget",
        displayName: "Show Target Marker",
        value: true
    });

    targetColor = new formattingSettings.ColorPicker({
        name: "targetColor",
        displayName: "Target Marker Color",
        value: { value: "#F97316" }
    });

    manualColorMode = new formattingSettings.ToggleSwitch({
        name: "manualColorMode",
        displayName: "Manual Color Mode",
        value: false
    });

    manualColor = new formattingSettings.ColorPicker({
        name: "manualColor",
        displayName: "Manual Arc Color",
        value: { value: "#0D9488" }
    });

    showNeedle = new formattingSettings.ToggleSwitch({
        name: "showNeedle",
        displayName: "Show Needle",
        value: true
    });

    showTicks = new formattingSettings.ToggleSwitch({
        name: "showTicks",
        displayName: "Show Tick Marks",
        value: true
    });

    tickCount = new formattingSettings.NumUpDown({
        name: "tickCount",
        displayName: "Tick Count",
        value: 5
    });

    valueFontSize = new formattingSettings.NumUpDown({
        name: "valueFontSize",
        displayName: "Value Font Size",
        value: 28
    });

    showPctComplete = new formattingSettings.ToggleSwitch({
        name: "showPctComplete",
        displayName: "Show % Complete",
        value: true
    });

    showMinMax = new formattingSettings.ToggleSwitch({
        name: "showMinMax",
        displayName: "Show Min/Max Labels",
        value: true
    });

    name: string = "gaugeSettings";
    displayName: string = "Gauge";
    slices: Array<FormattingSettingsSlice> = [
        this.gaugeType,
        this.arcThickness,
        this.showNeedle,
        this.showTicks,
        this.tickCount,
        this.showTarget,
        this.targetColor,
        this.manualColorMode,
        this.manualColor,
        this.valueFontSize,
        this.showPctComplete,
        this.showMinMax
    ];
}

// ── Zone colors card ───────────────────────────────────────────────────────────
class ZoneSettingsCard extends FormattingSettingsCard {
    showZones = new formattingSettings.ToggleSwitch({
        name: "showZones",
        displayName: "Show Color Zones",
        value: true
    });

    zone1MaxPct = new formattingSettings.NumUpDown({
        name: "zone1MaxPct",
        displayName: "Zone 1 Max % (Red zone)",
        value: 50
    });

    zone2MaxPct = new formattingSettings.NumUpDown({
        name: "zone2MaxPct",
        displayName: "Zone 2 Max % (Amber zone)",
        value: 80
    });

    zone1Color = new formattingSettings.ColorPicker({
        name: "zone1Color",
        displayName: "Zone 1 Color",
        value: { value: "#EF4444" }
    });

    zone2Color = new formattingSettings.ColorPicker({
        name: "zone2Color",
        displayName: "Zone 2 Color",
        value: { value: "#F59E0B" }
    });

    zone3Color = new formattingSettings.ColorPicker({
        name: "zone3Color",
        displayName: "Zone 3 Color (Good)",
        value: { value: "#10B981" }
    });

    trackColor = new formattingSettings.ColorPicker({
        name: "trackColor",
        displayName: "Track (background arc) Color",
        value: { value: "#E5E7EB" }
    });

    name: string = "zoneSettings";
    displayName: string = "Color Zones";
    slices: Array<FormattingSettingsSlice> = [
        this.showZones,
        this.zone1MaxPct,
        this.zone2MaxPct,
        this.zone1Color,
        this.zone2Color,
        this.zone3Color,
        this.trackColor
    ];
}

// ── Font card ──────────────────────────────────────────────────────────────────
class FontSettingsCard extends FormattingSettingsCard {
    fontFamily = new formattingSettings.ItemDropdown({
        name: "fontFamily",
        displayName: "Font Family",
        items: [
            { value: "Segoe UI",    displayName: "Segoe UI (default)" },
            { value: "Arial",       displayName: "Arial"               },
            { value: "Calibri",     displayName: "Calibri"             },
            { value: "Verdana",     displayName: "Verdana"             },
            { value: "Georgia",     displayName: "Georgia"             },
            { value: "Courier New", displayName: "Courier New"         }
        ],
        value: { value: "Segoe UI", displayName: "Segoe UI (default)" }
    });

    boldValue = new formattingSettings.ToggleSwitch({
        name: "boldValue",
        displayName: "Bold Value",
        value: true
    });

    valuePrefix = new formattingSettings.TextInput({
        name: "valuePrefix",
        displayName: "Value Prefix",
        placeholder: "e.g. $, ₹",
        value: ""
    });

    valueSuffix = new formattingSettings.TextInput({
        name: "valueSuffix",
        displayName: "Value Suffix",
        placeholder: "e.g. %",
        value: ""
    });

    valueColor = new formattingSettings.ColorPicker({
        name: "valueColor",
        displayName: "Value Text Color",
        value: { value: "#374151" }
    });

    name: string = "fontSettings";
    displayName: string = "Value & Font";
    slices: Array<FormattingSettingsSlice> = [
        this.fontFamily,
        this.boldValue,
        this.valuePrefix,
        this.valueSuffix,
        this.valueColor
    ];
}

// ── Pro card ───────────────────────────────────────────────────────────────────

// ── Model ──────────────────────────────────────────────────────────────────────
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    gaugeSettings = new GaugeSettingsCard();
    zoneSettings  = new ZoneSettingsCard();
    fontSettings  = new FontSettingsCard();
    cards = [
        this.gaugeSettings,
        this.zoneSettings,
        this.fontSettings
    ];
}
