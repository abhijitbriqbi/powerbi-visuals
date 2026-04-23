"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard  = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

// ── Donut card ─────────────────────────────────────────────────────────────────
class DonutSettingsCard extends FormattingSettingsCard {
    innerRadius = new formattingSettings.NumUpDown({
        name: "innerRadius",
        displayName: "Inner Radius % (hole size)",
        value: 55
    });

    outerRadius = new formattingSettings.NumUpDown({
        name: "outerRadius",
        displayName: "Outer Radius % of space",
        value: 85
    });

    sortOrder = new formattingSettings.ItemDropdown({
        name: "sortOrder",
        displayName: "Sort Order",
        items: [
            { value: "original", displayName: "Original order"       },
            { value: "desc",     displayName: "Value: Largest first" },
            { value: "asc",      displayName: "Value: Smallest first"}
        ],
        value: { value: "original", displayName: "Original order" }
    });

    startAngle = new formattingSettings.NumUpDown({
        name: "startAngle",
        displayName: "Start Angle (degrees, 0 = top)",
        value: 0
    });

    showCenter = new formattingSettings.ToggleSwitch({
        name: "showCenter",
        displayName: "Show Center Value",
        value: true
    });

    centerLabel = new formattingSettings.TextInput({
        name: "centerLabel",
        displayName: "Center Label",
        placeholder: "e.g. Total",
        value: "Total"
    });

    centerValueDecimals = new formattingSettings.NumUpDown({
        name: "centerValueDecimals",
        displayName: "Center Value Decimals",
        value: 0
    });

    centerValueFormat = new formattingSettings.ItemDropdown({
        name: "centerValueFormat",
        displayName: "Center Value Format",
        items: [
            { value: "auto",  displayName: "Auto (K/M)"  },
            { value: "full",  displayName: "Full number" }
        ],
        value: { value: "auto", displayName: "Auto (K/M)" }
    });

    borderWidth = new formattingSettings.NumUpDown({
        name: "borderWidth",
        displayName: "Segment Border Width",
        value: 2
    });

    minLabelPct = new formattingSettings.NumUpDown({
        name: "minLabelPct",
        displayName: "Min % to Show Label",
        value: 4
    });

    name: string = "donutSettings";
    displayName: string = "Donut";
    slices: Array<FormattingSettingsSlice> = [
        this.innerRadius,
        this.outerRadius,
        this.sortOrder,
        this.startAngle,
        this.showCenter,
        this.centerLabel,
        this.centerValueDecimals,
        this.centerValueFormat,
        this.borderWidth,
        this.minLabelPct
    ];
}

// ── Labels card ────────────────────────────────────────────────────────────────
class LabelSettingsCard extends FormattingSettingsCard {
    showLabels = new formattingSettings.ToggleSwitch({
        name: "showLabels",
        displayName: "Show Labels",
        value: true
    });

    labelFontSize = new formattingSettings.NumUpDown({
        name: "labelFontSize",
        displayName: "Label Font Size",
        value: 11
    });

    labelFormat = new formattingSettings.ItemDropdown({
        name: "labelFormat",
        displayName: "Label Format",
        items: [
            { value: "name_pct",   displayName: "Name + %"    },
            { value: "name_value", displayName: "Name + Value" },
            { value: "pct",        displayName: "% only"       },
            { value: "name",       displayName: "Name only"    },
            { value: "value",      displayName: "Value only"   }
        ],
        value: { value: "name_pct", displayName: "Name + %" }
    });

    labelColor = new formattingSettings.ColorPicker({
        name: "labelColor",
        displayName: "Label Color",
        value: { value: "#374151" }
    });

    boldLabels = new formattingSettings.ToggleSwitch({
        name: "boldLabels",
        displayName: "Bold Labels",
        value: false
    });

    fontFamily = new formattingSettings.ItemDropdown({
        name: "fontFamily",
        displayName: "Font Family",
        items: [
            { value: "Segoe UI",    displayName: "Segoe UI"   },
            { value: "Arial",       displayName: "Arial"       },
            { value: "Calibri",     displayName: "Calibri"     },
            { value: "Verdana",     displayName: "Verdana"     },
            { value: "Georgia",     displayName: "Georgia"     },
            { value: "Courier New", displayName: "Courier New" }
        ],
        value: { value: "Segoe UI", displayName: "Segoe UI" }
    });

    name: string = "labelSettings";
    displayName: string = "Labels";
    slices: Array<FormattingSettingsSlice> = [
        this.showLabels,
        this.labelFontSize,
        this.labelFormat,
        this.labelColor,
        this.boldLabels,
        this.fontFamily
    ];
}

// ── Colors card ────────────────────────────────────────────────────────────────
class ColorSettingsCard extends FormattingSettingsCard {
    color1  = new formattingSettings.ColorPicker({ name: "color1",  displayName: "Segment 1 Color",  value: { value: "#0D9488" } });
    color2  = new formattingSettings.ColorPicker({ name: "color2",  displayName: "Segment 2 Color",  value: { value: "#F97316" } });
    color3  = new formattingSettings.ColorPicker({ name: "color3",  displayName: "Segment 3 Color",  value: { value: "#3B82F6" } });
    color4  = new formattingSettings.ColorPicker({ name: "color4",  displayName: "Segment 4 Color",  value: { value: "#8B5CF6" } });
    color5  = new formattingSettings.ColorPicker({ name: "color5",  displayName: "Segment 5 Color",  value: { value: "#10B981" } });
    color6  = new formattingSettings.ColorPicker({ name: "color6",  displayName: "Segment 6 Color",  value: { value: "#EF4444" } });
    color7  = new formattingSettings.ColorPicker({ name: "color7",  displayName: "Segment 7 Color",  value: { value: "#F59E0B" } });
    color8  = new formattingSettings.ColorPicker({ name: "color8",  displayName: "Segment 8 Color",  value: { value: "#EC4899" } });
    color9  = new formattingSettings.ColorPicker({ name: "color9",  displayName: "Segment 9 Color",  value: { value: "#06B6D4" } });
    color10 = new formattingSettings.ColorPicker({ name: "color10", displayName: "Segment 10 Color", value: { value: "#84CC16" } });

    name: string = "colorSettings";
    displayName: string = "Segment Colors";
    slices: Array<FormattingSettingsSlice> = [
        this.color1, this.color2, this.color3, this.color4, this.color5,
        this.color6, this.color7, this.color8, this.color9, this.color10
    ];
}

// ── Legend card ────────────────────────────────────────────────────────────────
class LegendSettingsCard extends FormattingSettingsCard {
    showLegend = new formattingSettings.ToggleSwitch({
        name: "showLegend",
        displayName: "Show Legend",
        value: true
    });

    legendPosition = new formattingSettings.ItemDropdown({
        name: "legendPosition",
        displayName: "Position",
        items: [
            { value: "Right",  displayName: "Right"  },
            { value: "Left",   displayName: "Left"   },
            { value: "Top",    displayName: "Top"    },
            { value: "Bottom", displayName: "Bottom" }
        ],
        value: { value: "Right", displayName: "Right" }
    });

    legendFontSize = new formattingSettings.NumUpDown({
        name: "legendFontSize",
        displayName: "Legend Font Size",
        value: 11
    });

    name: string = "legendSettings";
    displayName: string = "Legend";
    slices: Array<FormattingSettingsSlice> = [
        this.showLegend,
        this.legendPosition,
        this.legendFontSize
    ];
}

// ── Pro card ───────────────────────────────────────────────────────────────────

// ── Model ──────────────────────────────────────────────────────────────────────
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    donutSettings  = new DonutSettingsCard();
    labelSettings  = new LabelSettingsCard();
    colorSettings  = new ColorSettingsCard();
    legendSettings = new LegendSettingsCard();
    cards = [
        this.donutSettings,
        this.labelSettings,
        this.colorSettings,
        this.legendSettings
    ];
}
