"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard  = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

// ── Chart Style ────────────────────────────────────────────────────────────────
class ChartStyleCard extends FormattingSettingsCard {
    layoutMode = new formattingSettings.ItemDropdown({
        name: "layoutMode",
        displayName: "Layout Mode",
        items: [
            { value: "Scatter", displayName: "Scatter" },
            { value: "Packed",  displayName: "Packed"  }
        ],
        value: { value: "Scatter", displayName: "Scatter" }
    });

    bubbleOpacity = new formattingSettings.NumUpDown({
        name: "bubbleOpacity",
        displayName: "Bubble Opacity %",
        value: 80
    });

    minBubbleSize = new formattingSettings.NumUpDown({
        name: "minBubbleSize",
        displayName: "Min Bubble Radius (px)",
        value: 8
    });

    maxBubbleSize = new formattingSettings.NumUpDown({
        name: "maxBubbleSize",
        displayName: "Max Bubble Radius (px)",
        value: 50
    });

    name: string        = "chartStyle";
    displayName: string = "Chart Style";
    slices: Array<FormattingSettingsSlice> = [
        this.layoutMode,
        this.bubbleOpacity,
        this.minBubbleSize,
        this.maxBubbleSize
    ];
}

// ── X Axis ─────────────────────────────────────────────────────────────────────
class XAxisCard extends FormattingSettingsCard {
    showAxis = new formattingSettings.ToggleSwitch({
        name: "showAxis",
        displayName: "Show X Axis",
        value: true
    });

    axisLabel = new formattingSettings.TextInput({
        name: "axisLabel",
        displayName: "X Axis Label",
        placeholder: "e.g. Revenue",
        value: ""
    });

    axisMin = new formattingSettings.NumUpDown({
        name: "axisMin",
        displayName: "Min Value (0 = auto)",
        value: 0
    });

    axisMax = new formattingSettings.NumUpDown({
        name: "axisMax",
        displayName: "Max Value (0 = auto)",
        value: 0
    });

    name: string        = "xAxisSettings";
    displayName: string = "X Axis";
    slices: Array<FormattingSettingsSlice> = [
        this.showAxis,
        this.axisLabel,
        this.axisMin,
        this.axisMax
    ];
}

// ── Y Axis ─────────────────────────────────────────────────────────────────────
class YAxisCard extends FormattingSettingsCard {
    showAxis = new formattingSettings.ToggleSwitch({
        name: "showAxis",
        displayName: "Show Y Axis",
        value: true
    });

    axisLabel = new formattingSettings.TextInput({
        name: "axisLabel",
        displayName: "Y Axis Label",
        placeholder: "e.g. Profit",
        value: ""
    });

    axisMin = new formattingSettings.NumUpDown({
        name: "axisMin",
        displayName: "Min Value (0 = auto)",
        value: 0
    });

    axisMax = new formattingSettings.NumUpDown({
        name: "axisMax",
        displayName: "Max Value (0 = auto)",
        value: 0
    });

    name: string        = "yAxisSettings";
    displayName: string = "Y Axis";
    slices: Array<FormattingSettingsSlice> = [
        this.showAxis,
        this.axisLabel,
        this.axisMin,
        this.axisMax
    ];
}

// ── Bubble Colors ──────────────────────────────────────────────────────────────
class BubbleColorsCard extends FormattingSettingsCard {
    color1  = new formattingSettings.ColorPicker({ name: "color1",  displayName: "Color 1",  value: { value: "#0D9488" } });
    color2  = new formattingSettings.ColorPicker({ name: "color2",  displayName: "Color 2",  value: { value: "#F97316" } });
    color3  = new formattingSettings.ColorPicker({ name: "color3",  displayName: "Color 3",  value: { value: "#3B82F6" } });
    color4  = new formattingSettings.ColorPicker({ name: "color4",  displayName: "Color 4",  value: { value: "#8B5CF6" } });
    color5  = new formattingSettings.ColorPicker({ name: "color5",  displayName: "Color 5",  value: { value: "#10B981" } });
    color6  = new formattingSettings.ColorPicker({ name: "color6",  displayName: "Color 6",  value: { value: "#EF4444" } });
    color7  = new formattingSettings.ColorPicker({ name: "color7",  displayName: "Color 7",  value: { value: "#F59E0B" } });
    color8  = new formattingSettings.ColorPicker({ name: "color8",  displayName: "Color 8",  value: { value: "#EC4899" } });
    color9  = new formattingSettings.ColorPicker({ name: "color9",  displayName: "Color 9",  value: { value: "#06B6D4" } });
    color10 = new formattingSettings.ColorPicker({ name: "color10", displayName: "Color 10", value: { value: "#84CC16" } });

    name: string        = "bubbleColors";
    displayName: string = "Bubble Colors";
    slices: Array<FormattingSettingsSlice> = [
        this.color1, this.color2, this.color3, this.color4, this.color5,
        this.color6, this.color7, this.color8, this.color9, this.color10
    ];
}

// ── Labels ─────────────────────────────────────────────────────────────────────
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

    name: string        = "labelSettings";
    displayName: string = "Labels";
    slices: Array<FormattingSettingsSlice> = [
        this.showLabels,
        this.labelFontSize,
        this.labelColor,
        this.boldLabels,
        this.fontFamily
    ];
}

// ── Legend ─────────────────────────────────────────────────────────────────────
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
            { value: "Bottom", displayName: "Bottom" },
            { value: "Top",    displayName: "Top"    },
            { value: "Left",   displayName: "Left"   },
            { value: "Right",  displayName: "Right"  },
            { value: "None",   displayName: "None"   }
        ],
        value: { value: "Bottom", displayName: "Bottom" }
    });

    legendFontSize = new formattingSettings.NumUpDown({
        name: "legendFontSize",
        displayName: "Font Size",
        value: 11
    });

    name: string        = "legendSettings";
    displayName: string = "Legend";
    slices: Array<FormattingSettingsSlice> = [
        this.showLegend,
        this.legendPosition,
        this.legendFontSize
    ];
}

// ── Tooltip ────────────────────────────────────────────────────────────────────
class TooltipSettingsCard extends FormattingSettingsCard {
    showTooltip = new formattingSettings.ToggleSwitch({
        name: "showTooltip",
        displayName: "Show Tooltip",
        value: true
    });

    name: string        = "tooltipSettings";
    displayName: string = "Tooltip";
    slices: Array<FormattingSettingsSlice> = [this.showTooltip];
}

// ── Briqlab Pro ────────────────────────────────────────────────────────────────

// ── Model ──────────────────────────────────────────────────────────────────────
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    chartStyle      = new ChartStyleCard();
    xAxisSettings   = new XAxisCard();
    yAxisSettings   = new YAxisCard();
    bubbleColors    = new BubbleColorsCard();
    labelSettings   = new LabelSettingsCard();
    legendSettings  = new LegendSettingsCard();
    tooltipSettings = new TooltipSettingsCard();
    cards = [
        this.chartStyle,
        this.xAxisSettings,
        this.yAxisSettings,
        this.bubbleColors,
        this.labelSettings,
        this.legendSettings,
        this.tooltipSettings
    ];
}
