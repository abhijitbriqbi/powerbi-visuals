"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard  = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

// ── Chart Style ────────────────────────────────────────────────────────────────
class ChartStyleCard extends FormattingSettingsCard {
    chartType = new formattingSettings.ItemDropdown({
        name: "chartType",
        displayName: "Chart Type",
        items: [
            { value: "Donut", displayName: "Donut" },
            { value: "Pie",   displayName: "Pie"   }
        ],
        value: { value: "Donut", displayName: "Donut" }
    });

    innerRadius = new formattingSettings.NumUpDown({
        name: "innerRadius",
        displayName: "Inner Radius %",
        value: 55
    });

    padAngle = new formattingSettings.NumUpDown({
        name: "padAngle",
        displayName: "Segment Gap",
        value: 1.5
    });

    name: string        = "chartStyle";
    displayName: string = "Chart Style";
    slices: Array<FormattingSettingsSlice> = [
        this.chartType,
        this.innerRadius,
        this.padAngle
    ];
}

// ── Center Display ─────────────────────────────────────────────────────────────
class CenterDisplayCard extends FormattingSettingsCard {
    showCenter = new formattingSettings.ToggleSwitch({
        name: "showCenter",
        displayName: "Show Center",
        value: true
    });

    centerLabel = new formattingSettings.TextInput({
        name: "centerLabel",
        displayName: "Center Label",
        placeholder: "e.g. Total",
        value: "Total"
    });

    centerValueSize = new formattingSettings.NumUpDown({
        name: "centerValueSize",
        displayName: "Value Font Size",
        value: 28
    });

    centerLabelSize = new formattingSettings.NumUpDown({
        name: "centerLabelSize",
        displayName: "Label Font Size",
        value: 11
    });

    name: string        = "centerDisplay";
    displayName: string = "Center Display";
    slices: Array<FormattingSettingsSlice> = [
        this.showCenter,
        this.centerLabel,
        this.centerValueSize,
        this.centerLabelSize
    ];
}

// ── Segments ───────────────────────────────────────────────────────────────────
class SegmentsCard extends FormattingSettingsCard {
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

    name: string        = "segments";
    displayName: string = "Segments";
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

    labelThreshold = new formattingSettings.NumUpDown({
        name: "labelThreshold",
        displayName: "Hide Below %",
        value: 4
    });

    labelFormat = new formattingSettings.ItemDropdown({
        name: "labelFormat",
        displayName: "Label Format",
        items: [
            { value: "Name",      displayName: "Name"     },
            { value: "Name (%)",  displayName: "Name (%)" },
            { value: "% only",    displayName: "% only"   }
        ],
        value: { value: "Name (%)", displayName: "Name (%)" }
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
        this.labelThreshold,
        this.labelFormat,
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
    chartStyle     = new ChartStyleCard();
    centerDisplay  = new CenterDisplayCard();
    segments       = new SegmentsCard();
    labelSettings  = new LabelSettingsCard();
    legendSettings = new LegendSettingsCard();
    tooltipSettings = new TooltipSettingsCard();
    cards = [
        this.chartStyle,
        this.centerDisplay,
        this.segments,
        this.labelSettings,
        this.legendSettings,
        this.tooltipSettings
    ];
}
