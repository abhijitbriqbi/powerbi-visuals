"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard  = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

// ── Chart card ─────────────────────────────────────────────────────────────────
class ChartSettingsCard extends FormattingSettingsCard {
    orientation = new formattingSettings.ItemDropdown({
        name: "orientation",
        displayName: "Orientation",
        items: [
            { value: "vertical",   displayName: "Vertical"   },
            { value: "horizontal", displayName: "Horizontal" }
        ],
        value: { value: "vertical", displayName: "Vertical" }
    });

    barMode = new formattingSettings.ItemDropdown({
        name: "barMode",
        displayName: "Bar Mode",
        items: [
            { value: "grouped",  displayName: "Grouped (side-by-side)" },
            { value: "stacked",  displayName: "Stacked"                },
            { value: "stacked100", displayName: "100% Stacked"         }
        ],
        value: { value: "grouped", displayName: "Grouped (side-by-side)" }
    });

    barColor = new formattingSettings.ColorPicker({
        name: "barColor",
        displayName: "Series 1 Color",
        value: { value: "#0D9488" }
    });

    comparisonColor = new formattingSettings.ColorPicker({
        name: "comparisonColor",
        displayName: "Series 2 Color",
        value: { value: "#F97316" }
    });

    series3Color = new formattingSettings.ColorPicker({
        name: "series3Color",
        displayName: "Series 3 Color",
        value: { value: "#3B82F6" }
    });

    cornerRadius = new formattingSettings.NumUpDown({
        name: "cornerRadius",
        displayName: "Bar Corner Radius",
        value: 4
    });

    barPadding = new formattingSettings.NumUpDown({
        name: "barPadding",
        displayName: "Bar Padding (0–50%)",
        value: 25
    });

    showBackground = new formattingSettings.ToggleSwitch({
        name: "showBackground",
        displayName: "Show Background",
        value: false
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Background Color",
        value: { value: "#FFFFFF" }
    });

    name: string = "chartSettings";
    displayName: string = "Chart";
    slices: Array<FormattingSettingsSlice> = [
        this.orientation,
        this.barMode,
        this.barColor,
        this.comparisonColor,
        this.series3Color,
        this.cornerRadius,
        this.barPadding,
        this.showBackground,
        this.backgroundColor
    ];
}

// ── Labels card ────────────────────────────────────────────────────────────────
class LabelSettingsCard extends FormattingSettingsCard {
    showLabels = new formattingSettings.ToggleSwitch({
        name: "showLabels",
        displayName: "Show Data Labels",
        value: false
    });

    labelFontSize = new formattingSettings.NumUpDown({
        name: "labelFontSize",
        displayName: "Label Font Size",
        value: 10
    });

    labelFormat = new formattingSettings.ItemDropdown({
        name: "labelFormat",
        displayName: "Label Format",
        items: [
            { value: "auto",    displayName: "Auto (K/M)"   },
            { value: "value",   displayName: "Full Value"   },
            { value: "pct",     displayName: "Percentage %"  }
        ],
        value: { value: "auto", displayName: "Auto (K/M)" }
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

    name: string = "labelSettings";
    displayName: string = "Data Labels";
    slices: Array<FormattingSettingsSlice> = [
        this.showLabels,
        this.labelFontSize,
        this.labelFormat,
        this.labelColor,
        this.boldLabels
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

    boldAxis = new formattingSettings.ToggleSwitch({
        name: "boldAxis",
        displayName: "Bold Axis Labels",
        value: false
    });

    italicAxis = new formattingSettings.ToggleSwitch({
        name: "italicAxis",
        displayName: "Italic Axis Labels",
        value: false
    });

    name: string = "fontSettings";
    displayName: string = "Font";
    slices: Array<FormattingSettingsSlice> = [
        this.fontFamily,
        this.boldAxis,
        this.italicAxis
    ];
}

// ── Number format card ─────────────────────────────────────────────────────────
class NumberFormatCard extends FormattingSettingsCard {
    decimalPlaces = new formattingSettings.NumUpDown({
        name: "decimalPlaces",
        displayName: "Decimal Places",
        value: 1
    });

    useKM = new formattingSettings.ToggleSwitch({
        name: "useKM",
        displayName: "Abbreviate (K / M / B)",
        value: true
    });

    prefix = new formattingSettings.TextInput({
        name: "prefix",
        displayName: "Value Prefix",
        placeholder: "e.g. $, ₹",
        value: ""
    });

    suffix = new formattingSettings.TextInput({
        name: "suffix",
        displayName: "Value Suffix",
        placeholder: "e.g. USD, kg",
        value: ""
    });

    name: string = "numberFormat";
    displayName: string = "Number Format";
    slices: Array<FormattingSettingsSlice> = [
        this.decimalPlaces,
        this.useKM,
        this.prefix,
        this.suffix
    ];
}

// ── Axis card ──────────────────────────────────────────────────────────────────
class AxisSettingsCard extends FormattingSettingsCard {
    xFontSize = new formattingSettings.NumUpDown({
        name: "xFontSize",
        displayName: "X Axis Font Size",
        value: 10
    });

    yFontSize = new formattingSettings.NumUpDown({
        name: "yFontSize",
        displayName: "Y Axis Font Size",
        value: 10
    });

    xTitle = new formattingSettings.TextInput({
        name: "xTitle",
        displayName: "X Axis Title",
        placeholder: "Category",
        value: ""
    });

    yTitle = new formattingSettings.TextInput({
        name: "yTitle",
        displayName: "Y Axis Title",
        placeholder: "Value",
        value: ""
    });

    showGridlines = new formattingSettings.ToggleSwitch({
        name: "showGridlines",
        displayName: "Show Gridlines",
        value: true
    });

    gridlineColor = new formattingSettings.ColorPicker({
        name: "gridlineColor",
        displayName: "Gridline Color",
        value: { value: "#E2E8F0" }
    });

    showZeroLine = new formattingSettings.ToggleSwitch({
        name: "showZeroLine",
        displayName: "Show Zero Line",
        value: false
    });

    name: string = "axisSettings";
    displayName: string = "Axes";
    slices: Array<FormattingSettingsSlice> = [
        this.xFontSize,
        this.yFontSize,
        this.xTitle,
        this.yTitle,
        this.showGridlines,
        this.gridlineColor,
        this.showZeroLine
    ];
}

// ── Pro card ───────────────────────────────────────────────────────────────────

// ── Model ──────────────────────────────────────────────────────────────────────
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    chartSettings  = new ChartSettingsCard();
    labelSettings  = new LabelSettingsCard();
    fontSettings   = new FontSettingsCard();
    numberFormat   = new NumberFormatCard();
    axisSettings   = new AxisSettingsCard();
    cards = [
        this.chartSettings,
        this.labelSettings,
        this.fontSettings,
        this.numberFormat,
        this.axisSettings
    ];
}
