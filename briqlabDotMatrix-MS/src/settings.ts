"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

class GridSettingsCard extends FormattingSettingsCard {
    dotShape = new formattingSettings.ItemDropdown({
        name: "dotShape",
        displayName: "Dot Shape",
        items: [
            { displayName: "Circle", value: "Circle" },
            { displayName: "Square", value: "Square" }
        ],
        value: { displayName: "Circle", value: "Circle" }
    });

    maxDots = new formattingSettings.NumUpDown({
        name: "maxDots",
        displayName: "Max Dots",
        value: 100
    });

    dotGap = new formattingSettings.NumUpDown({
        name: "dotGap",
        displayName: "Dot Gap",
        value: 3
    });

    showCenterText = new formattingSettings.ToggleSwitch({
        name: "showCenterText",
        displayName: "Show Center Text",
        value: true
    });

    achievedColor = new formattingSettings.ColorPicker({
        name: "achievedColor",
        displayName: "Achieved Color",
        value: { value: "#0D9488" }
    });

    emptyColor = new formattingSettings.ColorPicker({
        name: "emptyColor",
        displayName: "Empty Dot Color",
        value: { value: "#E2E8F0" }
    });

    fontFamily = new formattingSettings.ItemDropdown({
        name: "fontFamily",
        displayName: "Font Family",
        items: [
            { displayName: "Segoe UI", value: "Segoe UI" },
            { displayName: "Arial", value: "Arial" },
            { displayName: "Helvetica", value: "Helvetica" },
            { displayName: "Georgia", value: "Georgia" },
            { displayName: "Courier New", value: "Courier New" },
            { displayName: "Verdana", value: "Verdana" }
        ],
        value: { displayName: "Segoe UI", value: "Segoe UI" }
    });

    name: string = "gridSettings";
    displayName: string = "Grid Settings";
    slices: Array<FormattingSettingsSlice> = [
        this.dotShape, this.maxDots, this.dotGap, this.showCenterText,
        this.achievedColor, this.emptyColor, this.fontFamily
    ];
}

class LegendSettingsCard extends FormattingSettingsCard {
    showUnitLegend = new formattingSettings.ToggleSwitch({
        name: "showUnitLegend",
        displayName: "Show Unit Legend",
        value: true
    });

    showCategoryLegend = new formattingSettings.ToggleSwitch({
        name: "showCategoryLegend",
        displayName: "Show Category Legend",
        value: true
    });

    name: string = "legendSettings";
    displayName: string = "Legend Settings";
    slices: Array<FormattingSettingsSlice> = [this.showUnitLegend, this.showCategoryLegend];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    gridSettings = new GridSettingsCard();
    legendSettings = new LegendSettingsCard();
    cards = [this.gridSettings, this.legendSettings];
}
