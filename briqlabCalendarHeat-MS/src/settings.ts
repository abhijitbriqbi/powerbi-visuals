"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

class LayoutSettingsCard extends FormattingSettingsCard {
    startOfWeek = new formattingSettings.ItemDropdown({
        name: "startOfWeek",
        displayName: "Start of Week",
        items: [
            { displayName: "Monday", value: "Mon" },
            { displayName: "Sunday", value: "Sun" }
        ],
        value: { displayName: "Monday", value: "Mon" }
    });

    showMonthLabels = new formattingSettings.ToggleSwitch({
        name: "showMonthLabels",
        displayName: "Show Month Labels",
        value: true
    });

    showDayLabels = new formattingSettings.ToggleSwitch({
        name: "showDayLabels",
        displayName: "Show Day Labels",
        value: true
    });

    showSummaryStats = new formattingSettings.ToggleSwitch({
        name: "showSummaryStats",
        displayName: "Show Summary Stats",
        value: true
    });

    highlightWeekends = new formattingSettings.ToggleSwitch({
        name: "highlightWeekends",
        displayName: "Highlight Weekends",
        value: true
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

    name: string = "layoutSettings";
    displayName: string = "Layout Settings";
    slices: Array<FormattingSettingsSlice> = [
        this.startOfWeek, this.showMonthLabels, this.showDayLabels,
        this.showSummaryStats, this.highlightWeekends, this.fontFamily
    ];
}

class ColorSettingsCard extends FormattingSettingsCard {
    colorScale = new formattingSettings.ItemDropdown({
        name: "colorScale",
        displayName: "Color Scale",
        items: [
            { displayName: "Teal", value: "Teal" },
            { displayName: "Blue", value: "Blue" },
            { displayName: "Green", value: "Green" },
            { displayName: "Purple", value: "Purple" },
            { displayName: "Orange", value: "Orange" },
            { displayName: "Custom", value: "Custom" }
        ],
        value: { displayName: "Teal", value: "Teal" }
    });

    lowColor = new formattingSettings.ColorPicker({
        name: "lowColor",
        displayName: "Low Color",
        value: { value: "#CCFBF1" }
    });

    highColor = new formattingSettings.ColorPicker({
        name: "highColor",
        displayName: "High Color",
        value: { value: "#0D9488" }
    });

    nullColor = new formattingSettings.ColorPicker({
        name: "nullColor",
        displayName: "Empty Color",
        value: { value: "#E5E7EB" }
    });

    colorSteps = new formattingSettings.NumUpDown({
        name: "colorSteps",
        displayName: "Color Steps",
        value: 7
    });

    name: string = "colorSettings";
    displayName: string = "Color Settings";
    slices: Array<FormattingSettingsSlice> = [
        this.colorScale, this.lowColor, this.highColor, this.nullColor, this.colorSteps
    ];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    layoutSettings = new LayoutSettingsCard();
    colorSettings = new ColorSettingsCard();
    cards = [this.layoutSettings, this.colorSettings];
}
