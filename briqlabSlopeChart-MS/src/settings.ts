"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

export class ChartSettingsCard extends FormattingSettingsCard {
    period1Label = new formattingSettings.TextInput({
        name: "period1Label",
        displayName: "Period 1 Label",
        value: "Period 1",
        placeholder: "Period 1"
    });

    period2Label = new formattingSettings.TextInput({
        name: "period2Label",
        displayName: "Period 2 Label",
        value: "Period 2",
        placeholder: "Period 2"
    });

    colorBy = new formattingSettings.ItemDropdown({
        name: "colorBy",
        displayName: "Color By",
        value: { value: "Direction", displayName: "Direction" },
        items: [
            { value: "Direction", displayName: "Direction" },
            { value: "Category", displayName: "Category" },
            { value: "Uniform", displayName: "Uniform" }
        ]
    });

    lineWidth = new formattingSettings.NumUpDown({
        name: "lineWidth",
        displayName: "Line Width",
        value: 1.5,
        options: {
            minValue: { value: 0.5, type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 8, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    scaleLine = new formattingSettings.ToggleSwitch({
        name: "scaleLine",
        displayName: "Scale Line Width to Change",
        value: false
    });

    showSlopeLabels = new formattingSettings.ToggleSwitch({
        name: "showSlopeLabels",
        displayName: "Show Slope Labels",
        value: false
    });

    showRankChange = new formattingSettings.ToggleSwitch({
        name: "showRankChange",
        displayName: "Show Rank Change",
        value: false
    });

    dotSize = new formattingSettings.NumUpDown({
        name: "dotSize",
        displayName: "Dot Size",
        value: 8,
        options: {
            minValue: { value: 2, type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 20, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    fontFamily = new formattingSettings.FontPicker({
        name: "fontFamily",
        displayName: "Font Family",
        value: "Segoe UI, sans-serif"
    });

    name: string = "chartSettings";
    displayName: string = "Chart";
    slices: Array<FormattingSettingsSlice> = [
        this.period1Label,
        this.period2Label,
        this.colorBy,
        this.lineWidth,
        this.scaleLine,
        this.showSlopeLabels,
        this.showRankChange,
        this.dotSize,
        this.fontFamily
    ];
}

export class SummarySettingsCard extends FormattingSettingsCard {
    showSummary = new formattingSettings.ToggleSwitch({
        name: "showSummary",
        displayName: "Show Summary",
        value: true
    });

    name: string = "summarySettings";
    displayName: string = "Summary";
    slices: Array<FormattingSettingsSlice> = [this.showSummary];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    chartSettings = new ChartSettingsCard();
    summarySettings = new SummarySettingsCard();
    cards = [
        this.chartSettings,
        this.summarySettings
    ];
}
