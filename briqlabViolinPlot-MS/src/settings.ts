"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

export class DistributionSettingsCard extends FormattingSettingsCard {
    kdeBandwidth = new formattingSettings.ItemDropdown({
        name: "kdeBandwidth",
        displayName: "KDE Bandwidth",
        value: { value: "Auto", displayName: "Auto" },
        items: [
            { value: "Auto", displayName: "Auto" },
            { value: "Fine", displayName: "Fine" },
            { value: "Coarse", displayName: "Coarse" }
        ]
    });

    showBoxPlot = new formattingSettings.ToggleSwitch({
        name: "showBoxPlot",
        displayName: "Show Box Plot",
        value: true
    });

    showOutliers = new formattingSettings.ToggleSwitch({
        name: "showOutliers",
        displayName: "Show Outliers",
        value: true
    });

    violinWidth = new formattingSettings.NumUpDown({
        name: "violinWidth",
        displayName: "Violin Width %",
        value: 75,
        options: {
            minValue: { value: 10, type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 100, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    fillOpacity = new formattingSettings.NumUpDown({
        name: "fillOpacity",
        displayName: "Fill Opacity %",
        value: 20,
        options: {
            minValue: { value: 0, type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 100, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    fontFamily = new formattingSettings.FontPicker({
        name: "fontFamily",
        displayName: "Font Family",
        value: "Segoe UI, sans-serif"
    });

    name: string = "distributionSettings";
    displayName: string = "Distribution";
    slices: Array<FormattingSettingsSlice> = [
        this.kdeBandwidth,
        this.showBoxPlot,
        this.showOutliers,
        this.violinWidth,
        this.fillOpacity,
        this.fontFamily
    ];
}

export class StatsSettingsCard extends FormattingSettingsCard {
    showRefLine = new formattingSettings.ItemDropdown({
        name: "showRefLine",
        displayName: "Reference Line",
        value: { value: "None", displayName: "None" },
        items: [
            { value: "None", displayName: "None" },
            { value: "Mean", displayName: "Mean" },
            { value: "Median", displayName: "Median" },
            { value: "Both", displayName: "Both" }
        ]
    });

    refLineColor = new formattingSettings.ColorPicker({
        name: "refLineColor",
        displayName: "Reference Line Color",
        value: { value: "#F97316" }
    });

    name: string = "statsSettings";
    displayName: string = "Statistics";
    slices: Array<FormattingSettingsSlice> = [this.showRefLine, this.refLineColor];
}

export class LabelSettingsCard extends FormattingSettingsCard {
    showLabels = new formattingSettings.ToggleSwitch({
        name: "showLabels",
        displayName: "Show Labels",
        value: true
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Font Size",
        value: 11,
        options: {
            minValue: { value: 9, type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 13, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    name: string = "labelSettings";
    displayName: string = "Labels";
    slices: Array<FormattingSettingsSlice> = [this.showLabels, this.fontSize];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    distributionSettings = new DistributionSettingsCard();
    statsSettings = new StatsSettingsCard();
    labelSettings = new LabelSettingsCard();
    cards = [
        this.distributionSettings,
        this.statsSettings,
        this.labelSettings
    ];
}
