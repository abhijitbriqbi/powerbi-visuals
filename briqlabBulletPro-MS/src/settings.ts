"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

export class ChartSettingsCard extends FormattingSettingsCard {
    rowHeight = new formattingSettings.NumUpDown({
        name: "rowHeight",
        displayName: "Row Height",
        value: 36,
        options: {
            minValue: { value: 24, type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 60, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    bgBarHeight = new formattingSettings.NumUpDown({
        name: "bgBarHeight",
        displayName: "Background Bar Height %",
        value: 75,
        options: {
            minValue: { value: 10, type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 100, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    perfBarHeight = new formattingSettings.NumUpDown({
        name: "perfBarHeight",
        displayName: "Performance Bar Height %",
        value: 40,
        options: {
            minValue: { value: 10, type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 100, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    showComparative = new formattingSettings.ToggleSwitch({
        name: "showComparative",
        displayName: "Show Comparative Marker",
        value: true
    });

    showSummary = new formattingSettings.ToggleSwitch({
        name: "showSummary",
        displayName: "Show Summary Header",
        value: true
    });

    fontFamily = new formattingSettings.FontPicker({
        name: "fontFamily",
        displayName: "Font Family",
        value: "Segoe UI, sans-serif"
    });

    name: string = "chartSettings";
    displayName: string = "Chart";
    slices: Array<FormattingSettingsSlice> = [
        this.rowHeight,
        this.bgBarHeight,
        this.perfBarHeight,
        this.showComparative,
        this.showSummary,
        this.fontFamily
    ];
}

export class ColorSettingsCard extends FormattingSettingsCard {
    redZoneColor = new formattingSettings.ColorPicker({
        name: "redZoneColor",
        displayName: "Red Zone Color",
        value: { value: "#FEE2E2" }
    });

    amberZoneColor = new formattingSettings.ColorPicker({
        name: "amberZoneColor",
        displayName: "Amber Zone Color",
        value: { value: "#FEF3C7" }
    });

    greenZoneColor = new formattingSettings.ColorPicker({
        name: "greenZoneColor",
        displayName: "Green Zone Color",
        value: { value: "#DCFCE7" }
    });

    targetColor = new formattingSettings.ColorPicker({
        name: "targetColor",
        displayName: "Target Marker Color",
        value: { value: "#374151" }
    });

    comparativeColor = new formattingSettings.ColorPicker({
        name: "comparativeColor",
        displayName: "Comparative Marker Color",
        value: { value: "#374151" }
    });

    name: string = "colorSettings";
    displayName: string = "Colors";
    slices: Array<FormattingSettingsSlice> = [
        this.redZoneColor,
        this.amberZoneColor,
        this.greenZoneColor,
        this.targetColor,
        this.comparativeColor
    ];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    chartSettings = new ChartSettingsCard();
    colorSettings = new ColorSettingsCard();
    cards = [
        this.chartSettings,
        this.colorSettings
    ];
}
