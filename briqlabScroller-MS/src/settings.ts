"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

class ScrollSettingsCard extends FormattingSettingsCard {
    speed = new formattingSettings.NumUpDown({
        name: "speed",
        displayName: "Scroll Speed (px/sec)",
        value: 60,
        options: {
            minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 },
            maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 400 }
        }
    });

    direction = new formattingSettings.ItemDropdown({
        name: "direction",
        displayName: "Direction",
        value: { value: "left", displayName: "Left" },
        items: [
            { value: "left", displayName: "Left (standard)" },
            { value: "right", displayName: "Right" }
        ]
    });

    pauseOnHover = new formattingSettings.ToggleSwitch({
        name: "pauseOnHover",
        displayName: "Pause on Hover",
        value: true
    });

    gap = new formattingSettings.NumUpDown({
        name: "gap",
        displayName: "Gap Between Items (px)",
        value: 40,
        options: {
            minValue: { type: powerbi.visuals.ValidatorType.Min, value: 10 },
            maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 200 }
        }
    });

    name: string = "scrollSettings";
    displayName: string = "Scroll";
    slices: Array<FormattingSettingsSlice> = [
        this.speed,
        this.direction,
        this.pauseOnHover,
        this.gap
    ];
}

class AppearanceSettingsCard extends FormattingSettingsCard {
    transparentBackground = new formattingSettings.ToggleSwitch({
        name: "transparentBackground",
        displayName: "Transparent Background",
        value: false
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Background Color",
        value: { value: "#374151" }
    });

    textColor = new formattingSettings.ColorPicker({
        name: "textColor",
        displayName: "Text Color",
        value: { value: "#FFFFFF" }
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Font Size (px)",
        value: 14,
        options: {
            minValue: { type: powerbi.visuals.ValidatorType.Min, value: 8 },
            maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 36 }
        }
    });

    fontFamily = new formattingSettings.ItemDropdown({
        name: "fontFamily",
        displayName: "Font",
        value: { value: "Segoe UI", displayName: "Segoe UI" },
        items: [
            { value: "Segoe UI", displayName: "Segoe UI" },
            { value: "Arial", displayName: "Arial" },
            { value: "Calibri", displayName: "Calibri" },
            { value: "Courier New", displayName: "Courier New (monospace)" },
            { value: "Georgia", displayName: "Georgia" },
            { value: "Verdana", displayName: "Verdana" }
        ]
    });

    padding = new formattingSettings.NumUpDown({
        name: "padding",
        displayName: "Vertical Padding (px)",
        value: 8,
        options: {
            minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 },
            maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 40 }
        }
    });

    name: string = "appearanceSettings";
    displayName: string = "Appearance";
    slices: Array<FormattingSettingsSlice> = [
        this.transparentBackground,
        this.backgroundColor,
        this.textColor,
        this.fontSize,
        this.fontFamily,
        this.padding
    ];
}

class IndicatorSettingsCard extends FormattingSettingsCard {
    showArrows = new formattingSettings.ToggleSwitch({
        name: "showArrows",
        displayName: "Show Direction Arrows",
        value: true
    });

    showValue = new formattingSettings.ToggleSwitch({
        name: "showValue",
        displayName: "Show Value",
        value: true
    });

    showChange = new formattingSettings.ToggleSwitch({
        name: "showChange",
        displayName: "Show Change / Delta",
        value: true
    });

    positiveColor = new formattingSettings.ColorPicker({
        name: "positiveColor",
        displayName: "Positive Color",
        value: { value: "#10B981" }
    });

    negativeColor = new formattingSettings.ColorPicker({
        name: "negativeColor",
        displayName: "Negative Color",
        value: { value: "#EF4444" }
    });

    invertColors = new formattingSettings.ToggleSwitch({
        name: "invertColors",
        displayName: "Invert Colors (lower = better)",
        value: false
    });

    name: string = "indicatorSettings";
    displayName: string = "Indicators";
    slices: Array<FormattingSettingsSlice> = [
        this.showArrows,
        this.showValue,
        this.showChange,
        this.positiveColor,
        this.negativeColor,
        this.invertColors
    ];
}

class NumberSettingsCard extends FormattingSettingsCard {
    decimalPlaces = new formattingSettings.NumUpDown({
        name: "decimalPlaces",
        displayName: "Decimal Places",
        value: 2,
        options: {
            minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 },
            maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 6 }
        }
    });

    useKM = new formattingSettings.ToggleSwitch({
        name: "useKM",
        displayName: "Abbreviate (K / M / B)",
        value: false
    });

    prefix = new formattingSettings.TextInput({
        name: "prefix",
        displayName: "Value Prefix",
        placeholder: "e.g. $, €, ₹",
        value: ""
    });

    suffix = new formattingSettings.TextInput({
        name: "suffix",
        displayName: "Value Suffix",
        placeholder: "e.g. USD, kg",
        value: ""
    });

    changePrefix = new formattingSettings.TextInput({
        name: "changePrefix",
        displayName: "Change Prefix",
        placeholder: "e.g. +",
        value: ""
    });

    changeSuffix = new formattingSettings.TextInput({
        name: "changeSuffix",
        displayName: "Change Suffix",
        placeholder: "e.g. %",
        value: "%"
    });

    name: string = "numberSettings";
    displayName: string = "Number Format";
    slices: Array<FormattingSettingsSlice> = [
        this.decimalPlaces,
        this.useKM,
        this.prefix,
        this.suffix,
        this.changePrefix,
        this.changeSuffix
    ];
}

class SeparatorSettingsCard extends FormattingSettingsCard {
    separatorChar = new formattingSettings.TextInput({
        name: "separatorChar",
        displayName: "Separator",
        placeholder: "e.g. | or •",
        value: "  |  "
    });

    separatorColor = new formattingSettings.ColorPicker({
        name: "separatorColor",
        displayName: "Separator Color",
        value: { value: "#374151" }
    });

    name: string = "separatorSettings";
    displayName: string = "Separator";
    slices: Array<FormattingSettingsSlice> = [
        this.separatorChar,
        this.separatorColor
    ];
}

class CustomTextSettingsCard extends FormattingSettingsCard {
    useCustomText = new formattingSettings.ToggleSwitch({
        name: "useCustomText",
        displayName: "Use Custom Text (ignore data)",
        value: false
    });

    customText = new formattingSettings.TextInput({
        name: "customText",
        displayName: "Custom Text",
        placeholder: "Type your scrolling message here...",
        value: ""
    });

    name: string = "customTextSettings";
    displayName: string = "Custom Text";
    slices: Array<FormattingSettingsSlice> = [
        this.useCustomText,
        this.customText
    ];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    scrollSettings = new ScrollSettingsCard();
    appearanceSettings = new AppearanceSettingsCard();
    indicatorSettings = new IndicatorSettingsCard();
    numberSettings = new NumberSettingsCard();
    separatorSettings = new SeparatorSettingsCard();
    customTextSettings = new CustomTextSettingsCard();
    cards = [
        this.scrollSettings,
        this.appearanceSettings,
        this.indicatorSettings,
        this.numberSettings,
        this.separatorSettings,
        this.customTextSettings
    ];
}

