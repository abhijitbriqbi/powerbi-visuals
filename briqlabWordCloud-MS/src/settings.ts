"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

class WordSettingsCard extends FormattingSettingsCard {
    minFontSize = new formattingSettings.NumUpDown({
        name: "minFontSize",
        displayName: "Min Font Size",
        value: 10
    });

    maxFontSize = new formattingSettings.NumUpDown({
        name: "maxFontSize",
        displayName: "Max Font Size",
        value: 60
    });

    maxWords = new formattingSettings.NumUpDown({
        name: "maxWords",
        displayName: "Max Words",
        value: 100
    });

    rotations = new formattingSettings.ItemDropdown({
        name: "rotations",
        displayName: "Word Rotation",
        items: [
            { displayName: "None", value: "None" },
            { displayName: "Slight", value: "Slight" },
            { displayName: "Full", value: "Full" }
        ],
        value: { displayName: "None", value: "None" }
    });

    wordPadding = new formattingSettings.NumUpDown({
        name: "wordPadding",
        displayName: "Word Padding",
        value: 4
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

    name: string = "wordSettings";
    displayName: string = "Word Settings";
    slices: Array<FormattingSettingsSlice> = [
        this.minFontSize, this.maxFontSize, this.maxWords,
        this.rotations, this.wordPadding, this.fontFamily
    ];
}

class ColorSettingsCard extends FormattingSettingsCard {
    colorMode = new formattingSettings.ItemDropdown({
        name: "colorMode",
        displayName: "Color Mode",
        items: [
            { displayName: "Sentiment", value: "Sentiment" },
            { displayName: "Rank", value: "Rank" },
            { displayName: "Uniform", value: "Uniform" }
        ],
        value: { displayName: "Rank", value: "Rank" }
    });

    uniformColor = new formattingSettings.ColorPicker({
        name: "uniformColor",
        displayName: "Uniform Color",
        value: { value: "#0D9488" }
    });

    positiveColor = new formattingSettings.ColorPicker({
        name: "positiveColor",
        displayName: "Positive Color",
        value: { value: "#0D9488" }
    });

    negativeColor = new formattingSettings.ColorPicker({
        name: "negativeColor",
        displayName: "Negative Color",
        value: { value: "#EF4444" }
    });

    neutralColor = new formattingSettings.ColorPicker({
        name: "neutralColor",
        displayName: "Neutral Color",
        value: { value: "#E2E8F0" }
    });

    name: string = "colorSettings";
    displayName: string = "Color Settings";
    slices: Array<FormattingSettingsSlice> = [
        this.colorMode, this.uniformColor, this.positiveColor, this.negativeColor, this.neutralColor
    ];
}

class ExclusionSettingsCard extends FormattingSettingsCard {
    stopWords = new formattingSettings.TextInput({
        name: "stopWords",
        displayName: "Stop Words",
        value: "",
        placeholder: "the, a, an, is, ..."
    });

    caseSensitive = new formattingSettings.ToggleSwitch({
        name: "caseSensitive",
        displayName: "Case Sensitive",
        value: false
    });

    name: string = "exclusionSettings";
    displayName: string = "Exclusion Settings";
    slices: Array<FormattingSettingsSlice> = [this.stopWords, this.caseSensitive];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    wordSettings = new WordSettingsCard();
    colorSettings = new ColorSettingsCard();
    exclusionSettings = new ExclusionSettingsCard();
    cards = [this.wordSettings, this.colorSettings, this.exclusionSettings];
}
