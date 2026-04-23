"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard  = formattingSettings.SimpleCard;
import FormattingSettingsModel = formattingSettings.Model;

const fontFamilyOptions = [
    { value: "Segoe UI", displayName: "Segoe UI" },
    { value: "Arial",    displayName: "Arial" },
    { value: "Calibri",  displayName: "Calibri" },
    { value: "Verdana",  displayName: "Verdana" }
];

// ── Search Settings Card ─────────────────────────────────────────────────────

class SearchSettingsCard extends FormattingSettingsCard {
    placeholder = new formattingSettings.TextInput({
        name:         "placeholder",
        displayName:  "Placeholder text",
        description:  "Text shown in the search box when empty",
        value:        "Search…",
        placeholder:  "Search…"
    });

    maxResults = new formattingSettings.NumUpDown({
        name:        "maxResults",
        displayName: "Max results",
        description: "Maximum number of results to display",
        value:       10
    });

    showResultCount = new formattingSettings.ToggleSwitch({
        name:        "showResultCount",
        displayName: "Show result count",
        description: "Display the number of matching results above the list",
        value:       true
    });

    caseSensitive = new formattingSettings.ToggleSwitch({
        name:        "caseSensitive",
        displayName: "Case sensitive",
        description: "Enable case-sensitive matching",
        value:       false
    });

    fontFamily = new formattingSettings.ItemDropdown({
        name:        "fontFamily",
        displayName: "Font family",
        description: "Font used for the visual",
        items:       fontFamilyOptions,
        value:       fontFamilyOptions[0]
    });

    fontSize = new formattingSettings.NumUpDown({
        name:        "fontSize",
        displayName: "Font size",
        description: "Font size in pixels",
        value:       12
    });

    name:        string = "searchSettings";
    displayName: string = "Search Settings";
    slices = [
        this.placeholder,
        this.maxResults,
        this.showResultCount,
        this.caseSensitive,
        this.fontFamily,
        this.fontSize
    ];
}

// ── Color Settings Card ──────────────────────────────────────────────────────

class ColorSettingsCard extends FormattingSettingsCard {
    accentColor = new formattingSettings.ColorPicker({
        name:        "accentColor",
        displayName: "Accent color",
        description: "Highlight and selection color",
        value:       { value: "#0D9488" }
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name:        "backgroundColor",
        displayName: "Background color",
        description: "Visual background color",
        value:       { value: "#FFFFFF" }
    });

    textColor = new formattingSettings.ColorPicker({
        name:        "textColor",
        displayName: "Text color",
        description: "Primary text color",
        value:       { value: "#374151" }
    });

    borderColor = new formattingSettings.ColorPicker({
        name:        "borderColor",
        displayName: "Border color",
        description: "Input box border color",
        value:       { value: "#E2E8F0" }
    });

    name:        string = "colorSettings";
    displayName: string = "Colors";
    slices = [
        this.accentColor,
        this.backgroundColor,
        this.textColor,
        this.borderColor
    ];
}

// ── Pro Settings Card ────────────────────────────────────────────────────────

// ── Root Model ───────────────────────────────────────────────────────────────

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    searchSettings = new SearchSettingsCard();
    colorSettings  = new ColorSettingsCard();
    cards = [
        this.searchSettings,
        this.colorSettings
    ];
}
