"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

class DisplaySettingsCard extends FormattingSettingsCard {
    valueFontSize = new formattingSettings.NumUpDown({
        name: "valueFontSize",
        displayName: "Value Font Size",
        value: 32
    });

    labelFontSize = new formattingSettings.NumUpDown({
        name: "labelFontSize",
        displayName: "Label Font Size",
        value: 12
    });

    primaryColor = new formattingSettings.ColorPicker({
        name: "primaryColor",
        displayName: "Primary Color",
        value: { value: "#0D9488" }
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Background Color",
        value: { value: "#FFFFFF" }
    });

    showBorder = new formattingSettings.ToggleSwitch({
        name: "showBorder",
        displayName: "Show Border",
        value: false
    });

    borderColor = new formattingSettings.ColorPicker({
        name: "borderColor",
        displayName: "Border Color",
        value: { value: "#E2E8F0" }
    });

    borderRadius = new formattingSettings.NumUpDown({
        name: "borderRadius",
        displayName: "Border Radius",
        value: 8
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

    name: string = "displaySettings";
    displayName: string = "Display";
    slices: Array<FormattingSettingsSlice> = [
        this.valueFontSize,
        this.labelFontSize,
        this.fontFamily,
        this.primaryColor,
        this.backgroundColor,
        this.showBorder,
        this.borderColor,
        this.borderRadius
    ];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    displaySettings = new DisplaySettingsCard();
    cards = [this.displaySettings];
}
