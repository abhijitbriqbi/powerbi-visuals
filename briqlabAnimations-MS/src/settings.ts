"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard  = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

class AnimationSettingsCard extends FormattingSettingsCard {
    duration = new formattingSettings.NumUpDown({
        name: "duration",
        displayName: "Animation Duration (ms)",
        value: 1500
    });

    easing = new formattingSettings.ItemDropdown({
        name: "easing",
        displayName: "Easing",
        items: [
            { displayName: "Ease Out",    value: "easeOut"    },
            { displayName: "Linear",      value: "linear"     },
            { displayName: "Ease In Out", value: "easeInOut"  },
            { displayName: "Bounce",      value: "bounce"     }
        ],
        value: { displayName: "Ease Out", value: "easeOut" }
    });

    autoPlay = new formattingSettings.ToggleSwitch({
        name: "autoPlay",
        displayName: "Auto-play on Load",
        value: true
    });

    showProgressBar = new formattingSettings.ToggleSwitch({
        name: "showProgressBar",
        displayName: "Show Progress Bar",
        value: true
    });

    name: string = "animationSettings";
    displayName: string = "Animation";
    slices: Array<FormattingSettingsSlice> = [
        this.duration, this.easing, this.autoPlay, this.showProgressBar
    ];
}

class StyleSettingsCard extends FormattingSettingsCard {
    primaryColor = new formattingSettings.ColorPicker({
        name: "primaryColor",
        displayName: "Primary Color",
        value: { value: "#0D9488" }
    });

    secondaryColor = new formattingSettings.ColorPicker({
        name: "secondaryColor",
        displayName: "Secondary Color",
        value: { value: "#E2E8F0" }
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Card Background",
        value: { value: "#FFFFFF" }
    });

    fontFamily = new formattingSettings.ItemDropdown({
        name: "fontFamily",
        displayName: "Font Family",
        items: [
            { displayName: "Segoe UI",   value: "Segoe UI"   },
            { displayName: "Arial",      value: "Arial"      },
            { displayName: "Helvetica",  value: "Helvetica"  },
            { displayName: "Georgia",    value: "Georgia"    },
            { displayName: "Verdana",    value: "Verdana"    }
        ],
        value: { displayName: "Segoe UI", value: "Segoe UI" }
    });

    valueFontSize = new formattingSettings.NumUpDown({
        name: "valueFontSize",
        displayName: "Value Font Size",
        value: 36
    });

    labelFontSize = new formattingSettings.NumUpDown({
        name: "labelFontSize",
        displayName: "Label Font Size",
        value: 13
    });

    cardLayout = new formattingSettings.ItemDropdown({
        name: "cardLayout",
        displayName: "Layout",
        items: [
            { displayName: "Auto Grid", value: "grid"   },
            { displayName: "Row",       value: "row"    },
            { displayName: "Column",    value: "column" }
        ],
        value: { displayName: "Auto Grid", value: "grid" }
    });

    name: string = "styleSettings";
    displayName: string = "Style";
    slices: Array<FormattingSettingsSlice> = [
        this.primaryColor, this.secondaryColor, this.backgroundColor,
        this.fontFamily, this.valueFontSize, this.labelFontSize, this.cardLayout
    ];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    animationSettings = new AnimationSettingsCard();
    styleSettings     = new StyleSettingsCard();
    cards = [this.animationSettings, this.styleSettings];
}
