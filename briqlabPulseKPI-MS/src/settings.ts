"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard  = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

class CardSettingsCard extends FormattingSettingsCard {
    cardSize = new formattingSettings.ItemDropdown({
        name: "cardSize",
        displayName: "Card Size",
        items: [
            { value: "compact",  displayName: "Compact"  },
            { value: "standard", displayName: "Standard" },
            { value: "large",    displayName: "Large"    }
        ],
        value: { value: "standard", displayName: "Standard" }
    });

    showPulseRing = new formattingSettings.ToggleSwitch({
        name: "showPulseRing", displayName: "Show Pulse Ring", value: true
    });

    showSparkline = new formattingSettings.ToggleSwitch({
        name: "showSparkline", displayName: "Show Sparkline", value: true
    });

    showInsightText = new formattingSettings.ToggleSwitch({
        name: "showInsightText", displayName: "Show Insight Text", value: true
    });

    showProgressBar = new formattingSettings.ToggleSwitch({
        name: "showProgressBar", displayName: "Show Progress Bar", value: true
    });

    currencyPrefix = new formattingSettings.ItemDropdown({
        name: "currencyPrefix",
        displayName: "Currency Prefix",
        items: [
            { value: "None", displayName: "None" },
            { value: "₹",    displayName: "₹"    },
            { value: "$",    displayName: "$"    },
            { value: "€",    displayName: "€"    },
            { value: "£",    displayName: "£"    }
        ],
        value: { value: "None", displayName: "None" }
    });

    valueFormat = new formattingSettings.ItemDropdown({
        name: "valueFormat",
        displayName: "Value Format",
        items: [
            { value: "auto",  displayName: "Auto (K/M/B)" },
            { value: "K",     displayName: "Thousands (K)" },
            { value: "M",     displayName: "Millions (M)"  },
            { value: "B",     displayName: "Billions (B)"  },
            { value: "exact", displayName: "Exact"         }
        ],
        value: { value: "auto", displayName: "Auto (K/M/B)" }
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

    name: string = "cardSettings";
    displayName: string = "Card";
    slices: Array<FormattingSettingsSlice> = [
        this.cardSize, this.showPulseRing, this.showSparkline,
        this.showInsightText, this.showProgressBar,
        this.currencyPrefix, this.valueFormat, this.fontFamily
    ];
}

class ColorSettingsCard extends FormattingSettingsCard {
    primaryColor = new formattingSettings.ColorPicker({
        name: "primaryColor", displayName: "Primary Color", value: { value: "#0D9488" }
    });

    aboveTargetColor = new formattingSettings.ColorPicker({
        name: "aboveTargetColor", displayName: "Above Target Color", value: { value: "#10B981" }
    });

    belowTargetColor = new formattingSettings.ColorPicker({
        name: "belowTargetColor", displayName: "Below Target Color", value: { value: "#EF4444" }
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor", displayName: "Background Color", value: { value: "#FFFFFF" }
    });

    name: string = "colorSettings";
    displayName: string = "Colors";
    slices: Array<FormattingSettingsSlice> = [
        this.primaryColor, this.aboveTargetColor, this.belowTargetColor, this.backgroundColor
    ];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    cardSettings  = new CardSettingsCard();
    colorSettings = new ColorSettingsCard();
    cards = [this.cardSettings, this.colorSettings];
}
