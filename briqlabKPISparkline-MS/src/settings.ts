"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard  = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

// ── Value ──────────────────────────────────────────────────────────────────────
class ValueSettingsCard extends FormattingSettingsCard {
    currencyPrefix = new formattingSettings.ItemDropdown({
        name: "currencyPrefix",
        displayName: "Currency Prefix",
        items: [
            { value: "None", displayName: "None" },
            { value: "Rs",   displayName: "Rs"   },
            { value: "$",    displayName: "$"    },
            { value: "EUR",  displayName: "EUR"  },
            { value: "GBP",  displayName: "GBP"  }
        ],
        value: { value: "None", displayName: "None" }
    });

    valueSize = new formattingSettings.NumUpDown({
        name: "valueSize",
        displayName: "Value Font Size",
        value: 42
    });

    showTrend = new formattingSettings.ToggleSwitch({
        name: "showTrend",
        displayName: "Show Trend Indicator",
        value: true
    });

    countUpAnimation = new formattingSettings.ToggleSwitch({
        name: "countUpAnimation",
        displayName: "Count-Up Animation",
        value: true
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

    name: string        = "valueSettings";
    displayName: string = "Value";
    slices: Array<FormattingSettingsSlice> = [
        this.currencyPrefix,
        this.valueSize,
        this.fontFamily,
        this.showTrend,
        this.countUpAnimation
    ];
}

// ── Target ─────────────────────────────────────────────────────────────────────
class TargetSettingsCard extends FormattingSettingsCard {
    showTargetBar = new formattingSettings.ToggleSwitch({
        name: "showTargetBar",
        displayName: "Show Target Bar",
        value: true
    });

    targetLabel = new formattingSettings.TextInput({
        name: "targetLabel",
        displayName: "Target Label",
        placeholder: "vs Target",
        value: "vs Target"
    });

    aboveTargetColor = new formattingSettings.ColorPicker({
        name: "aboveTargetColor",
        displayName: "Above Target Color",
        value: { value: "#10B981" }
    });

    belowTargetColor = new formattingSettings.ColorPicker({
        name: "belowTargetColor",
        displayName: "Below Target Color",
        value: { value: "#EF4444" }
    });

    name: string        = "targetSettings";
    displayName: string = "Target";
    slices: Array<FormattingSettingsSlice> = [
        this.showTargetBar,
        this.targetLabel,
        this.aboveTargetColor,
        this.belowTargetColor
    ];
}

// ── Sparkline ──────────────────────────────────────────────────────────────────
class SparklineSettingsCard extends FormattingSettingsCard {
    showSparkline = new formattingSettings.ToggleSwitch({
        name: "showSparkline",
        displayName: "Show Sparkline",
        value: true
    });

    sparklineType = new formattingSettings.ItemDropdown({
        name: "sparklineType",
        displayName: "Sparkline Type",
        items: [
            { value: "Area", displayName: "Area" },
            { value: "Line", displayName: "Line" },
            { value: "Bar",  displayName: "Bar"  }
        ],
        value: { value: "Area", displayName: "Area" }
    });

    sparklineHeight = new formattingSettings.NumUpDown({
        name: "sparklineHeight",
        displayName: "Sparkline Height %",
        value: 35
    });

    sparklineColor = new formattingSettings.ItemDropdown({
        name: "sparklineColor",
        displayName: "Sparkline Color",
        items: [
            { value: "Auto",   displayName: "Auto (matches status)" },
            { value: "Manual", displayName: "Manual"                }
        ],
        value: { value: "Auto", displayName: "Auto (matches status)" }
    });

    manualSparklineColor = new formattingSettings.ColorPicker({
        name: "manualSparklineColor",
        displayName: "Manual Color",
        value: { value: "#0D9488" }
    });

    name: string        = "sparklineSettings";
    displayName: string = "Sparkline";
    slices: Array<FormattingSettingsSlice> = [
        this.showSparkline,
        this.sparklineType,
        this.sparklineHeight,
        this.sparklineColor,
        this.manualSparklineColor
    ];
}

// ── Card Style ─────────────────────────────────────────────────────────────────
class CardStyleCard extends FormattingSettingsCard {
    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Background Color",
        value: { value: "#FFFFFF" }
    });

    accentColor = new formattingSettings.ItemDropdown({
        name: "accentColor",
        displayName: "Accent Stripe",
        items: [
            { value: "Auto",   displayName: "Auto (matches status)" },
            { value: "Manual", displayName: "Manual"                }
        ],
        value: { value: "Auto", displayName: "Auto (matches status)" }
    });

    manualAccentColor = new formattingSettings.ColorPicker({
        name: "manualAccentColor",
        displayName: "Manual Accent Color",
        value: { value: "#0D9488" }
    });

    borderRadius = new formattingSettings.NumUpDown({
        name: "borderRadius",
        displayName: "Border Radius",
        value: 8
    });

    showShadow = new formattingSettings.ToggleSwitch({
        name: "showShadow",
        displayName: "Show Shadow",
        value: true
    });

    name: string        = "cardStyle";
    displayName: string = "Card Style";
    slices: Array<FormattingSettingsSlice> = [
        this.backgroundColor,
        this.accentColor,
        this.manualAccentColor,
        this.borderRadius,
        this.showShadow
    ];
}

// ── Briqlab Pro ────────────────────────────────────────────────────────────────

// ── Model ──────────────────────────────────────────────────────────────────────
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    valueSettings    = new ValueSettingsCard();
    targetSettings   = new TargetSettingsCard();
    sparklineSettings = new SparklineSettingsCard();
    cardStyle        = new CardStyleCard();
    cards = [
        this.valueSettings,
        this.targetSettings,
        this.sparklineSettings,
        this.cardStyle
    ];
}
