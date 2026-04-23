"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard  = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

class RingSettingsCard extends FormattingSettingsCard {
    trackWidth = new formattingSettings.NumUpDown({
        name: "trackWidth", displayName: "Track Width (px)", value: 18
    });
    ringGap = new formattingSettings.NumUpDown({
        name: "ringGap", displayName: "Ring Gap (px)", value: 6
    });
    maxRings = new formattingSettings.NumUpDown({
        name: "maxRings", displayName: "Max Rings", value: 6
    });
    autoColor = new formattingSettings.ToggleSwitch({
        name: "autoColor", displayName: "Auto Color by Achievement", value: true
    });
    showMilestones = new formattingSettings.ToggleSwitch({
        name: "showMilestones", displayName: "Show Milestones", value: false
    });
    fontFamily = new formattingSettings.ItemDropdown({
        name: "fontFamily",
        displayName: "Font Family",
        items: [
            { value: "Segoe UI",    displayName: "Segoe UI"    },
            { value: "Arial",       displayName: "Arial"        },
            { value: "Calibri",     displayName: "Calibri"      },
            { value: "Verdana",     displayName: "Verdana"      },
            { value: "Georgia",     displayName: "Georgia"      },
            { value: "Courier New", displayName: "Courier New"  }
        ],
        value: { value: "Segoe UI", displayName: "Segoe UI" }
    });

    name: string = "ringSettings";
    displayName: string = "Rings";
    slices: Array<FormattingSettingsSlice> = [
        this.trackWidth, this.ringGap, this.maxRings,
        this.autoColor, this.showMilestones, this.fontFamily
    ];
}

class LabelSettingsCard extends FormattingSettingsCard {
    showLabels = new formattingSettings.ToggleSwitch({
        name: "showLabels", displayName: "Show Labels", value: true
    });
    labelFontSize = new formattingSettings.NumUpDown({
        name: "labelFontSize", displayName: "Label Font Size", value: 11
    });

    name: string = "labelSettings";
    displayName: string = "Labels";
    slices: Array<FormattingSettingsSlice> = [this.showLabels, this.labelFontSize];
}

class CenterSettingsCard extends FormattingSettingsCard {
    showCenter = new formattingSettings.ToggleSwitch({
        name: "showCenter", displayName: "Show Center", value: true
    });
    summaryMetric = new formattingSettings.ItemDropdown({
        name: "summaryMetric",
        displayName: "Summary Metric",
        items: [
            { value: "Average %", displayName: "Average %" },
            { value: "Count Met", displayName: "Count Met" }
        ],
        value: { value: "Average %", displayName: "Average %" }
    });

    name: string = "centerSettings";
    displayName: string = "Center";
    slices: Array<FormattingSettingsSlice> = [this.showCenter, this.summaryMetric];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    ringSettings   = new RingSettingsCard();
    labelSettings  = new LabelSettingsCard();
    centerSettings = new CenterSettingsCard();
    cards = [this.ringSettings, this.labelSettings, this.centerSettings];
}
