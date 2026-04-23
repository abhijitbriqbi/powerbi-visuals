"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard  = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

class LayoutSettingsCard extends FormattingSettingsCard {
    nodeWidth = new formattingSettings.NumUpDown({
        name: "nodeWidth", displayName: "Node Width (px)", value: 20
    });
    nodeGap = new formattingSettings.NumUpDown({
        name: "nodeGap", displayName: "Node Gap (px)", value: 12
    });
    flowOpacity = new formattingSettings.NumUpDown({
        name: "flowOpacity", displayName: "Flow Opacity %", value: 60
    });
    showFlowLabels = new formattingSettings.ToggleSwitch({
        name: "showFlowLabels", displayName: "Show Flow Labels", value: true
    });
    minLabelPct = new formattingSettings.NumUpDown({
        name: "minLabelPct", displayName: "Min Label Threshold %", value: 5
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

    name: string = "layoutSettings";
    displayName: string = "Layout";
    slices: Array<FormattingSettingsSlice> = [
        this.nodeWidth, this.nodeGap, this.flowOpacity,
        this.showFlowLabels, this.minLabelPct, this.fontFamily
    ];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    layoutSettings = new LayoutSettingsCard();
    cards = [this.layoutSettings];
}
