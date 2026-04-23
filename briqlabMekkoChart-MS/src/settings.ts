"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard  = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

class ChartSettingsCard extends FormattingSettingsCard {
    barGap = new formattingSettings.NumUpDown({
        name: "barGap", displayName: "Bar Gap (px)", value: 2
    });
    showSegmentLabels = new formattingSettings.ToggleSwitch({
        name: "showSegmentLabels", displayName: "Show Segment Labels", value: true
    });
    labelThreshold = new formattingSettings.NumUpDown({
        name: "labelThreshold", displayName: "Label Threshold %", value: 8
    });
    showXScale = new formattingSettings.ToggleSwitch({
        name: "showXScale", displayName: "Show X Scale", value: true
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

    name: string = "chartSettings";
    displayName: string = "Chart";
    slices: Array<FormattingSettingsSlice> = [
        this.barGap, this.showSegmentLabels, this.labelThreshold,
        this.showXScale, this.fontFamily
    ];
}

class LegendSettingsCard extends FormattingSettingsCard {
    showLegend = new formattingSettings.ToggleSwitch({
        name: "showLegend", displayName: "Show Legend", value: true
    });
    legendPosition = new formattingSettings.ItemDropdown({
        name: "legendPosition",
        displayName: "Legend Position",
        items: [
            { value: "Bottom", displayName: "Bottom" },
            { value: "Top",    displayName: "Top"    }
        ],
        value: { value: "Bottom", displayName: "Bottom" }
    });

    name: string = "legendSettings";
    displayName: string = "Legend";
    slices: Array<FormattingSettingsSlice> = [this.showLegend, this.legendPosition];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    chartSettings  = new ChartSettingsCard();
    legendSettings = new LegendSettingsCard();
    cards = [this.chartSettings, this.legendSettings];
}
