"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard  = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

class RadarSettingsCard extends FormattingSettingsCard {
    gridRings = new formattingSettings.NumUpDown({
        name: "gridRings", displayName: "Grid Rings", value: 5
    });
    maxEntities = new formattingSettings.NumUpDown({
        name: "maxEntities", displayName: "Max Entities", value: 5
    });
    fillOpacity = new formattingSettings.NumUpDown({
        name: "fillOpacity", displayName: "Fill Opacity %", value: 12
    });
    showDots = new formattingSettings.ToggleSwitch({
        name: "showDots", displayName: "Show Dots", value: true
    });
    showBenchmark = new formattingSettings.ToggleSwitch({
        name: "showBenchmark", displayName: "Show Benchmark", value: true
    });
    showScoreCards = new formattingSettings.ToggleSwitch({
        name: "showScoreCards", displayName: "Show Score Cards", value: true
    });

    name: string = "radarSettings";
    displayName: string = "Radar";
    slices: Array<FormattingSettingsSlice> = [
        this.gridRings, this.maxEntities, this.fillOpacity,
        this.showDots, this.showBenchmark, this.showScoreCards
    ];
}

class LabelSettingsCard extends FormattingSettingsCard {
    axisFontSize = new formattingSettings.NumUpDown({
        name: "axisFontSize", displayName: "Axis Font Size", value: 11
    });
    showAxisValues = new formattingSettings.ToggleSwitch({
        name: "showAxisValues", displayName: "Show Axis Values", value: true
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

    name: string = "labelSettings";
    displayName: string = "Labels";
    slices: Array<FormattingSettingsSlice> = [
        this.axisFontSize, this.showAxisValues, this.fontFamily
    ];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    radarSettings = new RadarSettingsCard();
    labelSettings = new LabelSettingsCard();
    cards = [this.radarSettings, this.labelSettings];
}
