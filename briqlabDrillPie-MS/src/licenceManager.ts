"use strict";
import powerbi from "powerbi-visuals-api";

// Plan ID registered in Microsoft Partner Center — must match EXACTLY.
const BRIQLAB_PRO_PLAN_ID = "BriqlabPro";

let cachedStatus: boolean | null = null;
let lastCheck = 0;
const CACHE_MS = 5 * 60 * 1000; // 5 minutes

/**
 * Check Microsoft IAP licence — no external API calls.
 * All licence verification is done by the Power BI host.
 */
export async function checkMicrosoftLicence(
    host: powerbi.extensibility.visual.IVisualHost
): Promise<boolean> {
    const now = Date.now();
    if (cachedStatus !== null && now - lastCheck < CACHE_MS) return cachedStatus;

    try {
        const lm = (host as any).licenseManager;
        if (!lm) { cachedStatus = false; lastCheck = now; return false; }

        const info = await lm.getAvailableServicePlans();
        if (!info?.plans?.length) { cachedStatus = false; lastCheck = now; return false; }

        cachedStatus = info.plans.some(
            (p: any) => p.spIdentifier === BRIQLAB_PRO_PLAN_ID &&
                        (p.state === "Active" || p.state === "active")
        );
        lastCheck = now;
        return cachedStatus;
    } catch {
        cachedStatus = false;
        lastCheck = now;
        return false;
    }
}

export function resetLicenceCache(): void {
    cachedStatus = null;
    lastCheck = 0;
}
