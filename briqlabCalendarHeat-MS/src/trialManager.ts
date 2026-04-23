"use strict";

const TRIAL_DAYS  = 4;
const STORAGE_KEY = "briqlab_trial_calendarheat_start";
const PURCHASE_URL =
    "https://appsource.microsoft.com/en-us/product/power-bi-visuals/briqlab.briqlabbriqlabcalendarheat";

export function getTrialStartDate(): Date {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (!stored) {
        const now = new Date();
        localStorage.setItem(STORAGE_KEY, now.toISOString());
        return now;
    }
    return new Date(stored);
}

export function getTrialDaysRemaining(): number {
    const elapsed = Date.now() - getTrialStartDate().getTime();
    return Math.max(0, TRIAL_DAYS - Math.floor(elapsed / 86_400_000));
}

export function isTrialExpired(): boolean {
    return getTrialDaysRemaining() === 0;
}

export function getPurchaseUrl(): string { return PURCHASE_URL; }

export function getButtonText(): string { return "Buy on AppSource →"; }
