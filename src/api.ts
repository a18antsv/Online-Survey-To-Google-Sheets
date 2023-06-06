import "dotenv/config"
import fetch from "node-fetch";

const {
    BASE_PATH,
} = process.env;

export interface CountryData {
    PKEY: string;
    ev_tier: number;
    is_capi: number;
    country_code: number;
    country_name: string;
    MAIN_QUOTA: number;
    MAIN_COUNT: number;
    MAIN_PERCENT: number;
    QUOTA_COUNT: number;
    QUOTA_PERCENT: number;
    EV_B_QUOTA: string;
    EV_B_COUNT: number;
    EV_B_PERCENT: number;
    TESLA_B_QUOTA: string;
    TESLA_B_COUNT: number;
    TESLA_B_PERCENT: number;
}

export interface MainData {
    QSLABEL: string;
    Total_cnt: number;
    Total_Complete: number;
    Total_Percent: number;
    Ownner_Cnt: number;
    Ownner_Complete: number;
    Ownner_Percent: number;
    Intender_Cnt: number;
    Intender_Complete: number;
    Intender_Percent: number;
}

export interface EvData {
    QSLABEL: string;
    Total_cnt: number;
    Ownner_Complete: number;
    Ownner_Percent: number;
    Intender_Complete: number;
    Intender_Percent: number;
}

export interface BasicData {
    QSLABEL: string;
    Total_Cnt: number;
    Total_Complete: number;
    Total_Percent: number;
}

export interface SelectionData {
    BRAND: string;
    D4_45: number;
    D4A: number;
    D4_123: number;
    D4B: number;
    A3_BRAND: number;
}

export const apiCall = <T>(endpoint: string, params = new Map<string, string>()): Promise<T> => {
    const url = new URL(BASE_PATH + endpoint);

    params.forEach((value, key) => {
        url.searchParams.append(key, value);
    });

    return fetch(url.toString(), {method: "POST"})
        .then(response => {
            if (!response.ok) {
                throw new Error(response.statusText)
            }

            return response.json() as Promise<T>
        });
}
