"use strict";

/**
 * 表示単位（万・億など）。棒グラフ（visuals/barChart/src/unitUtils.ts）と同じ考え方・同じ保存値で、
 * このビジュアルで使う分だけを写した。棒グラフは公開用ソースを単独で切り出すので、
 * 共通化（packages/core へ移す）は公開の段取りと一緒に決める。
 */

import { textMeasurementService } from "powerbi-visuals-utils-formattingutils";

export interface UnitDefinition {
    exponent: number;
    divisor: number;
    /** 解決された語（日本語なら 万・億、英語なら K・M・bn・T） */
    unitWord: string;
    unitWordJa: string;
    /** 標準（英語）の語。1000 の冪以外は語が無いので空 */
    unitWordStd: string;
}

export const UNIT_DEFINITIONS: Record<string, UnitDefinition> = {
    "0": { exponent: 0, divisor: 1, unitWord: "", unitWordJa: "", unitWordStd: "" },
    "1": { exponent: 1, divisor: 10, unitWord: "十", unitWordJa: "十", unitWordStd: "" },
    "2": { exponent: 2, divisor: 100, unitWord: "百", unitWordJa: "百", unitWordStd: "" },
    "3": { exponent: 3, divisor: 1_000, unitWord: "千", unitWordJa: "千", unitWordStd: "K" },
    "4": { exponent: 4, divisor: 10_000, unitWord: "万", unitWordJa: "万", unitWordStd: "" },
    "5": { exponent: 5, divisor: 100_000, unitWord: "十万", unitWordJa: "十万", unitWordStd: "" },
    "6": { exponent: 6, divisor: 1_000_000, unitWord: "百万", unitWordJa: "百万", unitWordStd: "M" },
    "7": { exponent: 7, divisor: 10_000_000, unitWord: "千万", unitWordJa: "千万", unitWordStd: "" },
    "8": { exponent: 8, divisor: 100_000_000, unitWord: "億", unitWordJa: "億", unitWordStd: "" },
    "9": { exponent: 9, divisor: 1_000_000_000, unitWord: "十億", unitWordJa: "十億", unitWordStd: "bn" },
    "10": { exponent: 10, divisor: 10_000_000_000, unitWord: "百億", unitWordJa: "百億", unitWordStd: "" },
    "11": { exponent: 11, divisor: 100_000_000_000, unitWord: "千億", unitWordJa: "千億", unitWordStd: "" },
    "12": { exponent: 12, divisor: 1_000_000_000_000, unitWord: "兆", unitWordJa: "兆", unitWordStd: "T" },
};

/** 自動単位で使う段（小さい順）。日本語は 千・万・百万・億・兆、標準は K・M・bn・T */
const AUTO_KEYS_JA = ["0", "3", "4", "6", "8", "12"];
const AUTO_KEYS_STD = ["0", "3", "6", "9", "12"];

/** precision が "auto" のときの小数桁の上限（formatValue と同じ） */
const AUTO_FRACTION_DIGITS = 2;

function roundedScaledAbs(abs: number, divisor: number, precision: string): number {
    const digits = precision === "auto" ? AUTO_FRACTION_DIGITS : Math.max(0, parseInt(precision, 10) || 0);
    const factor = Math.pow(10, digits);
    return Math.round((abs / divisor) * factor) / factor;
}

/** 小さい順の divisor から abs に合う段を選ぶ。丸めて次の段に届くときは繰り上げる */
function pickUnitIndex(abs: number, divisors: number[], precision: string): number {
    let index = 0;
    for (let i = divisors.length - 1; i >= 0; i--) {
        if (abs >= divisors[i]) {
            index = i;
            break;
        }
    }
    while (
        index + 1 < divisors.length &&
        roundedScaledAbs(abs, divisors[index], precision) >= divisors[index + 1] / divisors[index]
    ) {
        index++;
    }
    return index;
}

export function resolveAutoUnitKey(maxAbsValue: number, notation = "japanese", precision = "auto"): string {
    const keys = notation === "standard" ? AUTO_KEYS_STD : AUTO_KEYS_JA;
    const divisors = keys.map((k) => UNIT_DEFINITIONS[k].divisor);
    return keys[pickUnitIndex(Math.abs(maxAbsValue), divisors, precision)];
}

/** 標準表記では、K・M・bn・T の語が無い桁を 1 つ下の 1000 の冪に読み替える（語が付かずに桁を読み違えるのを防ぐ） */
export function toStandardUnitKey(key: string): string {
    const def = UNIT_DEFINITIONS[key];
    if (!def) return "0";
    return String(Math.floor(def.exponent / 3) * 3);
}

/** 選ばれたキー（auto を含む）から単位を決める */
export function resolveUnit(unitTypeKey: string, maxAbsValue: number, notation = "japanese", precision = "auto"): UnitDefinition {
    const standard = notation === "standard";
    let key = unitTypeKey === "auto" ? resolveAutoUnitKey(maxAbsValue, notation, precision) : unitTypeKey;
    if (!UNIT_DEFINITIONS[key]) key = "0";
    if (standard) key = toStandardUnitKey(key);
    const base = UNIT_DEFINITIONS[key];
    return { ...base, unitWord: standard ? base.unitWordStd : base.unitWordJa };
}

/** 数値を単位で割って書式化する。precision が auto なら小数 2 桁まで */
export function formatValue(value: number, divisor: number, precision: string): string {
    const scaled = value / divisor;
    if (precision === "auto") {
        return scaled.toLocaleString("ja-JP", { maximumFractionDigits: AUTO_FRACTION_DIGITS, minimumFractionDigits: 0 });
    }
    const digits = Math.max(0, parseInt(precision, 10) || 0);
    return scaled.toLocaleString("ja-JP", { minimumFractionDigits: digits, maximumFractionDigits: digits });
}

/** Y 軸の上に出す単位の表記。「(百万円)」。語も追加文字も無ければ空 */
export function unitBadgeOf(unitWord: string, unitText: string): string {
    const composed = `${unitWord || ""}${(unitText || "").trim()}`;
    return composed ? `(${composed})` : "";
}

// --- 文字の幅 ------------------------------------------------------------------

export interface FontSpec {
    family: string;
    /** ピクセル */
    size: number;
    bold?: boolean;
    italic?: boolean;
}

/**
 * 文字の幅。DOM があれば実測し、無ければ（テスト）全角・半角の近似で返す。
 * テストは vitest.setup.ts で textMeasurementService を決定的な近似に差し替えている
 */
export function measureTextWidth(text: string, font: FontSpec): number {
    try {
        if (typeof document !== "undefined") {
            const width = textMeasurementService.measureSvgTextWidth({
                text,
                fontFamily: font.family,
                fontSize: `${font.size}px`,
                fontWeight: font.bold ? "bold" : "normal",
                fontStyle: font.italic ? "italic" : "normal",
            });
            if (width > 0) return width;
        }
    } catch {
        // 近似に落とす
    }
    let w = 0;
    for (let i = 0; i < text.length; i++) {
        const code = text.charCodeAt(i);
        w += code >= 0x20 && code <= 0x7e ? font.size * 0.6 : font.size;
    }
    return font.bold ? w * 1.05 : w;
}

/** WCAG の相対輝度（#RRGGBB・#RGB）。読めなければ null */
export function luminanceOf(color: string): number | null {
    const hex = color.trim().replace("#", "");
    const full = hex.length === 3 ? hex.split("").map((c) => c + c).join("") : hex;
    if (!/^[0-9a-f]{6}$/i.test(full)) return null;
    const value = Number.parseInt(full, 16);
    const channel = (shift: number) => {
        const c = ((value >> shift) & 0xff) / 255;
        return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
    };
    return 0.2126 * channel(16) + 0.7152 * channel(8) + 0.0722 * channel(0);
}

/** 2 つの色のコントラスト比（WCAG、1〜21）。読めない色があれば 21 とみなす */
export function contrastRatio(a: string, b: string): number {
    const la = luminanceOf(a);
    const lb = luminanceOf(b);
    if (la === null || lb === null) return 21;
    return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

/** 背景の色に対して読みやすい文字の色（白か黒）。WCAG の相対輝度で決める（棒グラフと同じ） */
export function contrastingText(backgroundColor: string): string {
    const luminance = luminanceOf(backgroundColor);
    if (luminance === null) return "#252423";
    return luminance > 0.4 ? "#252423" : "#FFFFFF";
}

/** 読める文字の濃い色（白の反対） */
export const DARK_TEXT = "#252423";

/**
 * 下の色 base の上で読める文字色。prefer（既定の色）のコントラスト比が minimum 以上なら prefer、
 * 足りなければ白と濃い灰色のうちコントラストの高いほう
 */
export function readableText(base: string, prefer: string, minimum: number): string {
    if (contrastRatio(prefer, base) >= minimum) return prefer;
    return contrastRatio("#FFFFFF", base) >= contrastRatio(DARK_TEXT, base) ? "#FFFFFF" : DARK_TEXT;
}

/**
 * 幅に収まるように先頭を「…」で切る。上の階層から並べた名前で、一番下のレベルの名前を残すのに使う。
 * 1 文字も入らなければ空
 */
export function truncateStartToWidth(text: string, maxWidth: number, font: FontSpec): string {
    if (measureTextWidth(text, font) <= maxWidth) return text;
    const ellipsis = "…";
    for (let n = 1; n < text.length; n++) {
        const candidate = ellipsis + text.slice(n);
        if (measureTextWidth(candidate, font) <= maxWidth) return candidate;
    }
    return "";
}

/** 幅に収まるように末尾を「…」で切る。1 文字も入らなければ空 */
export function truncateToWidth(text: string, maxWidth: number, font: FontSpec): string {
    if (measureTextWidth(text, font) <= maxWidth) return text;
    const ellipsis = "…";
    for (let n = text.length - 1; n > 0; n--) {
        const candidate = text.slice(0, n) + ellipsis;
        if (measureTextWidth(candidate, font) <= maxWidth) return candidate;
    }
    return "";
}
