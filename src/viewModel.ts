"use strict";

/**
 * 棒の並びを組み立てる。DataView の読み替えは data.ts。
 *
 * 棒は「合計」と「増減」の 2 種類だけ（docs/waterfall.md の「設計」）。
 * - 合計の棒：軸から立つ。起点・イベントの値・小計・最後の合計
 * - 増減の棒：直前までの累計から浮く。項目ごとの差や値（まとめたものは「その他」）。系列があれば系列の色で積む
 *
 * 形は入っている枠から決める。
 * - 比べる（形 1）：値にメジャーが 2 つ以上、またはイベント＋値 1 つ。イベントの合計 → 差 → イベントの合計
 * - 合流（形 2）：値 1 つ。項目の値 → 最後の合計
 * - 起点（形 3・5）：起点＋値 1 つ。起点 → 項目の値 → 最後の合計（起点＋増減）
 * - 小計（形 4）：どの形にも付く。小計の列の値が同じ項目をまとめ、その最後に累計の棒を立てる
 *
 * どの形でも、増減の合計 ＝ 最後の合計 − 最初の合計 になる（検算できる形）。
 */

import powerbi from "powerbi-visuals-api";
import { scaleLinear } from "d3-scale";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import {
    BREAK_STYLES,
    BreakStyle,
    CONNECT_MODES,
    DEFAULT_FONT_FAMILY,
    DEFAULT_LABEL_BACKGROUND,
    DEFAULT_LABEL_BACKGROUND_TRANSPARENCY,
    DEFAULT_OTHERS_LABEL,
    DEFAULT_TOTAL_LABEL,
    DataDrivenFormat,
    LABEL_ORIENTATIONS,
    LABEL_POSITIONS,
    LEGEND_POSITIONS,
    TITLE_STYLES,
    LINE_STYLES,
    LegendPosition,
    ORDERS,
    ORIENTATIONS,
    ORIGIN_MODES,
    Orientation,
    RATE_TYPES,
    VisualFormattingSettingsModel,
    dropdownValue,
} from "./settings";
import { Formatter, LEVEL_SEPARATOR, ParsedItem, Parsed, parse } from "./data";
import { formatValue, resolveUnit, unitBadgeOf } from "./unitUtils";

import DataView = powerbi.DataView;
import DataViewObject = powerbi.DataViewObject;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISandboxExtendedColorPalette = powerbi.extensibility.ISandboxExtendedColorPalette;
import ISelectionId = powerbi.visuals.ISelectionId;

export { BLANK_TEXT, LEVEL_SEPARATOR } from "./data";

export const SHAPES = { compare: "compare", merge: "merge", origin: "origin" } as const;
export type Shape = (typeof SHAPES)[keyof typeof SHAPES];

/**
 * 棒の種類。total = 合計（軸から立つ）、increase・decrease・others = 増減（累計から浮く）、
 * subtotal = 比べる形の小計（その区切りの差の合計。浮いた棒で、累計は進めない）
 */
export const BAR_KINDS = {
    total: "total",
    increase: "increase",
    decrease: "decrease",
    others: "others",
    subtotal: "subtotal",
} as const;
export type BarKind = (typeof BAR_KINDS)[keyof typeof BAR_KINDS];

/** 合計の棒の中身。event = イベントの値、origin = 起点、subtotal = 小計（比べる形では区切りの差）、final = 最後の合計 */
export type TotalRole = "event" | "origin" | "subtotal" | "final";

export interface TooltipItem {
    displayName: string;
    value: string;
}

/** ツールヒントを出すもの（棒、目標の線）。選ぶ ID が無ければ空 */
export interface TooltipTarget {
    tooltip: TooltipItem[];
    selectionIds: ISelectionId[];
}

/** 棒の中の 1 区画。系列が無ければ棒 1 本に 1 つ */
export interface Segment {
    /** 積み始め（正の区画は下端、負の区画は上端） */
    from: number;
    /** 積み終わり */
    to: number;
    color: string;
    /** 系列の添字。系列が無ければ null */
    series: number | null;
    /** ハイライトの該当分の終わり（from から積む）。ハイライトが無い・該当しなければ null */
    highlightTo: number | null;
    selectionIds: ISelectionId[];
}

export interface Bar {
    /** React の key。形・区切り・項目で一意 */
    key: string;
    kind: BarKind;
    /** 合計の棒の中身。増減の棒は null */
    totalRole: TotalRole | null;
    /** 軸の名前。項目は一番下のレベルの名前 */
    label: string;
    /** 「すべて展開」のときの上の階層の名前（" / " でつなぐ）。無ければ空 */
    parentLabel: string;
    /** 読み上げ・ツールヒント用の名前（上の階層から） */
    fullLabel: string;
    /** 累計の始まり。合計の棒は 0 */
    from: number;
    /** 累計の終わり。合計の棒は値 */
    to: number;
    /** 増減の棒は増減、合計の棒は値 */
    value: number;
    /** データ ラベル（表示単位で書式化済み） */
    labelText: string;
    /** 率（書式で選んだとき）。無ければ空 */
    rateText: string;
    /** 棒の色（系列で積むときは区画ごとの色が segments にある） */
    color: string;
    segments: Segment[];
    /** ハイライトがあって、この棒が全部は該当しない（薄く描いて該当分を重ねる） */
    dimmed: boolean;
    /** 選択に使う ID。項目は行ごと。選べない棒は空 */
    selectionIds: ISelectionId[];
    tooltip: TooltipItem[];
}

export interface Tick {
    value: number;
    label: string;
}

export interface AxisInfo {
    min: number;
    max: number;
    ticks: Tick[];
    /**
     * 合計の棒を切ったか。bottom = 軸の最小が 0 より大きい（正の合計の棒の根元を切った）、
     * top = 軸の最大が 0 より小さい（負の合計の棒の根元を切った）、none = 0 が軸に入っている
     */
    cut: "none" | "bottom" | "top";
}

export interface LegendItem {
    name: string;
    color: string;
    /** 系列を選ぶ ID。増加・減少などの項目は null */
    selectionId: ISelectionId | null;
    /** 系列か、棒の種類（増加・減少・合計・区切りの差・その他）か、目標の線 */
    kind: "series" | BarKind | "target";
    /** 目標の線の印（線の太さと種類。グラフの線と同じ種類で、凡例では太さを 3px までにする）。棒の項目は無し */
    line?: { width: number; dash: string | null };
    /** 印を左右に分けて塗る色（区切りの差：増加・減少）。1 つなら 1 色 */
    split?: string[];
    /** 印の濃さ（区切りの差は透かす）。無ければ 1 */
    opacity?: number;
}

/** 凡例の印の線の太さの上限（太い線は凡例の行に収まらない） */
const LEGEND_LINE_MAX = 3;

/** 凡例の名前（標準と同じ「増加・減少・合計」） */
export const LEGEND_NAMES = { increase: "増加", decrease: "減少", total: "合計", subtotal: "区切りの差" } as const;

export interface TargetLine {
    value: number;
    /** 目標は目標のメジャーの表示名、定数線は書式の名前（空もある）。凡例に出す名前 */
    name: string;
    /** グラフの中に出す名前と値（書式で選ぶ。目標は既定で出さない）。空なら出さない */
    label: string;
    /** 線にマウスを当てたときのツールヒント（名前と、モデルの書式の値）。空なら出さない */
    tooltip: TooltipItem[];
    color: string;
    width: number;
    dash: string | null;
}

/** 文字の見た目。size はポイント（描画でピクセルに直す）。color が空なら自動 */
export interface TextStyle {
    family: string;
    size: number;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    color: string;
}

export interface AxisTitle {
    text: string;
    font: TextStyle;
}

export interface ViewStyle {
    orientation: Orientation;
    /** 軸を切った印の形 */
    breakStyle: BreakStyle;
    connectors: { show: boolean; color: string; width: number; dash: string | null };
    dataLabels: {
        show: boolean;
        /** color が空なら自動（棒の外は灰色、棒の中は棒の色に合わせて白か黒） */
        font: TextStyle;
        /** 合計・小計・区切りの差の棒のラベル。color が空なら値のラベルと同じ（それも空なら自動） */
        totalFont: TextStyle;
        orientation: "horizontal" | "vertical";
        position: string;
        /** 背景。出さなければ null */
        background: { color: string; opacity: number } | null;
    };
    categoryAxis: {
        show: boolean;
        font: TextStyle;
        /** 合計・小計・区切りの差の棒の項目名 */
        totalFont: TextStyle;
        /** 項目名に使う大きさの上限（ビュー全体に対する割合 0〜1）。縦向きは高さ、横向きは幅 */
        maxShare: number;
        /** 帯の最小の幅（px）。これより狭くなるとスクロールする。0 ならスクロールしない */
        minCategoryWidth: number;
        title: AxisTitle | null;
    };
    valueAxis: { show: boolean; font: TextStyle; switchPosition: boolean; invert: boolean; title: AxisTitle | null };
    gridlines: { show: boolean; color: string; opacity: number; width: number; dash: string | null };
    legend: { show: boolean; position: LegendPosition; font: TextStyle; title: string };
    columns: { categorySpacing: number; outerPadding: number | null };
    /** 区切りの差の棒の濃さ（1 − 透過性）。ハイコントラストでは 1（ほかの棒と同じ） */
    subtotalOpacity: number;
    /** 軸を切った印の塗り（背景と同じ色） */
    background: string;
    /** ハイコントラストのときの前景・背景・選択の色。そうでなければ null */
    highContrast: { foreground: string; background: string; selected: string } | null;
}

export interface ViewModel {
    isEmpty: boolean;
    /** 描けないときに画面に出す文。isEmpty のときだけ使う */
    message: string;
    /** 描けるが、知らせたいこと（行が落ちたなど）。無ければ空 */
    notice: string;
    shape: Shape | null;
    bars: Bar[];
    /** イベントの名前（形 1）。書式の左端・右端の選択肢になる。形 1 でなければ空 */
    events: string[];
    leftEvent: string;
    rightEvent: string;
    axis: AxisInfo;
    /** 値の軸の単位（(百万円) など）。出さないときは空 */
    unitBadge: string;
    legend: LegendItem[];
    target: TargetLine | null;
    /** 定数線（固定の値）。出さなければ null */
    constantLine: TargetLine | null;
    hasHighlights: boolean;
    /** Power BI がまだ行を残している（metadata.segment） */
    truncated: boolean;
    /** 書式ペインに流し込むデータ次第の中身 */
    format: DataDrivenFormat;
    style: ViewStyle;
}

/** 比べる形の小計の名前に付ける語（「売上総利益の差」） */
export const SUBTOTAL_DIFF_SUFFIX = "の差";

/** 階層が無いときの、形 1 の増減の棒の名前 */
export const DIFF_TEXT = "差";
export const CUMULATIVE_TEXT = "累計";

/** 軸を切るときの余白。棒の根元側を広めに取り、合計の棒の長さが読めるようにする */
const CUT_PAD_BELOW = 0.3;
const CUT_PAD_ABOVE = 0.1;
const TICK_COUNT = 5;
/** ハイライトの該当分が全部か（丸めの誤差を許す） */
const EPSILON = 1e-9;

export const LANDING_MESSAGE =
    "「カテゴリ」と「値」を入れてください。値にメジャーを 2 つ（計画・実績など）入れるか、「比較」と「値」を入れると、その差を項目ごとに並べます。";

export const LOADING_NOTICE = "続きの行を読み込んでいます…";

export const TRUNCATED_NOTICE = "Power BI の行の上限で、すべての行を読み込めていません。合計・小計・両端は読み込めた行で計算しています。";

const RATE_NAMES: Record<string, string> = {
    [RATE_TYPES.growth]: "伸び率",
    [RATE_TYPES.contribution]: "寄与率",
    [RATE_TYPES.share]: "構成比",
};

function colorOf(slice: { value: { value: string } }): string {
    return slice.value?.value ?? "";
}

export function dashOf(style: string, width: number): string | null {
    if (style === LINE_STYLES.dashed) return `${Math.max(2, width * 4)} ${Math.max(2, width * 2)}`;
    if (style === LINE_STYLES.dotted) return `${Math.max(1, width)} ${Math.max(2, width * 2)}`;
    return null;
}

const clamp = (value: number, lo: number, hi: number) => Math.max(lo, Math.min(hi, value));
/** 数でなければ fallback（0 は 0 のまま） */
const finiteOr = (value: unknown, fallback: number): number => {
    const n = typeof value === "number" ? value : Number(value);
    return value === null || value === undefined || value === "" || !Number.isFinite(n) ? fallback : n;
};
/** 区切りの差の透過性の既定（%） */
const DEFAULT_SUBTOTAL_TRANSPARENCY = 50;

/** FontControl の値。size はポイント */
function fontOf(control: formattingSettings.FontControl, color: string): TextStyle {
    return {
        family: control.fontFamily.value || DEFAULT_FONT_FAMILY,
        size: Number(control.fontSize.value) || 9,
        bold: !!control.bold?.value,
        italic: !!control.italic?.value,
        underline: !!control.underline?.value,
        color,
    };
}

/**
 * グリッド線の線種。棒グラフと同じく、点線は細かい点、破線は 4px 刻み。
 * 「幅で拡大縮小」がオンなら、点線・破線の模様を線の幅に比例させる
 */
export function gridDashOf(style: string, width: number, scaleWithWidth: boolean): string | null {
    const w = Math.max(1, width);
    if (style === LINE_STYLES.dotted) return scaleWithWidth ? `${w} ${2 * w}` : "1 3";
    if (style === LINE_STYLES.dashed) return scaleWithWidth ? `${3 * w} ${3 * w}` : "4 4";
    return null;
}

function styleOf(settings: VisualFormattingSettingsModel, host: IVisualHost): ViewStyle {
    const palette = host.colorPalette as ISandboxExtendedColorPalette;
    const hc = palette.isHighContrast
        ? {
              foreground: palette.foreground?.value ?? "#000000",
              background: palette.background?.value ?? "#FFFFFF",
              selected: palette.foregroundSelected?.value ?? palette.foreground?.value ?? "#000000",
          }
        : null;
    // ハイコントラストでは、文字・線はすべて前景色にする
    const ink = (color: string) => hc?.foreground ?? color;
    const connectorWidth = clamp(Number(settings.connectors.width.value) || 1, 0.5, 10);
    const labels = settings.dataLabels;
    const gridWidth = clamp(Number(settings.gridlines.horizontalWidth.value) || 1, 0.5, 10);
    const outerPadding = settings.columns.outerPadding.value;
    const legendPosition = dropdownValue(settings.legend.position, LEGEND_POSITIONS.topLeft);
    return {
        orientation:
            dropdownValue(settings.layout.orientation, ORIENTATIONS.vertical) === ORIENTATIONS.horizontal
                ? ORIENTATIONS.horizontal
                : ORIENTATIONS.vertical,
        breakStyle:
            dropdownValue(settings.valueAxis.breakStyle, BREAK_STYLES.slash) === BREAK_STYLES.wave ? BREAK_STYLES.wave : BREAK_STYLES.slash,
        connectors: {
            show: settings.connectors.show.value ?? true,
            color: ink(colorOf(settings.connectors.color) || palette.foregroundNeutralTertiary?.value || "#A19F9D"),
            width: connectorWidth,
            dash: dashOf(dropdownValue(settings.connectors.lineStyle, LINE_STYLES.solid), connectorWidth),
        },
        dataLabels: {
            show: labels.show.value ?? true,
            // 空は自動のまま渡す（描画で棒の中か外かで決める）。ハイコントラストは前景色
            font: fontOf(labels.font, hc ? hc.foreground : colorOf(labels.color)),
            totalFont: fontOf(labels.totalFont, hc ? hc.foreground : colorOf(labels.totalColor) || colorOf(labels.color)),
            orientation:
                dropdownValue(labels.orientation, LABEL_ORIENTATIONS.horizontal) === LABEL_ORIENTATIONS.vertical ? "vertical" : "horizontal",
            position: dropdownValue(labels.position, LABEL_POSITIONS.auto),
            background:
                (labels.backgroundShow.value ?? false) && !hc
                    ? {
                          color: colorOf(labels.backgroundColor) || DEFAULT_LABEL_BACKGROUND,
                          opacity: (100 - clamp(finiteOr(labels.backgroundTransparency.value, DEFAULT_LABEL_BACKGROUND_TRANSPARENCY), 0, 100)) / 100,
                      }
                    : null,
        },
        categoryAxis: {
            show: settings.categoryAxis.show.value ?? true,
            font: fontOf(settings.categoryAxis.font, ink(colorOf(settings.categoryAxis.labelColor) || "#605E5C")),
            totalFont: fontOf(
                settings.categoryAxis.totalFont,
                ink(colorOf(settings.categoryAxis.totalColor) || colorOf(settings.categoryAxis.labelColor) || "#605E5C")
            ),
            maxShare: clamp(Number(settings.categoryAxis.maxHeight.value) || 25, 5, 100) / 100,
            minCategoryWidth: Math.max(0, Number(settings.categoryAxis.minCategoryWidth.value) || 0),
            title: null,
        },
        valueAxis: {
            show: settings.valueAxis.show.value ?? true,
            font: fontOf(settings.valueAxis.font, ink(colorOf(settings.valueAxis.labelColor) || "#605E5C")),
            switchPosition: settings.valueAxis.switchPosition.value ?? false,
            invert: settings.valueAxis.invertRange.value ?? false,
            title: null,
        },
        gridlines: {
            show: settings.gridlines.horizontalShow.value ?? true,
            color: ink(colorOf(settings.gridlines.horizontalColor) || "#E1DFDD"),
            opacity: hc ? 0.5 : 1 - clamp(Number(settings.gridlines.horizontalTransparency.value) || 0, 0, 100) / 100,
            width: gridWidth,
            dash: gridDashOf(
                dropdownValue(settings.gridlines.horizontalStyle, LINE_STYLES.dotted),
                gridWidth,
                settings.gridlines.horizontalScaleWithWidth.value ?? false
            ),
        },
        legend: {
            show: settings.legend.show.value ?? true,
            position: ((Object.values(LEGEND_POSITIONS) as string[]).includes(legendPosition)
                ? legendPosition
                : LEGEND_POSITIONS.topLeft) as LegendPosition,
            font: fontOf(settings.legend.font, ink(colorOf(settings.legend.labelColor) || "#605E5C")),
            title: "",
        },
        columns: {
            categorySpacing: clamp(Number(settings.columns.categorySpacing.value ?? 20), 0, 50),
            outerPadding: typeof outerPadding === "number" && Number.isFinite(outerPadding) ? clamp(outerPadding, 0, 100) : null,
        },
        subtotalOpacity: hc ? 1 : 1 - clamp(finiteOr(settings.columns.subtotalTransparency.value, DEFAULT_SUBTOTAL_TRANSPARENCY), 0, 100) / 100,
        background: hc?.background ?? palette.background?.value ?? "#FFFFFF",
        highContrast: hc,
    };
}

const NO_FORMAT: DataDrivenFormat = {
    events: [],
    leftEvent: "",
    rightEvent: "",
    hasOrigin: false,
    hasTarget: false,
    seriesTargets: [],
    horizontal: false,
    subtotalDiff: false,
};

function emptyOf(style: ViewStyle, message: string): ViewModel {
    return {
        isEmpty: true,
        message,
        notice: "",
        shape: null,
        bars: [],
        events: [],
        leftEvent: "",
        rightEvent: "",
        axis: { min: 0, max: 1, ticks: [], cut: "none" },
        unitBadge: "",
        legend: [],
        target: null,
        constantLine: null,
        hasHighlights: false,
        truncated: false,
        format: { ...NO_FORMAT, horizontal: style.orientation === ORIENTATIONS.horizontal },
        style,
    };
}

interface Colors {
    increase: string;
    decrease: string;
    total: string;
    others: string;
}

/**
 * 色を toward に amount だけ寄せる（#RRGGBB だけ。読めなければそのまま）。
 * 透かして描く区切りの差の見た目の色（背景と混ざった色）を出し、中に置くラベルの文字色を決めるのに使う
 */
export function blend(color: string, toward: string, amount: number): string {
    const parse = (c: string) => {
        const match = /^#([0-9a-f]{6})$/i.exec(c.trim());
        return match ? parseInt(match[1], 16) : null;
    };
    const from = parse(color);
    const to = parse(toward);
    if (from === null || to === null) return color;
    const channel = (shift: number) => {
        const a = (from >> shift) & 0xff;
        const b = (to >> shift) & 0xff;
        return Math.round(a + (b - a) * amount);
    };
    return `#${[16, 8, 0].map((shift) => channel(shift).toString(16).padStart(2, "0")).join("").toUpperCase()}`;
}

/** 区切りの差の色：差が 0 以上なら増加の色、マイナスなら減少の色（0 は増加に寄せる） */
export const subtotalColorOf = (delta: number, colors: { increase: string; decrease: string }): string =>
    delta >= 0 ? colors.increase : colors.decrease;

/** 色。書式で指定していなければテーマの色にし、書式ペインにもその色を出す */
function resolveColors(settings: VisualFormattingSettingsModel, host: IVisualHost): Colors {
    const palette = host.colorPalette as ISandboxExtendedColorPalette;
    const card = settings.columns;
    // Desktop の colorPalette は呼んだ順にテーマのデータの色を割り当てる。合計を先に呼び、1 色目にする
    const themeTotal = palette.getColor("total").value;
    const themeIncrease = palette.positive?.value ?? palette.getColor("increase").value;
    const themeDecrease = palette.negative?.value ?? palette.getColor("decrease").value;
    const themeOthers = palette.foregroundNeutralSecondary?.value ?? palette.getColor("others").value;
    const pick = (slice: typeof card.totalFill, theme: string): string => {
        const chosen = colorOf(slice);
        if (chosen) return chosen;
        slice.value = { value: theme };
        return theme;
    };
    const total = pick(card.totalFill, themeTotal);
    return {
        total,
        increase: pick(card.increaseFill, themeIncrease),
        decrease: pick(card.decreaseFill, themeDecrease),
        others: pick(card.othersFill, themeOthers),
    };
}

function withSign(text: string, value: number, plus: boolean): string {
    return plus && value > 0 && !text.startsWith("+") ? `+${text}` : text;
}

function percentText(rate: number, signed: boolean): string {
    const text = `${(rate * 100).toFixed(1)}%`;
    return signed && rate > 0 ? `+${text}` : text;
}

/** 並べ替え。同じ値のときは元の順（安定） */
function sortEntries<T extends { delta: number }>(entries: T[], order: string): T[] {
    if (order === ORDERS.data) return entries;
    const key = order === ORDERS.absDelta ? (e: T) => Math.abs(e.delta) : (e: T) => e.delta;
    return entries
        .map((entry, index) => ({ entry, index }))
        .sort((a, b) => key(b.entry) - key(a.entry) || a.index - b.index)
        .map(({ entry }) => entry);
}

/**
 * 「その他」にまとめる。増減の絶対値の大きい keep 個を残す（残したものの並びはそのまま）。
 * まとめるのが 1 つだけなら、その項目をそのまま出す（棒の数は同じで、名前が分かるほうがよい）
 */
function splitOthers<T extends { delta: number }>(entries: T[], keep: number): { kept: T[]; merged: T[] } {
    if (entries.length - keep < 2) return { kept: entries, merged: [] };
    const ranked = entries
        .map((entry, index) => ({ entry, index }))
        .sort((a, b) => Math.abs(b.entry.delta) - Math.abs(a.entry.delta) || a.index - b.index);
    const keepSet = new Set(ranked.slice(0, keep).map(({ entry }) => entry));
    return {
        kept: entries.filter((entry) => keepSet.has(entry)),
        merged: entries.filter((entry) => !keepSet.has(entry)),
    };
}

/** 棒の区画が描く範囲（系列で積むと、正と負の区画で上下に広がる） */
export function extentOf(bar: Bar): [number, number] {
    const ends = bar.segments.flatMap((s) => [s.from, s.to]);
    return ends.length ? [Math.min(...ends), Math.max(...ends)] : [Math.min(bar.from, bar.to), Math.max(bar.from, bar.to)];
}

export interface AxisOptions {
    /** 0 から描く（最小値を入れたときは最小値が優先） */
    startAtZero: boolean;
    /** 書式の最小値・最大値。空なら null（自動） */
    start: number | null;
    end: number | null;
    /** 範囲を丸める（自動の側だけを切りのよい値にする） */
    round: boolean;
}

/** 書式の最小値・最大値の文字を数に直す。空・数でなければ null（自動） */
export function boundOf(text: string | undefined): number | null {
    const trimmed = (text ?? "").trim().replace(/,/g, "");
    if (!trimmed) return null;
    const value = Number(trimmed);
    return Number.isFinite(value) ? value : null;
}

/**
 * 軸の範囲。0 から描かないときは、増減が見やすいように合計の棒を切る。
 * 累計が 0 をまたぐときや、形 2 のように 0 から積むときは 0 が範囲に入るので切らない。
 * 書式の最小値・最大値は、切る・0 から描くより優先する（入れた値のまま。丸めない）。
 * 最小値が 0 より大きい（最大値が 0 より小さい）と、合計の棒の根元が軸の外になるので「切った」ことになる
 */
export function axisOf(bars: Bar[], options: AxisOptions, formatTick: (v: number) => string, extra: number[] = []): AxisInfo {
    const { startAtZero, start, end, round } = options;
    const points: number[] = [...extra];
    for (const bar of bars) {
        if (bar.kind === BAR_KINDS.total) points.push(bar.to);
        else points.push(...extentOf(bar));
    }
    if ((startAtZero && start === null) || !points.length) points.push(0);
    const low = Math.min(...points);
    const high = Math.max(...points);
    let lo = low;
    let hi = high;
    if (!startAtZero) {
        const span = hi - lo || Math.abs(hi) * 0.1 || 1;
        lo -= span * CUT_PAD_BELOW;
        hi += span * CUT_PAD_ABOVE;
        // 余白のせいで 0 をまたがない（0 をまたぐのはデータがまたぐときだけ）
        if (low >= 0 && lo < 0) lo = 0;
        if (high <= 0 && hi > 0) hi = 0;
    }
    if (start !== null) lo = start;
    if (end !== null) hi = end;
    if (lo >= hi) hi = lo + (Math.abs(lo) * 0.1 || 1);
    let min = lo;
    let max = hi;
    if (round) {
        const [niceMin, niceMax] = scaleLinear().domain([lo, hi]).nice(TICK_COUNT).domain();
        if (start === null) min = niceMin;
        if (end === null) max = niceMax;
    }
    const cut = min > 0 ? "bottom" : max < 0 ? "top" : "none";
    return {
        min,
        max,
        ticks: scaleLinear()
            .domain([min, max])
            .ticks(TICK_COUNT)
            .map((value) => ({ value, label: formatTick(value) })),
        cut,
    };
}

/** 系列 1 つぶんの増減。値の枠 before → after（形 2・3 は before が null） */
interface Part {
    series: number;
    delta: number;
    /** ハイライトの該当分の増減。該当しなければ null */
    highlight: number | null;
}

const sumOf = (values: Array<number | null>): number | null =>
    values.reduce<number | null>((sum, v) => (v === null ? sum : (sum ?? 0) + v), null);

export function transform(
    dataView: DataView | undefined,
    host: IVisualHost,
    settings: VisualFormattingSettingsModel
): ViewModel {
    const style = styleOf(settings, host);
    const parsed: Parsed = parse(dataView, host, LANDING_MESSAGE);
    if (parsed.error) return emptyOf(style, parsed.error);

    const { sources, series, items, sections, hasSeries, hasHighlights } = parsed;
    const origin = parsed.origin;
    const shape: Shape = sources.length >= 2 ? SHAPES.compare : origin ? SHAPES.origin : SHAPES.merge;
    if (shape === SHAPES.compare && origin) {
        return emptyOf(
            style,
            "「起点」は、値が 1 つで「比較」を使わないときに入れます。比べる形（「比較」、または 2 つ以上の値）では「起点」を外してください。"
        );
    }

    // --- 左端・右端のイベント（形 1） -------------------------------------------------
    const eventNames = shape === SHAPES.compare ? sources.map((s) => s.name) : [];
    const layoutObject = dataView?.metadata?.objects?.layout as DataViewObject | undefined;
    // 保存値は populate の時点では選択肢に無く、スライスには入らない。生の objects から読む
    const savedIndex = (property: "leftEvent" | "rightEvent"): number => {
        const raw = layoutObject?.[property];
        return typeof raw === "string" ? eventNames.indexOf(raw) : -1;
    };
    const leftIndex = shape === SHAPES.compare && savedIndex("leftEvent") >= 0 ? savedIndex("leftEvent") : 0;
    const rightIndex =
        shape === SHAPES.compare ? (savedIndex("rightEvent") >= 0 ? savedIndex("rightEvent") : sources.length - 1) : 0;

    // --- 色 --------------------------------------------------------------------------------
    const colors = resolveColors(settings, host);
    const palette = host.colorPalette;
    const seriesColors = series.map((s) => {
        if (!hasSeries) return "";
        // 保存先は「列」カードの fill（凡例の値ごとの selector 付き。棒グラフと同じ）
        const saved = (s.group.objects?.columns?.fill as powerbi.Fill | undefined)?.solid?.color;
        return saved ? String(saved) : palette.getColor(s.name).value;
    });

    // --- 設定 --------------------------------------------------------------------------------
    const order = dropdownValue(settings.layout.order, ORDERS.data);
    const connectMode = dropdownValue(settings.layout.connectMode, CONNECT_MODES.ends);
    const originMode = dropdownValue(settings.layout.originMode, ORIGIN_MODES.first);
    const othersOn = settings.others.show.value ?? false;
    const othersKeep = Math.max(1, Math.min(1000, Math.round(Number(settings.others.count.value) || 1)));
    const othersLabel = (settings.others.label.value ?? "").trim() || DEFAULT_OTHERS_LABEL;
    const totalShow = settings.layout.totalShow.value ?? true;
    const totalLabel = (settings.layout.totalLabel.value ?? "").trim() || DEFAULT_TOTAL_LABEL;
    const rateType = dropdownValue(settings.dataLabels.rateType, RATE_TYPES.none);

    // --- 棒を組み立てる -----------------------------------------------------------------
    const bars: Bar[] = [];
    /** 率の分母（棒の添字ごと）。最後の合計は組み立てたあとで決まるので、あとで割る */
    const rateBase: Array<{ growth: number | null; contribution: number | null } | null> = [];

    const totalBar = (
        role: TotalRole,
        name: string,
        value: number,
        highlight: number | null,
        key: string,
        ids: ISelectionId[],
        tooltip: TooltipItem[]
    ) => {
        const fully = highlight !== null && Math.abs(highlight - value) <= Math.abs(value) * EPSILON + EPSILON;
        bars.push({
            key,
            kind: BAR_KINDS.total,
            totalRole: role,
            label: name,
            parentLabel: "",
            fullLabel: name,
            from: 0,
            to: value,
            value,
            labelText: "",
            rateText: "",
            color: colors.total,
            segments: [
                {
                    from: 0,
                    to: value,
                    color: colors.total,
                    series: null,
                    highlightTo: hasHighlights && highlight !== null && !fully ? highlight : null,
                    selectionIds: ids,
                },
            ],
            dimmed: hasHighlights && !fully,
            selectionIds: ids,
            tooltip,
        });
        rateBase.push(null);
    };

    /**
     * 増減の棒。系列があれば系列の順に積む。正の区画は累計から上へ、負の区画は下へ積む
     * （順に積むと、向きが変わったところで区画が重なって長さが読めなくなるため）。累計は正味の増減で進む
     */
    const deltaBar = (
        key: string,
        kind: BarKind,
        labels: { label: string; parentLabel: string; fullLabel: string },
        level: number,
        parts: Part[],
        idsOfSeries: (s: number) => ISelectionId[],
        ids: ISelectionId[],
        tooltip: TooltipItem[],
        base: { growth: number | null; contribution: number | null }
    ) => {
        const delta = parts.reduce((sum, p) => sum + p.delta, 0);
        const color =
            kind === BAR_KINDS.others ? colors.others : delta < 0 ? colors.decrease : colors.increase;
        let up = level;
        let down = level;
        const segments: Segment[] = [];
        let fully = true;
        for (const part of parts) {
            if (part.highlight === null || Math.abs(part.highlight - part.delta) > Math.abs(part.delta) * EPSILON + EPSILON) {
                fully = false;
            }
            if (part.delta === 0) continue;
            const from = part.delta > 0 ? up : down;
            const to = from + part.delta;
            if (part.delta > 0) up = to;
            else down = to;
            segments.push({
                from,
                to,
                color: hasSeries ? seriesColors[part.series] : color,
                series: hasSeries ? part.series : null,
                highlightTo: hasHighlights && part.highlight !== null ? from + part.highlight : null,
                selectionIds: hasSeries ? idsOfSeries(part.series) : ids,
            });
        }
        if (!segments.length) {
            segments.push({ from: level, to: level, color, series: null, highlightTo: null, selectionIds: ids });
        }
        // 全部が該当する棒は、薄くも重ねもしない
        const dimmed = hasHighlights && !fully;
        if (!dimmed) segments.forEach((segment) => (segment.highlightTo = null));
        bars.push({
            key,
            kind,
            totalRole: null,
            ...labels,
            from: level,
            to: level + delta,
            value: delta,
            labelText: "",
            rateText: "",
            color,
            segments,
            dimmed,
            selectionIds: ids,
            tooltip,
        });
        rateBase.push(base);
        return level + delta;
    };

    /**
     * 上の階層のうち、どの項目でも同じレベルは棒ごとに出さない。ドリルダウンすると親のレベルも届くが
     * （事業A → 製品A1 に入ると、どの棒も「事業A / 製品A1」）、全部同じ親を棒ごとに繰り返しても読めない。
     * 「すべて展開」で親が混ざるレベルから下は出す。読み上げ・ツールヒントの名前（fullLabel）は上から全部
     */
    const commonDepth = (() => {
        const withLevels = items.filter((item) => item.levels.length > 1);
        if (!withLevels.length) return 0;
        const maxDepth = Math.min(...withLevels.map((item) => item.levels.length - 1));
        let depth = 0;
        while (depth < maxDepth && withLevels.every((item) => item.levels[depth] === withLevels[0].levels[depth])) depth++;
        return depth;
    })();
    const itemLabels = (item: ParsedItem) => {
        if (!item.levels.length) {
            const name = shape === SHAPES.compare ? DIFF_TEXT : sources[0].name;
            return { label: name, parentLabel: "", fullLabel: name };
        }
        return {
            label: item.levels[item.levels.length - 1],
            parentLabel: item.levels.slice(commonDepth, -1).join(LEVEL_SEPARATOR),
            fullLabel: item.levels.join(LEVEL_SEPARATOR),
        };
    };

    /** 値の枠 k の合計（全項目・全系列） */
    const totalOf = (k: number) => items.reduce((sum, item) => sum + (sumOf(item.values[k]) ?? 0), 0);
    const highlightTotalOf = (k: number) =>
        hasHighlights ? items.reduce<number | null>((sum, item) => {
            const h = sumOf(item.highlights[k]);
            return h === null ? sum : (sum ?? 0) + h;
        }, null) : null;

    /**
     * 増減の棒を区切りごとに並べ、区切りの最後に小計の棒を立てる。
     * from / to は値の枠の添字。形 2・3 では from が null（項目の値そのものが増減）。
     * level はそこまでの累計、hLevel はハイライトの累計
     */
    const runSegment = (
        from: number | null,
        to: number,
        start: { level: number; hLevel: number | null; contributionBase: number | null },
        segmentKey: string
    ): { level: number; hLevel: number | null } => {
        let level = start.level;
        let hLevel = start.hLevel;
        const format: Formatter = sources[to].format;
        const partsOf = (item: ParsedItem): Part[] | null => {
            const after = item.values[to];
            const before = from === null ? null : item.values[from];
            const any = after.some((v) => v !== null) || (before?.some((v) => v !== null) ?? false);
            if (!any) return null;
            return series.map((_, s) => {
                const d = (after[s] ?? 0) - (before?.[s] ?? 0);
                const ha = item.highlights[to][s];
                const hb = from === null ? null : item.highlights[from][s];
                const highlight = ha === null && hb === null ? null : (ha ?? 0) - (hb ?? 0);
                return { series: s, delta: d, highlight };
            });
        };
        /** ツールヒント。形 1 は左端の値・右端の値・差、形 2・3 は値と累計。系列があれば系列ごとの増減 */
        const describe = (
            name: TooltipItem | null,
            before: number | null,
            after: number | null,
            parts: Part[],
            levelAfter: number
        ): TooltipItem[] => {
            const delta = parts.reduce((sum, p) => sum + p.delta, 0);
            return [
                ...(name ? [name] : []),
                ...(from === null
                    ? [
                          { displayName: sources[to].name, value: format(delta) },
                          { displayName: CUMULATIVE_TEXT, value: format(levelAfter) },
                      ]
                    : [
                          { displayName: sources[from].name, value: sources[from].format(before) },
                          { displayName: sources[to].name, value: format(after) },
                          { displayName: DIFF_TEXT, value: withSign(format(delta), delta, true) },
                      ]),
                ...(hasSeries
                    ? parts
                          .filter((p) => p.delta !== 0)
                          .map((p) => ({ displayName: series[p.series].name, value: withSign(format(p.delta), p.delta, true) }))
                    : []),
            ];
        };
        const advanceHighlight = (parts: Part[]) => {
            if (hLevel === null) return;
            hLevel += parts.reduce((sum, p) => sum + (p.highlight ?? 0), 0);
        };

        for (const section of sections) {
            const sectionStart = level;
            const hSectionStart = hLevel;
            const entries = section.items
                .map((item) => ({ item, parts: partsOf(item) }))
                .filter((entry): entry is { item: ParsedItem; parts: Part[] } => entry.parts !== null)
                .map((entry) => ({ ...entry, delta: entry.parts.reduce((sum, p) => sum + p.delta, 0) }));
            if (!entries.length) continue;
            const sorted = sortEntries(entries, order);
            const { kept, merged } = othersOn ? splitOthers(sorted, othersKeep) : { kept: sorted, merged: [] };

            for (const { item, parts } of kept) {
                const labels = itemLabels(item);
                const before = from === null ? null : sumOf(item.values[from]);
                const after = sumOf(item.values[to]);
                const delta = parts.reduce((sum, p) => sum + p.delta, 0);
                level = deltaBar(
                    `${segmentKey}|${item.key}`,
                    delta < 0 ? BAR_KINDS.decrease : BAR_KINDS.increase,
                    labels,
                    level,
                    parts,
                    (s) => parsed.idsOf(item.rows, s),
                    parsed.idsOf(item.rows),
                    [
                        ...describe(
                            parsed.hierarchyName ? { displayName: parsed.hierarchyName, value: labels.fullLabel } : null,
                            before,
                            after,
                            parts,
                            level + delta
                        ),
                        // 足したフィールドは、右端（形 2・3 は値）の行で評価されたもの
                        ...parsed.tooltipItemsOf(item.sourceRows[to]),
                    ],
                    { growth: before, contribution: start.contributionBase }
                );
                advanceHighlight(parts);
            }

            if (merged.length) {
                const parts: Part[] = series.map((_, s) => {
                    const ofSeries = merged.map((e) => e.parts[s]);
                    const highlights = ofSeries.map((p) => p.highlight);
                    return {
                        series: s,
                        delta: ofSeries.reduce((sum, p) => sum + p.delta, 0),
                        highlight: highlights.every((h) => h === null) ? null : highlights.reduce<number>((sum, h) => sum + (h ?? 0), 0),
                    };
                });
                const rows = merged.flatMap((e) => e.item.rows);
                const before = from === null ? null : sumOf(merged.map((e) => sumOf(e.item.values[from])));
                const after = sumOf(merged.map((e) => sumOf(e.item.values[to])));
                const delta = parts.reduce((sum, p) => sum + p.delta, 0);
                level = deltaBar(
                    `${segmentKey}|${section.key}|others`,
                    BAR_KINDS.others,
                    { label: othersLabel, parentLabel: "", fullLabel: othersLabel },
                    level,
                    parts,
                    (s) => parsed.idsOf(rows, s),
                    parsed.idsOf(rows),
                    describe({ displayName: othersLabel, value: `${merged.length} 項目` }, before, after, parts, level + delta),
                    { growth: before, contribution: start.contributionBase }
                );
                advanceHighlight(parts);
            }

            if (section.name !== null && from === null) {
                // 合流する形・起点から積む形の小計は、そこまでの累計（軸から立つ棒）
                totalBar("subtotal", section.name, level, hLevel, `${segmentKey}|${section.key}|subtotal`, [], [
                    { displayName: section.name, value: format(level) },
                ]);
            } else if (section.name !== null && from !== null) {
                // 比べる形の小計は、その区切りの差の合計（浮いた棒）。累計（計画 ＋ そこまでの差）は
                // 計画の値でも実績の値でもなく読み違えやすいため（docs/waterfall.md の「小計」）
                const name = `${section.name}${SUBTOTAL_DIFF_SUFFIX}`;
                const delta = level - sectionStart;
                const before = sumOf(entries.map((e) => sumOf(e.item.values[from])));
                const after = sumOf(entries.map((e) => sumOf(e.item.values[to])));
                const hDelta = hLevel === null || hSectionStart === null ? null : hLevel - hSectionStart;
                const fully = hDelta !== null && Math.abs(hDelta - delta) <= Math.abs(delta) * EPSILON + EPSILON;
                const ids = parsed.idsOf(entries.flatMap((e) => e.item.rows));
                bars.push({
                    key: `${segmentKey}|${section.key}|subtotal`,
                    kind: BAR_KINDS.subtotal,
                    totalRole: "subtotal",
                    label: name,
                    parentLabel: "",
                    fullLabel: name,
                    from: sectionStart,
                    to: level,
                    value: delta,
                    labelText: "",
                    rateText: "",
                    color: subtotalColorOf(delta, colors),
                    segments: [
                        {
                            from: sectionStart,
                            to: level,
                            color: subtotalColorOf(delta, colors),
                            series: null,
                            highlightTo: hasHighlights && hDelta !== null && !fully ? sectionStart + hDelta : null,
                            selectionIds: ids,
                        },
                    ],
                    dimmed: hasHighlights && !fully,
                    selectionIds: ids,
                    tooltip: [
                        { displayName: name, value: withSign(format(delta), delta, true) },
                        { displayName: sources[from].name, value: sources[from].format(before) },
                        { displayName: sources[to].name, value: format(after) },
                    ],
                });
                rateBase.push({ growth: before, contribution: start.contributionBase });
            }
        }
        return { level, hLevel };
    };

    const eventTotal = (index: number, key: string) => {
        const source = sources[index];
        const total = totalOf(index);
        totalBar("event", source.name, total, highlightTotalOf(index), key, source.selectionIds, [
            { displayName: source.name, value: source.format(total) },
        ]);
    };

    let finalLevel = 0;
    if (shape === SHAPES.compare) {
        // 両端だけなら [左, 右]、すべてなら左から右までのイベントを順に（右が左より前なら逆向きにたどる）
        const step = rightIndex >= leftIndex ? 1 : -1;
        const sequence =
            connectMode === CONNECT_MODES.all
                ? Array.from({ length: Math.abs(rightIndex - leftIndex) + 1 }, (_, k) => leftIndex + k * step)
                : [leftIndex, rightIndex];
        eventTotal(sequence[0], "event0");
        let state = { level: totalOf(sequence[0]), hLevel: highlightTotalOf(sequence[0]) };
        for (let k = 1; k < sequence.length; k++) {
            state = runSegment(
                sequence[k - 1],
                sequence[k],
                { ...state, contributionBase: totalOf(sequence[k - 1]) },
                `seg${k}`
            );
            eventTotal(sequence[k], `event${k}`);
        }
        finalLevel = totalOf(sequence[sequence.length - 1]);
    } else {
        let state: { level: number; hLevel: number | null } = { level: 0, hLevel: hasHighlights ? 0 : null };
        let contributionBase: number | null = null;
        if (shape === SHAPES.origin) {
            // 起点は項目ごとに届く。最初の項目の値か、項目の合計（書式で選ぶ）
            const pick = (values: Array<number | null>) =>
                originMode === ORIGIN_MODES.sum ? (sumOf(values) ?? 0) : (values.find((v) => v !== null) ?? 0);
            const value = pick(origin!.perItem);
            const highlight = hasHighlights ? pick(origin!.perItemHighlight) : null;
            totalBar("origin", origin!.name, value, highlight, "origin", [], [{ displayName: origin!.name, value: origin!.format(value) }]);
            state = { level: value, hLevel: highlight };
            contributionBase = value;
        }
        state = runSegment(null, 0, { ...state, contributionBase }, "seg1");
        finalLevel = state.level;
        if (totalShow) {
            totalBar("final", totalLabel, state.level, state.hLevel, "final", [], [
                { displayName: totalLabel, value: sources[0].format(state.level) },
            ]);
        }
    }

    // --- 率 -------------------------------------------------------------------------------
    if (rateType !== RATE_TYPES.none) {
        bars.forEach((bar, i) => {
            const base = rateBase[i];
            if (!base) return;
            let rate: number | null = null;
            // 分母は絶対値で割る。費用（マイナス）が減ったときも、差の向き（利益への効き目）のまま率を出す
            if (rateType === RATE_TYPES.growth && shape === SHAPES.compare && base.growth) rate = bar.value / Math.abs(base.growth);
            if (rateType === RATE_TYPES.contribution && base.contribution) rate = bar.value / Math.abs(base.contribution);
            if (rateType === RATE_TYPES.share && shape === SHAPES.merge && finalLevel) rate = bar.value / Math.abs(finalLevel);
            if (rate === null) return;
            bar.rateText = percentText(rate, rateType !== RATE_TYPES.share);
            bar.tooltip.push({ displayName: RATE_NAMES[rateType], value: bar.rateText });
        });
    }

    // --- 単位・ラベル・軸 ---------------------------------------------------------------
    const axisSettings = settings.valueAxis;
    const notation = dropdownValue(axisSettings.unitNotation, "japanese");
    const axisPrecision = dropdownValue(axisSettings.precision, "auto");
    const labelPrecision = dropdownValue(settings.dataLabels.precision, "auto");
    // 自動の表示単位は増減の大きさで決める。両端（合計）で決めると、差が 0.19 億のように潰れる
    const maxAbs = (kind: (bar: Bar) => boolean) =>
        bars.filter(kind).reduce((m, bar) => Math.max(m, Math.abs(bar.value)), 0);
    const deltaMax = maxAbs((bar) => bar.kind !== BAR_KINDS.total);
    const unitBase = deltaMax || maxAbs((bar) => bar.kind === BAR_KINDS.total);
    const unit = resolveUnit(dropdownValue(axisSettings.unitType, "auto"), unitBase, notation, axisPrecision);
    // データ ラベルの表示単位。「Y 軸と同じ」でなければラベル自身の単位で出し、単位の語を付ける（棒グラフの合計ラベルと同じ）
    const labelUnitKey = dropdownValue(settings.dataLabels.unitType, "auto");
    const labelUnit = labelUnitKey === "auto" ? null : resolveUnit(labelUnitKey, unitBase, notation, labelPrecision);
    const labelOf = (v: number) =>
        labelUnit
            ? `${formatValue(v, labelUnit.divisor, labelPrecision)}${labelUnit.unitWord}`
            : formatValue(v, unit.divisor, labelPrecision);
    const plusSign = settings.dataLabels.plusSign.value ?? true;
    for (const bar of bars) {
        const text = labelOf(bar.value);
        bar.labelText = bar.kind === BAR_KINDS.total ? text : withSign(text, bar.value, plusSign);
    }

    const ink = (color: string) => style.highContrast?.foreground ?? color;

    // --- 目標 -------------------------------------------------------------------------------
    // 目標も項目ごとに届く。どの項目も同じ値ならその値、違えば項目の合計（項目に割り振った目標）
    let target: TargetLine | null = null;
    const targetValues = (parsed.target?.perItem ?? []).filter((v): v is number => v !== null);
    if (parsed.target && targetValues.length && (settings.target.show.value ?? true)) {
        const value = targetValues.every((v) => v === targetValues[0])
            ? targetValues[0]
            : targetValues.reduce((sum, v) => sum + v, 0);
        const width = clamp(Number(settings.target.width.value) || 1.5, 0.5, 10);
        target = {
            value,
            name: parsed.target.name,
            // 名前は凡例に出す。グラフの中の名前と値は既定で出さない（ユーザー「目標のラベルの位置が違和感」）
            label: (settings.target.labelShow.value ?? false) ? `${parsed.target.name} ${labelOf(value)}` : "",
            tooltip: [{ displayName: parsed.target.name, value: parsed.target.format(value) }],
            color: ink(colorOf(settings.target.color) || "#252423"),
            width,
            dash: dashOf(dropdownValue(settings.target.lineStyle, LINE_STYLES.dashed), width),
        };
    }

    // --- 定数線（固定の値。軸は広げない） ---------------------------------------------------
    let constantLine: TargetLine | null = null;
    const constant = settings.constantLine;
    if (constant.show.value ?? false) {
        const value = Number(constant.value.value) || 0;
        const width = clamp(Number(constant.width.value) || 1, 0.5, 10);
        const name = (constant.labelText.value ?? "").trim();
        constantLine = {
            value,
            name,
            // 定数線は標準と同じくグラフの中に名前と値を出す（凡例には入れない）
            label: (constant.labelShow.value ?? true) ? `${name ? `${name} ` : ""}${labelOf(value)}` : "",
            tooltip: [],
            color: ink(colorOf(constant.color) || "#605E5C"),
            width,
            dash: dashOf(dropdownValue(constant.lineStyle, LINE_STYLES.dashed), width),
        };
    }

    const axis = axisOf(
        bars,
        {
            startAtZero: axisSettings.startAtZero.value ?? false,
            start: boundOf(axisSettings.start.value),
            end: boundOf(axisSettings.end.value),
            round: axisSettings.roundRange.value ?? true,
        },
        (v) => formatValue(v, unit.divisor, axisPrecision),
        target ? [target.value] : []
    );

    // --- 軸のタイトル -----------------------------------------------------------------------
    // 自動は、値の軸は値のフィールドの名前、項目の軸は見えているレベルの名前（ドリルした親のレベルは除く）
    const unitLabel = `${unit.unitWord}${(axisSettings.unitText.value ?? "").trim()}`;
    const titleOf = (styleValue: string, text: string, unitText: string): string => {
        if (styleValue === TITLE_STYLES.showUnitOnly) return unitText;
        if (styleValue === TITLE_STYLES.showBoth) return unitText ? `${text} (${unitText})` : text;
        return text;
    };
    const valueTitleStyle = dropdownValue(axisSettings.titleStyle, TITLE_STYLES.showTitleOnly);
    // タイトルに単位を入れる（単位のみ・両方）ときは、上の単位のラベルを出さない（同じ単位を 2 か所に出さない）
    const unitInTitle =
        (axisSettings.titleShow.value ?? true) &&
        unitLabel !== "" &&
        (valueTitleStyle === TITLE_STYLES.showUnitOnly || valueTitleStyle === TITLE_STYLES.showBoth);
    if (axisSettings.titleShow.value ?? true) {
        const text = titleOf(
            valueTitleStyle,
            (axisSettings.titleText.value ?? "").trim() || parsed.valueName,
            unitLabel
        );
        style.valueAxis.title = text ? { text, font: fontOf(axisSettings.titleFont, ink(colorOf(axisSettings.titleColor) || "#252423")) } : null;
    }
    const categorySettings = settings.categoryAxis;
    if (categorySettings.titleShow.value ?? true) {
        const text = titleOf(
            dropdownValue(categorySettings.titleStyle, TITLE_STYLES.showTitleOnly),
            (categorySettings.titleText.value ?? "").trim() || parsed.levelNames.slice(commonDepth).join(LEVEL_SEPARATOR),
            ""
        );
        style.categoryAxis.title = text
            ? { text, font: fontOf(categorySettings.titleFont, ink(colorOf(categorySettings.titleColor) || "#252423")) }
            : null;
    }

    // --- 凡例 ---------------------------------------------------------------------------------
    // 標準と同じ「増加・減少・合計（・区切りの差・その他）」。棒にある種類だけ出す。
    // 系列で積むときは、増減の棒は系列の色なので「系列 ＋ 合計（・区切りの差）」
    const present = (kind: BarKind) => bars.some((b) => b.kind === kind);
    // 区切りの差は 1 項目。印は、グラフにある符号の色（両方あれば左半分が増加、右半分が減少）を透かして描く
    const subtotalLegend = (): LegendItem[] => {
        const diffs = bars.filter((b) => b.kind === BAR_KINDS.subtotal);
        if (!diffs.length) return [];
        const split = [
            ...(diffs.some((b) => b.value >= 0) ? [colors.increase] : []),
            ...(diffs.some((b) => b.value < 0) ? [colors.decrease] : []),
        ];
        return [{ name: LEGEND_NAMES.subtotal, color: split[0], split, opacity: style.subtotalOpacity, selectionId: null, kind: BAR_KINDS.subtotal }];
    };
    const legend: LegendItem[] = [
        ...(hasSeries
            ? series.map((s, i): LegendItem => ({ name: s.name, color: seriesColors[i], selectionId: s.selectionId, kind: "series" }))
            : []),
        ...(!hasSeries && present(BAR_KINDS.increase)
            ? [{ name: LEGEND_NAMES.increase, color: colors.increase, selectionId: null, kind: BAR_KINDS.increase } as LegendItem]
            : []),
        ...(!hasSeries && present(BAR_KINDS.decrease)
            ? [{ name: LEGEND_NAMES.decrease, color: colors.decrease, selectionId: null, kind: BAR_KINDS.decrease } as LegendItem]
            : []),
        ...(present(BAR_KINDS.total)
            ? [{ name: LEGEND_NAMES.total, color: colors.total, selectionId: null, kind: BAR_KINDS.total } as LegendItem]
            : []),
        ...subtotalLegend(),
        ...(!hasSeries && present(BAR_KINDS.others)
            ? [{ name: othersLabel, color: colors.others, selectionId: null, kind: BAR_KINDS.others } as LegendItem]
            : []),
        // 目標の線は最後。印は線（グラフの線と同じ色・種類）、名前は目標のメジャーの表示名
        ...(target
            ? [
                  {
                      name: target.name,
                      color: target.color,
                      selectionId: null,
                      kind: "target",
                      line: {
                          width: Math.min(target.width, LEGEND_LINE_MAX),
                          dash: dashOf(dropdownValue(settings.target.lineStyle, LINE_STYLES.dashed), Math.min(target.width, LEGEND_LINE_MAX)),
                      },
                  } as LegendItem,
              ]
            : []),
    ];
    if (settings.legend.titleShow.value ?? false) {
        style.legend.title = (settings.legend.titleText.value ?? "").trim() || parsed.seriesName;
    }

    return {
        isEmpty: false,
        message: "",
        notice: parsed.truncated ? TRUNCATED_NOTICE : "",
        shape,
        bars,
        events: eventNames,
        leftEvent: eventNames[leftIndex] ?? "",
        rightEvent: eventNames[rightIndex] ?? "",
        axis,
        unitBadge:
            (axisSettings.unitShow.value ?? true) && !unitInTitle ? unitBadgeOf(unit.unitWord, axisSettings.unitText.value ?? "") : "",
        legend,
        target,
        constantLine,
        hasHighlights,
        truncated: parsed.truncated,
        format: {
            events: eventNames,
            leftEvent: eventNames[leftIndex] ?? "",
            rightEvent: eventNames[rightIndex] ?? "",
            hasOrigin: !!origin,
            hasTarget: !!parsed.target,
            seriesTargets: hasSeries
                ? series.map((s, i) => ({
                      name: s.name,
                      color: seriesColors[i],
                      selector: s.selectionId?.getSelector() ?? {},
                  }))
                : [],
            horizontal: style.orientation === ORIENTATIONS.horizontal,
            subtotalDiff: bars.some((b) => b.kind === BAR_KINDS.subtotal),
        },
        style,
    };
}
