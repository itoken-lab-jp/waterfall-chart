"use strict";

/**
 * 座標の計算。縦向き・横向きのどちらも、ここで画面の座標まで出す（App.tsx は SVG に写すだけ）。
 *
 * - 縦：項目が左から右、値が下から上（範囲の反転で上から下）。項目名は下。入りきらなければ斜め（-45°）にし、
 *   それでも重なれば間引く。帯が「カテゴリの最小幅」より狭くなると横にスクロールする（標準と同じ）
 * - 横：項目が上から下、値が左から右。項目名は左。帯が 2 行ぶんあれば上の階層を 2 行目に、
 *   無ければ 1 行で上の階層から出し、入りきらないところは前を「…」で省く。帯が狭いと縦にスクロールする
 *
 * スクロールするときの座標：棒・項目名・グリッド線などは「中身」の座標（ビューがスクロールの長さまで広がったときの座標）で出し、
 * App が scroll の範囲だけ見せる。値の軸の目盛り・タイトル・凡例・単位・目標と定数線の名前はスクロールしない
 */

import powerbi from "powerbi-visuals-api";

import { AxisInfo, BAR_KINDS, Bar, LegendItem, Segment, TextStyle, TooltipItem, ViewModel, blend, extentOf } from "./viewModel";
import { FontSpec, headWithEllipsis, measureTextWidth, truncateToWidth } from "./unitUtils";
import { autoLabelColor } from "./shared/color";
import { BREAK_STYLES, LABEL_POSITIONS, ORIENTATIONS } from "./settings";
import { LEVEL_SEPARATOR } from "./data";
import { Candidate, LabelRequest, Obstacle, placeLabels } from "./placement";
import { PT_TO_PX } from "./shared/text";
import { rotatedLabelLeadIn } from "./shared/rotatedLabels";
import { Halo, haloStyle } from "./shared/halo";
import { legendBaselineOf, legendLineGlyphWidth, legendMarkerRadius, lineInkSpan } from "./shared/legend";
import { layoutLegend } from "./shared/legendLayout";
import { Rect, CornerSide, Corner, cornerRadiusOf, barPath } from "./shared/barShape";

export type { Rect, CornerSide, Corner };
export { cornerRadiusOf, barPath };
import { LegendArrow, LegendPaging } from "./shared/LegendArrowButton";
import { recommendedTickCount, tickTargetOf, fitTicks, endMarginOf, rowTicks, nextColumnTry, ColumnTry } from "./shared/ticks";
import { placeReferenceLabels } from "./shared/referenceLine";
import { maxOf, minOf } from "./shared/minMax";
import { axisColumnWidths, truncateUnitText } from "./shared/unitBadge";
import { titleInkOf, valueTitleBaselineFromEdge, valueTitleBaselineToColumn } from "./shared/valueTitleGap";
import {
    AXIS_TEXT_BELOW,
    AXIS_TEXT_DROP,
    AXIS_TITLE_ASCENT,
    axisTextBottomOf,
    axisTextDescentOf,
    BAR_TEXT_BOTTOM,
    BAR_TITLE_BOTTOM,
    bottomTitleBelowOf,
    COLUMN_TEXT_BOTTOM,
    COLUMN_TITLE_BOTTOM,
    H_LABEL_BACKGROUND_PAD,
    SIDE_NAME_LEAD,
    SIDE_NAME_PLOT_GAP,
    SIDE_TITLE_NAME_GAP,
    sidePlotLeftOf,
    sideTitleBaselineOf,
    sideTitleInkRightOf,
} from "./shared/axisBands";
import {
    DRILL_ROW_SHARE_MAX,
    DRILL_TEXT_ASCENT,
    LABEL_ROOM_SHARE_MAX,
    OUTSIDE_LABEL_PAD,
    TOP_ROW_RISE,
    TOP_TEXT_GAP,
    UNIT_TEXT_ASCENT,
    columnBadgeRiseOf,
    outsideTopHeadroomOf,
    plainTopOf,
    topRowNeedOf,
} from "./shared/plotTop";

export { PT_TO_PX, recommendedTickCount };

import IViewport = powerbi.IViewport;

/** 書式の文字サイズはポイント。SVG にはピクセルで渡す */
export const FONT_FAMILY = '"Segoe UI", wf_segoe-ui_normal, helvetica, arial, sans-serif';
/** 軸を切ったとき、切った側の端にとる印の余白 */
export const BREAK_ZONE = 16;
/** 印の隙間（背景の色）の太さと、斜線の傾き・波の振れ幅 */
const BREAK_GAP = 3;
const BREAK_SLANT = 4;
const BREAK_WAVE = 1.5;
const SIN45 = Math.SQRT1_2;
/** 斜めの項目名と描く範囲の下端のあいだ */
const ROTATED_GAP = 8;
/**
 * スクロールしないとき、斜めの項目名の左端とビジュアルの左端（左に凡例を置けば凡例の分の右の端）のあいだに残す余白（px）。
 * 20px（Power BI はビジュアルの枠で描画を切り抜くので、枠線・角丸の下で字が欠けないよう見込む）。
 * 字が大きいときは、この余白と文字の半分の大きいほうをとる（rotatedLeftMarginOf）。1.15.8.42 までは 4px（と文字の半分）
 */
const ROTATED_LEFT_MARGIN = 20;
/**
 * スクロールしない斜めの項目名の左の端の余白（px）。ROTATED_LEFT_MARGIN と文字の半分の大きいほう。名前の字の左の端は
 * 「中心 − 長さ × sin45 − 文字の半分」と見積もるので、文字の半分が余白より大きい（約 30pt を超える）と、余白まで省いた名前も
 * 字の左の端が左の端より外へ出て、出さなかった（1.15.8.43 の初めの作り）。30pt 以下は変わらない
 */
const rotatedLeftMarginOf = (fontSize: number) => Math.max(ROTATED_LEFT_MARGIN, fontSize * 0.5);
const NOTICE_HEIGHT = 16;
/**
 * 横向きの上の行（ドリルの位置・単位のラベル）の上に空ける余白（上に凡例を置けば、凡例の箱の下から）。題名が折り返したときに、
 * 単位のラベルが題名の最後の行にかぶって見えないように空ける（1.3.0.0 までは単位が上端から 2px ほどだった）。
 * 上の凡例は共通の決まり（shared/legend）で、ビューの上の端から箱の中の余白をとって置く（1.15.8.26 までは、上の端から 8px ＋ 4px に置いていた）。
 * 縦向きの描く範囲の上の行（単位のラベル・ドリルの位置）は、字の四角の上を TOP_TEXT_GAP 空ける（標準と同じ。shared/plotTop）。
 * ビジュアルはビューの外（上）には描いていない（.wf-container が viewport の大きさで overflow: hidden）
 */
const TOP_MARGIN = 8;
/** スクロール バーの太さぶん、描く範囲を空ける */
export const SCROLLBAR = 12;
/** 描く範囲の幅・高さがこれより小さければ、図を描かない */
const MIN_PLOT = 10;
/**
 * 横向きの描く範囲の右の端（ラベルを出すなら、一番広いラベルを右の端まで伸びる棒に付けたときの字の送りの右の端）と、
 * ビューの右の端の間 (px)。標準の集合横棒と同じ
 * （2026-10-20 の Desktop で、ラベルなしの 800 の線も、ラベルの字の送りの右の端も、ビューの右の端から約 10px）
 */
const H_PLOT_RIGHT = 10;
/** データ ラベルの背景の余白 */
const LABEL_PADDING = 2;

/**
 * データ ラベルを棒の端（つなぐ端）に付けて置くときの数。標準の集合縦棒・集合横棒を Desktop（全般のレスポンシブはオフ）で
 * 測った置き方に合わせた。標準の集合縦棒の 24pt・内側上では、棒の上の端から数字のいちばん上まで 10.7px だった。
 * 下の数では 10.86px（数字の高さを 0.7 文字とみて 8.3 ＋ (0.78 − 0.7) × 32px）、9pt では 9.26px になる。
 * 標準のウォーターフォールの 24pt・内側上は 13.9px で、3px ほど広いが、集合縦棒の数にそろえる。
 *
 * 値の向きに文字の高さが並ぶラベル（縦向きの棒の横書き、横向きの棒の縦書き。値と率の 2 行はまとまりで見る）：
 * - 数字の四角（1 行目のベースライン − 0.78 文字 〜 最後の行のベースライン ＋ 0.08 文字）の、つなぐ端に近い側を端から 8.3px 離す
 * - 描く物（背景を描くなら背景の四角、描かないなら文字の四角）の近い側が端から 1px より近づくなら、1px 離れる所まで離す
 * - 棒の中に入るのは「端から文字の四角の遠い側まで ＋ 2 ≤ 棒の長さ」のとき。内側中央は「数字の四角の高さ ≤ 棒の長さ」
 * 値の向きに文字の幅が並ぶラベル（横向きの棒の横書き、縦向きの棒の縦書き）：
 * - 文字の端を端から 7.5px 離す。棒の中に入るのは「7.5 ＋ 文字の幅 ＋ 1 ≤ 棒の長さ」、内側中央は「文字の幅 ≤ 棒の長さ」
 */
const LABEL_EDGE = {
    /** 数字の四角の近い側と端のあいだ */
    digitGap: 8.3,
    /** 数字の四角の上と下（ベースラインから、文字の大きさに対する割合） */
    digitAscent: 0.78,
    digitDescent: 0.08,
    /** 描く物の近い側と端のあいだの最小 */
    inkGap: 1,
    /** 棒の中に入るときに、文字の四角の遠い側に残す長さ */
    fitSlack: 2,
    /** 文字の幅が並ぶラベルの、文字の端と端のあいだ */
    textGap: 7.5,
    /** 文字の幅が並ぶラベルが棒の中に入るときに残す長さ */
    textFitSlack: 1,
} as const;
/** 文字の四角の上と下（ベースラインから、文字の大きさに対する割合。boxOf・backgroundOf と同じ見積もり） */
const TEXT_ASCENT = 0.8;
const TEXT_DESCENT = 0.22;

/**
 * 値の向きに文字の高さが並ぶラベルを、端 edge に付けて置く。値の向きの座標で、1 行目のベースライン（baseline）と、
 * 端から文字の四角の遠い側までの長さ（reach）を返す。dir はラベルが端のどちら側か（正なら座標の大きい側＝文字の下の向き、負なら小さい側）。
 * span は 1 行目から最後の行までのベースラインのあいだ、inkPad は描く物が文字の四角の外に出る分（背景の余白。無ければ 0）
 */
function attachedBaseline(edge: number, dir: number, size: number, span: number, inkPad: number): { baseline: number; reach: number } {
    const e = LABEL_EDGE;
    if (dir > 0) {
        // 近い側は数字の四角の上（1 行目）
        let baseline = edge + e.digitGap + e.digitAscent * size;
        const ink = baseline - TEXT_ASCENT * size - inkPad;
        if (ink < edge + e.inkGap) baseline += edge + e.inkGap - ink;
        return { baseline, reach: baseline + span + TEXT_DESCENT * size - edge };
    }
    // 近い側は数字の四角の下（最後の行）
    let baseline = edge - e.digitGap - e.digitDescent * size - span;
    const ink = baseline + span + TEXT_DESCENT * size + inkPad;
    if (ink > edge - e.inkGap) baseline -= ink - (edge - e.inkGap);
    return { baseline, reach: edge - (baseline - TEXT_ASCENT * size) };
}


/** SVG の文字の見た目（ピクセル） */
export interface RenderFont {
    family: string;
    size: number;
    weight: "normal" | "bold";
    style: "normal" | "italic";
    decoration: "none" | "underline";
    color: string;
}

/** 角を丸める端（画面の向き） */

export interface SegmentLayout {
    segment: Segment;
    rect: Rect;
    /** ハイライトの該当分。無ければ null */
    highlight: Rect | null;
    /** 角を丸める端（1 つか、浮いた棒は両端）と半径。丸めなければ null */
    corner: Corner | null;
    /** ハイライトの該当分の角（区画と同じ端を、該当分の長さに収まる半径で丸める）。丸めなければ null */
    highlightCorner: Corner | null;
}


export interface TextLayout {
    x: number;
    y: number;
    anchor: "start" | "middle" | "end";
    lines: string[];
    /** 回転（度）。(x, y) を中心に回す */
    rotate: number;
    font: RenderFont;
    lineHeight: number;
    /** 背景（(x, y) を原点にした、回す前の座標）。無ければ null */
    background?: (Rect & { color: string; opacity: number }) | null;
    /** 文字の縁。無ければ付けない */
    halo?: Halo;
}

export interface BarLayout {
    bar: Bar;
    /** 棒が描く範囲（フォーカスの枠に使う） */
    rect: Rect;
    segments: SegmentLayout[];
    label: TextLayout | null;
    category: (TextLayout & { title: string }) | null;
}

export interface Line {
    x1: number;
    y1: number;
    x2: number;
    y2: number;
}

export interface BreakMark {
    /** 背景色で塗る隙間 */
    gap: string;
    /** 隙間の両側の細い線 */
    edges: string[];
    /** 値の方向の座標で、印の真ん中と広がり（テスト・重なりの確かめ用） */
    center: number;
    extent: [number, number];
}

export interface LegendLayout extends LegendPaging {
    title: TextLayout | null;
    /**
     * y は色の四角の上端（線の印は y + swatch / 2 に引く）、textY は名前のベースライン、swatchWidth は印の幅（目標の線は、模様と端の形で
     * 見える左右の端の幅）、textX は名前の字の送りの左の端（印の見える右の端から legendMarkerGap）。lineX1・lineX2 は目標の線の両端
     * （線の見える左の端が x に来る所から線の印の長さ。破線・点線は最後のダッシュが線の終わりより手前で終わるので、線の終わりは
     * 印の右の端より外にあることがある）。棒の項目は無し
     */
    items: Array<{ item: LegendItem; index: number; x: number; swatchWidth: number; y: number; textX: number; textY: number; text: string; lineX1?: number; lineX2?: number }>;
    font: RenderFont;
    /** 色の四角の一辺（px）。文字の大きさに比例する */
    swatch: number;
    /**
     * 凡例の箱（ビューの座標）。凡例はこの大きさで切り抜いて描く（縮めても入らない名前・タイトルが箱からはみ出しても、
     * 描く範囲・軸・項目名に重ねず、はみ出した所でマウスを取らない）
     */
    box: Rect;
}

export interface ReferenceLine {
    line: Line;
    color: string;
    width: number;
    dash: string | null;
    cap?: "round" | "square" | "butt";
    /** 名前と値（スクロールしない側に置く）。出さなければ null */
    label: TextLayout | null;
    /** 線にマウスを当てたときのツールヒント。空なら出さない（定数線） */
    tooltip: TooltipItem[];
    /** 線の不透明度（定数線の透過性）。無ければ 1 */
    opacity?: number;
    /** 棒の後ろ（back）か前（front）に描く。無ければ前 */
    layer?: "back" | "front";
    /** 網掛け領域（中身の座標）と塗り。無ければ塗らない */
    shade?: (Rect & { color: string; opacity: number }) | null;
}

export interface ChartLayout {
    horizontal: boolean;
    /** 見えている描く範囲（スクロールしても動かない） */
    plot: Rect;
    /** スクロール。axis が null ならスクロールしない。region は見せる窓（ビューの座標）、content は中身の長さ（スクロールの方向） */
    scroll: { axis: "x" | "y" | null; region: Rect; content: number };
    bars: BarLayout[];
    ticks: TextLayout[];
    gridlines: Line[];
    breaks: BreakMark[];
    connectors: Line[];
    target: ReferenceLine | null;
    constantLine: ReferenceLine | null;
    legend: LegendLayout | null;
    valueTitle: TextLayout | null;
    categoryTitle: TextLayout | null;
    badge: TextLayout | null;
    /** ドリルダウンした位置。ドリルしていない・出さないときは null */
    drillPath: (TextLayout & { title: string }) | null;
    notice: { x: number; y: number } | null;
    /** 項目名を斜めにした（縦向きで入りきらない） */
    rotatedCategories: boolean;
    /** 項目名を間引いた間隔（1 = 全部出す） */
    categoryStep: number;
}

export const renderFontOf = (style: TextStyle, color = style.color): RenderFont => ({
    family: style.family || FONT_FAMILY,
    size: style.size * PT_TO_PX,
    weight: style.bold ? "bold" : "normal",
    style: style.italic ? "italic" : "normal",
    decoration: style.underline ? "underline" : "none",
    color,
});

const specOf = (font: RenderFont): FontSpec => ({
    family: font.family,
    size: font.size,
    bold: font.weight === "bold",
    italic: font.style === "italic",
});

/** 1 行で出すときの名前。上の階層（どの項目でも同じレベルは除いたもの）から並べる */
const pathOf = (b: Bar): string => (b.parentLabel ? `${b.parentLabel}${LEVEL_SEPARATOR}${b.label}` : b.label);
// 1 行の項目名（斜め・横向き）は、入らなければ標準の集合縦棒と同じく末尾を「…」で省く（「海外（…」。truncateToWidth）。
// 上の階層から並べた名前（「親 / 子」）も末尾を省く（標準は階層を展開して名前を連結しても末尾を省く。子の名前が見えなくなることがある）。
// 1.15.8.42 までは頭を省いた（「…エリア」。一番下のレベルの名前を残すため）

/**
 * 切った印の形。u は棒の幅の方向（u0〜u1）、t は値の方向の座標で、center を真ん中に置く。
 * 斜線は幅いっぱいで BREAK_SLANT だけ傾いた 2 本の線、波線は半波 4 つの波。あいだを背景の色で抜く
 */
export function breakMark(style: string, horizontal: boolean, center: number, u0: number, u1: number): BreakMark {
    const point = (u: number, t: number) => (horizontal ? `${t},${u}` : `${u},${t}`);
    const half = BREAK_GAP / 2;
    if (style === "wave") {
        const count = 4;
        const step = (u1 - u0) / count;
        /** t の位置の波。forward なら u0 → u1、そうでなければ u1 → u0 へ描く */
        const wave = (t: number, forward: boolean, move: string) => {
            const sign = (k: number) => (k % 2 === 0 ? -1 : 1);
            if (forward) {
                let d = `${move}${point(u0, t)} Q${point(u0 + step / 2, t + sign(0) * BREAK_WAVE * 2)} ${point(u0 + step, t)}`;
                for (let k = 2; k <= count; k++) d += ` T${point(u0 + k * step, t)}`;
                return d;
            }
            const last = count - 1;
            let d = `${move}${point(u1, t)} Q${point(u1 - step / 2, t + sign(last) * BREAK_WAVE * 2)} ${point(u1 - step, t)}`;
            for (let k = 2; k <= count; k++) d += ` T${point(u1 - k * step, t)}`;
            return d;
        };
        const top = center - half;
        const bottom = center + half;
        return {
            gap: `${wave(top, true, "M")} ${wave(bottom, false, "L")} Z`,
            edges: [wave(top, true, "M"), wave(bottom, true, "M")],
            center,
            extent: [top - BREAK_WAVE * 2, bottom + BREAK_WAVE * 2],
        };
    }
    // 斜線：「/」の向き（縦向きは右上がり、横向きは上が右）
    const lean = BREAK_SLANT / 2;
    const at = (u: number, offset: number) => point(u, center + offset + (u === u0 ? lean : -lean));
    return {
        gap: `M${at(u0, -half)} L${at(u1, -half)} L${at(u1, half)} L${at(u0, half)} Z`,
        edges: [`M${at(u0, -half)} L${at(u1, -half)}`, `M${at(u0, half)} L${at(u1, half)}`],
        center,
        extent: [center - half - lean, center + half + lean],
    };
}

/**
 * 背景の長方形（(x, y) を原点に、回す前の座標）。文字の寄せ方と行数から出す。padX は字の送りの左右の端から外の余白
 * （横向きの棒の横書きのデータ ラベルは H_LABEL_BACKGROUND_PAD。ほかは上下と同じ LABEL_PADDING）
 */
function backgroundOf(
    lines: string[],
    anchor: TextLayout["anchor"],
    font: RenderFont,
    lineHeight: number,
    fill: { color: string; opacity: number } | null,
    padX = LABEL_PADDING
): TextLayout["background"] {
    if (!fill) return null;
    const width = Math.max(0, maxOf(lines.map((l) => measureTextWidth(l, specOf(font))))) + padX * 2;
    const x = anchor === "start" ? -padX : anchor === "end" ? -width + padX : -width / 2;
    const top = -font.size * 0.8 - LABEL_PADDING;
    const height = font.size * 1.02 + (lines.length - 1) * lineHeight + LABEL_PADDING * 2;
    return { x, y: top, width, height, color: fill.color, opacity: fill.opacity };
}

/** 長方形どうしが重なるか（辺が接するだけなら重ならない） */
const hitsBox = (a: Rect, b: Rect): boolean =>
    a.x < b.x + b.width && a.x + a.width > b.x && a.y < b.y + b.height && a.y + a.height > b.y;

/** 四角を障害物にする */
const asBox = (box: Rect): Obstacle => ({ kind: "box", box });

/** 点線の枠の小計の地の濃さ（差の色を背景に混ぜる割合） */
const SUBTOTAL_TINT = 0.15;

/**
 * 比べる形の小計の塗り方。color は差の色（増加・減少）。
 * fill は棒の塗り（hatch のときは斜線の模様を App が敷く）、stroke・dash は枠、under は棒の中の見た目の色（ラベルの文字色を決める）
 */
export function subtotalPaint(
    style: Pick<ViewModel["style"], "subtotalStyle" | "subtotalOpacity" | "background">,
    color: string
): { fill: string; fillOpacity: number; stroke: string | null; dash: string | null; hatch: boolean; under: string } {
    const bg = style.background;
    switch (style.subtotalStyle) {
        case "fill":
            return {
                fill: color,
                fillOpacity: style.subtotalOpacity,
                stroke: null,
                dash: null,
                hatch: false,
                under: style.subtotalOpacity < 1 ? blend(color, bg, 1 - style.subtotalOpacity) : color,
            };
        case "hatch":
            return { fill: bg, fillOpacity: 1, stroke: color, dash: null, hatch: true, under: bg };
        case "dashed": {
            const tint = blend(color, bg, 1 - SUBTOTAL_TINT);
            return { fill: tint, fillOpacity: 1, stroke: color, dash: "3 2", hatch: false, under: tint };
        }
        default:
            return { fill: bg, fillOpacity: 1, stroke: color, dash: null, hatch: false, under: bg };
    }
}

/**
 * 文字の外枠（画面の座標）。pad は四方に足す余白（padX を渡せば、字の送りの左右だけその余白）。回転は 0° と -90° だけ見る
 * （ラベル・目標と定数線の名前）。高さは backgroundOf と同じ見積もり（上に 0.8 文字、下に 0.22 文字）
 */
export function boxOf(text: TextLayout, pad: number, padX = pad): Rect {
    const spec = specOf(text.font);
    const width = Math.max(0, maxOf(text.lines.map((l) => measureTextWidth(l, spec)))) + padX * 2;
    const x0 = text.anchor === "start" ? -padX : text.anchor === "end" ? -width + padX : -width / 2;
    const y0 = -text.font.size * 0.8 - pad;
    const height = text.font.size * 1.02 + (text.lines.length - 1) * text.lineHeight + pad * 2;
    // -90° は (dx, dy) → (dy, -dx)
    if (text.rotate === -90) return { x: text.x + y0, y: text.y - x0 - width, width: height, height: width };
    return { x: text.x + x0, y: text.y + y0, width, height };
}

/** ラベルの四角に足す余白（背景の無いラベル）。ほかの物とくっつきすぎないように */
const HIT_PAD = 1;

/** 文字の縁が文字の四角の外へ出る幅（縁の太さの半分）。文字と縁が同じ色で縁を描かないなら 0 */
function haloOutsetOf(text: TextLayout): number {
    return text.halo && haloStyle(text.halo, text.font.color).strokeWidth ? text.halo.width / 2 : 0;
}

/**
 * ほかのラベルとの重なりを見る四角。背景を描くなら背景の四角、描かないなら文字の四角に HIT_PAD を足したもの。
 * 文字の縁も描くものなので入れる（背景からはみ出す分も）。棒や線とは文字の四角で見る（縁は棒や線に掛けてよい）
 */
export function hitBoxOf(text: TextLayout): Rect {
    const halo = haloOutsetOf(text);
    if (!text.background) return boxOf(text, HIT_PAD + halo);
    // 背景の左右の余白は背景の幅から（横向きの棒の横書きのデータ ラベルは上下より広い）
    const textWidth = Math.max(0, maxOf(text.lines.map((l) => measureTextWidth(l, specOf(text.font)))));
    const pad = Math.max(LABEL_PADDING, halo);
    return boxOf(text, pad, Math.max((text.background.width - textWidth) / 2, halo));
}

/** 名前の側の余白を取り直す回数の上限 */
const NEAR_PASSES = 3;
/** 置いたデータ ラベルの数 */
const shownLabels = (layout: ChartLayout): number => layout.bars.filter((b) => b.label).length;
/**
 * 空け直した結果 next を採るか。出るラベルが増えたとき、または数が同じで好ましい置き場所（最初の候補）に置けたラベルが増えたとき
 */
const betterPlaced = (next: { layout: ChartLayout; firstPlaced: number }, current: { layout: ChartLayout; firstPlaced: number }): boolean => {
    const more = shownLabels(next.layout) - shownLabels(current.layout);
    return more > 0 || (more === 0 && next.firstPlaced > current.firstPlaced);
};
/** 値の軸の範囲を、描く範囲の長さで決まる目盛りの本数で丸め直す回数の上限（丸め直して長さが変わり、本数が行き来しても止める） */
const AXIS_PASSES = 2;

/**
 * 座標の計算。描けないほど狭いとき・棒が無いときは null。
 *
 * 1 回目は項目名の側（縦は下、横は左）に余白をとらずに置く。値の端にある棒のラベルが項目名の側へはみ出すなら、
 * はみ出した分だけ描く範囲と項目名のあいだを空けて置き直す（ラベルが項目名に重ならないように。空けるのは要るときだけ）。
 * 空け直した結果は、出るラベルが増えたとき、または数が同じで好ましい置き場所（最初の候補）に置けたラベルが増えたときだけ採る
 * （1.15.8.30 までは何も増えなくても採り、描く範囲だけが狭くなる形があった）
 */
/** legendPage は、凡例が入りきらないときに出すページ（0 から。範囲の外なら端のページ） */
export function layoutOf(viewModel: ViewModel, viewport: IViewport, legendPage = 0): ChartLayout | null {
    // 空けるのはビューの 1/4 まで。それでも入らないラベルは出さない
    const cap = (viewModel.style.orientation === ORIENTATIONS.horizontal ? viewport.width : viewport.height) * 0.25;
    let tried = layoutRounded(viewModel, viewport, 0, legendPage);
    let result = tried.result;
    // 空け直した結果は、出るラベルか、最初の候補に置けたラベルが増えたときだけ採る（betterPlaced。どちらも増えなければ、
    // 空けずに組んだ描く範囲の高さ・幅を残す。数だけを見ると、空けて棒の外に置けるラベルが棒の中に入ったままになる）。
    // 採らなかった結果にはみ出しが残れば、そこからさらに空けてみる（1 回目で増えなくても、もう少し空ければ入るラベルがある）
    let room = 0;
    for (let pass = 0; result && tried.result && tried.result.nearOverflow > 0 && pass < NEAR_PASSES; pass++) {
        const wider = Math.min(cap, room + Math.ceil(tried.result.nearOverflow) + 1);
        if (wider <= room) break;
        // 前に丸めた軸から組む（描く範囲の長さが変わって本数が変われば、丸め直す）
        const next = layoutRounded(tried.viewModel, viewport, wider, legendPage);
        // 空けると描く範囲が狭くなりすぎるときは、はみ出すラベルを出さない版のまま
        if (!next.result) break;
        room = wider;
        tried = next;
        if (betterPlaced(next.result, result)) result = next.result;
    }
    // 描く範囲が足りない・位置の文字が 1 文字も入らないときは、ドリルの位置をとらずに組み直す
    // （位置の行のせいで図ごと消えない・空の位置で行の高さだけとらないように）
    if (viewModel.drillPath && !result?.layout.drillPath) return layoutOf({ ...viewModel, drillPath: "" }, viewport, legendPage);
    // 延ばせない端のラベルのために描く範囲の上の端を下げて図が消えるなら、横にスクロールする図は、下げる所をスクロール バーの分まで上げて
    // 組み直す（keepPlot）。目盛りの列を単位の幅まで広げて横にスクロールするようになった低い図は、下にスクロール バーの分をとるので、
    // 前の版の上の端まで下げると描く範囲が残らない。組み直した図も、描く範囲の長さで決まる本数で丸め直す（layoutRounded）
    if (!result) result = layoutRounded(viewModel, viewport, 0, legendPage, true).result;
    return result?.layout ?? null;
}

/**
 * 値の軸の範囲を、描く範囲の長さで決まる目盛りの本数の目安（目盛りと同じ本数。tickTargetOf）で丸めて組む（標準と同じ）。
 * 描く範囲の長さは組んでみないと分からないので、組んで決まった本数で軸を丸め直し（AxisInfo の roundedWith）、範囲が変われば組み直す。
 * 丸め直しても範囲が変わらなければ、そこで止める（ふつうは 1 回）。丸め直した軸の viewModel も返す（名前の側の余白を取り直すときに使う）。
 * AXIS_PASSES 回まで丸め直しても落ち着かない（丸め直すと描く範囲の長さが閾値をまたいで本数が行き来する）ときは、使った軸のまま、
 * 目盛りも軸を丸めた本数（roundCount）で引いて組み直す（範囲と目盛りを同じ本数から出す。描いた長さの本数のままだと、
 * 範囲を丸めた本数より少ない本数で引いて、目盛りが 1 本になることがある）
 */
function layoutRounded(
    viewModel: ViewModel,
    viewport: IViewport,
    nearRoom: number,
    legendPage: number,
    keepPlot = false
): { result: { layout: ChartLayout; nearOverflow: number; tickTarget: number; firstPlaced: number } | null; viewModel: ViewModel } {
    const sameRange = (a: AxisInfo, b: AxisInfo) => a.min === b.min && a.max === b.max;
    let current = viewModel;
    let result = layoutWith(current, viewport, nearRoom, legendPage, keepPlot);
    for (let pass = 0; result && pass < AXIS_PASSES; pass++) {
        const axis = current.axis.roundedWith?.(result.tickTarget);
        if (!axis || sameRange(axis, current.axis)) break;
        const next = layoutWith({ ...current, axis }, viewport, nearRoom, legendPage, keepPlot);
        if (!next) break;
        current = { ...current, axis };
        result = next;
    }
    const roundCount = current.axis.roundCount;
    if (result && roundCount !== undefined && result.tickTarget !== roundCount) {
        const drawn = current.axis.roundedWith?.(result.tickTarget);
        if (drawn && !sameRange(drawn, current.axis)) result = layoutWith(current, viewport, nearRoom, legendPage, keepPlot, roundCount) ?? result;
    }
    return { result, viewModel: current };
}

/**
 * 1 回分の座標。返す tickTarget は、描く範囲の長さで決まる目盛りの本数の目安（範囲を丸め直すのに使う）。
 * ticksAt を渡せば、目盛りはその本数の目安で引く（丸め直しが落ち着かないとき。layoutRounded）。
 * columnTry は、縦向きの目盛りの列の幅の組み直しの途中の状態（はじめは無し。下の nextColumnTry）
 */
function layoutWith(
    viewModel: ViewModel,
    viewport: IViewport,
    nearRoom: number,
    legendPage = 0,
    keepPlot = false,
    ticksAt?: number,
    columnTry?: ColumnTry
): { layout: ChartLayout; nearOverflow: number; tickTarget: number; firstPlaced: number } | null {
    const { style, axis, bars } = viewModel;
    if (!bars.length) return null;
    const horizontal = style.orientation === ORIENTATIONS.horizontal;
    const W = viewport.width;
    const H = viewport.height;
    const n = bars.length;

    const tickFont = renderFontOf(style.valueAxis.font);
    const categoryFont = renderFontOf(style.categoryAxis.font);
    const labelFont = renderFontOf(style.dataLabels.font);
    const legendFont = renderFontOf(style.legend.font);
    // 合計・小計の棒のラベルと項目名は別の文字（既定は太字）
    const isTotalBar = (b: Bar) => b.kind === BAR_KINDS.total || b.kind === BAR_KINDS.subtotal;
    const totalCategoryFont = renderFontOf(style.categoryAxis.totalFont);
    const categoryFontOf = (b: Bar) => (isTotalBar(b) ? totalCategoryFont : categoryFont);
    /** 項目名の行の高さ・文字の大きさ（大きいほうの文字） */
    const categorySize = Math.max(categoryFont.size, totalCategoryFont.size);
    const categoryLineHeight = categorySize + 3;
    const totalLabelFont = renderFontOf(style.dataLabels.totalFont);
    const labelFontOf = (b: Bar) => (isTotalBar(b) ? totalLabelFont : labelFont);
    /** ラベルの背景が文字の四角の外に出る分 */
    const labelInkPad = style.dataLabels.background ? LABEL_PADDING : 0;
    const valueTitle = style.valueAxis.title;
    const categoryTitle = style.categoryAxis.title;
    const valueTitleFont = valueTitle ? renderFontOf(valueTitle.font) : null;
    const categoryTitleFont = categoryTitle ? renderFontOf(categoryTitle.font) : null;
    /**
     * X 軸のタイトルの字の画素の上下の出の見積もり（titleInkOf。字・太字・下線（文字の大きさと画面の倍率で）から）。縦向きは、下へ下りる字・下線の下の端がビューに入るだけ
     * ベースラインの下をとり（bottomTitleBelowOf）、横向き（-90° に回した名前の側のタイトル）は、既定の組より高い字の分だけ内側に置く
     * （sideTitleBaselineOf）。既定の組の字だけなら標準の位置のまま。1.15.8.38 の最初の作りでは字によらず既定の組で置き、字の端が切れた
     */
    const categoryTitleInk =
        categoryTitle && categoryTitleFont
            ? titleInkOf(categoryTitle.text, categoryTitleFont.decoration === "underline", { fontPx: categoryTitleFont.size, bold: categoryTitleFont.weight === "bold" })
            : undefined;
    const valueTitleSize = valueTitleFont ? valueTitleFont.size + 6 : 0;
    /**
     * 縦向きの X 軸のタイトルの行（ベースラインより上の字の高さと、ベースラインからビューの下の端（下の凡例・知らせがあればその上の端）まで）。
     * 標準の集合縦棒と同じ（shared/axisBands）。1.15.8.37 までは文字の大きさ ＋ 6px（ベースラインは行の上の端から 2px ＋ 0.8 文字）
     */
    const categoryTitleSize = categoryTitleFont ? categoryTitleFont.size * AXIS_TITLE_ASCENT + bottomTitleBelowOf(COLUMN_TITLE_BOTTOM, categoryTitleFont.size, categoryTitleInk) : 0;
    /**
     * 横向きの名前の欄の左の端（基準の左の端から）。名前の側の並びは標準の集合横棒と同じ（shared/axisBands）：タイトルの字の画素の左の端を
     * 基準の左の端に置き、字の右の端から SIDE_TITLE_NAME_GAP で一番長い名前の字の左の端。名前の欄は左の端に SIDE_NAME_LEAD をとる。
     * 1.15.8.37 までは、タイトルの行に文字の大きさ ＋ 6px をとり（ベースラインは 4px ＋ 0.8 文字）、名前の欄の左右に 6px ずつとった
     */
    const sideTitleSize = categoryTitleFont ? sideTitleInkRightOf(categoryTitleFont.size, categoryTitleInk) + SIDE_TITLE_NAME_GAP - SIDE_NAME_LEAD : 0;
    /** 横向きの名前の欄の、一番長い名前の字の左右に足す幅（左の端に SIDE_NAME_LEAD、描く範囲の左の端まで SIDE_NAME_PLOT_GAP） */
    const namePads = SIDE_NAME_LEAD + SIDE_NAME_PLOT_GAP;
    const verticalText = style.dataLabels.orientation === "vertical";
    const hasParent = bars.some((b) => b.parentLabel);
    const invert = style.valueAxis.invert;
    const switchAxis = style.valueAxis.switchPosition;

    // --- 凡例 ------------------------------------------------------------------------------------
    // 組み方（行の高さ・余白・字と印の位置・2 行・ページを送る矢印・タイトル・左右の列の幅）は共通の決まり（shared/legendLayout）。
    // 知らせの行があれば、その上の範囲に置く。印は、棒の項目は四角、目標は横に引く線。目標の線の印の幅は、線の模様（破線・点線）と
    // 端の形で見える左右の端の幅（shared/legend の lineInkSpan。決まった線の印の長さより広くも狭くもなる）にして、名前を線の見える
    // 右の端から印と名前のあいだ、線の見える左の端を前の名前から項目のあいだに置く。どこも見えない線は決まった長さのまま
    const noticeHeight = viewModel.notice ? NOTICE_HEIGHT : 0;
    const swatch = legendMarkerRadius(legendFont.size) * 2;
    const lineGlyphLength = legendLineGlyphWidth(legendFont.size);
    const targetInkOf = (item: LegendItem) => (item.line ? lineInkSpan(lineGlyphLength, item.line.width, item.line.cap, item.line.dash) : null);
    const legendBox =
        style.legend.show && viewModel.legend.length > 0
            ? layoutLegend({
                  entries: viewModel.legend.map((item) => {
                      const ink = targetInkOf(item);
                      return { name: item.name, glyphWidth: item.kind === "target" ? (ink ? ink.end - ink.start : lineGlyphLength) : swatch };
                  }),
                  title: style.legend.title,
                  font: specOf(legendFont),
                  position: style.legend.position,
                  width: W,
                  height: H - noticeHeight,
                  page: legendPage,
                  rows: style.legend.rows,
              })
            : null;
    /**
     * 凡例の箱の分。描く範囲は、凡例を除いた所に、凡例が無いときと同じ決まりで置く。下の凡例も、箱のすぐ上を下の帯の下の端にする
     * （下の帯の作りは標準の集合縦棒・集合横棒と同じ）。1.15.8.39 までは、項目名の欄の下が 4px だった分を補うために
     * 箱の上に 10px（低い図では縮めた）の隙間をとり、1.15.8.40 の最初の作りでは下の帯を標準にそろえたあとも残っていて、下の帯が 10px 広すぎた
     */
    const legendReserve = legendBox?.reserve ?? { top: 0, bottom: 0, left: 0, right: 0 };
    // ドリルダウンした位置。縦向きは、単位のラベルを出さなくても描く範囲の上の行（単位のラベルと同じ行）に置き、行を積まない
    // （2026-09-25 ユーザー決定。縦に積むとグラフエリアが狭くなるため）。横向き（単位は右下）は、上の凡例の下に 1 行とる
    const drillPathFont = renderFontOf(style.categoryAxis.font);
    // 単位のラベルの文字（既定は目盛りと同じ）。「プロット エリアの右上」なら、横向きでも描く範囲の上の行に置く
    const badgeFont = renderFontOf(style.valueAxis.unitFont);
    const badgeTopRight = style.valueAxis.unitPosition === "plotTopRight";
    // 行の高さは、文字の下側（0.22 文字、boxOf と同じ見積もり）まで収まるようにとる（大きいフォントでも下の物に掛けない）
    const drillPathRow = viewModel.drillPath && horizontal ? Math.ceil(drillPathFont.size * 1.3) + 6 : 0;
    const reserve = {
        // 縦向きは、上の凡例の下の端（凡例が無ければビューの上の端）。描く範囲の上の端と上の行は、そこから測る（下の縦向きの組み方）。
        // 横向きは、その下に TOP_MARGIN とドリルの位置の行をとる
        top: horizontal ? legendReserve.top + TOP_MARGIN + drillPathRow : legendReserve.top,
        bottom: legendReserve.bottom,
        left: legendReserve.left,
        right: legendReserve.right,
    };
    // 定数線のデータ ラベルは標準と同じく描く範囲の中に置くので、余白は目標の名前だけ
    const referenceLabelRoom = viewModel.target?.label ? labelFont.size + 6 : 0;

    // 軸を切ったときは、切った側の端に印の余白をとる（目盛りはその内側に並ぶ）。
    // 「切った印」をなしにしたときは、標準と同じく印も余白もとらず、軸の外の分をそのまま省く
    const marked = style.breakStyle !== BREAK_STYLES.none;
    const lowZone = marked && axis.cut === "bottom" ? BREAK_ZONE : 0;
    const highZone = marked && axis.cut === "top" ? BREAK_ZONE : 0;
    /**
     * 出す目盛り。length は描く範囲の長さ（本数の目安を決める）、gap は 2 つの値の画面の上の間隔 (px)。
     * 目盛りの本数の目安は、標準と同じく描く範囲の長さで決める（縦は高さ 150・300px、横は幅 300・500px で 3・5・8 まで）。
     * 「目盛りの本数 (目安)」があればその本数。値の軸の範囲も、この本数の目安で丸める（layoutRounded）。丸め直しが落ち着かないときは、
     * 範囲を丸めた本数で引く（ticksAt）。グリッド線も同じ目盛り。
     * 縦は、目安を、目盛りを引く範囲（axis.min〜axis.max。棒の上の外のラベルの分だけ延ばした所と、軸を切った印の余白は入れない）の長さに
     * 数字の行（目盛りの文字の大きさ）が入る数までに下げて取り、隣との間隔が狭くてもそれ以上は減らさない（shared/ticks の rowTicks。
     * 標準の集合縦棒と同じ）。1.15.8.39 までは、隣の間隔が文字の 1.4 倍以上になるまで本数を減らし、36pt の目盛りで標準より少なかった。
     * 横は、数字が重なるとき（隣の間隔がいちばん長い数字 ＋ 8px より狭いとき）は本数を減らして取り直す（fitTicks。前のまま）
     */
    const ticksShownAt = (length: number, gap: (a: number, b: number) => number): typeof axis.ticks => {
        if (!axis.ticksFor) return axis.ticks;
        const target = ticksAt ?? tickTargetOf(style.valueAxis.tickCount, length, horizontal ? "horizontal" : "vertical");
        return horizontal
            ? fitTicks(
                  axis.ticksFor,
                  target,
                  (ticks) => minOf(ticks.slice(1).map((t, i) => gap(t.value, ticks[i].value))),
                  (ticks) => Math.max(0, maxOf(ticks.map((t) => measureTextWidth(t.label, specOf(tickFont))))) + 8
              )
            : rowTicks(axis.ticksFor, target, gap(axis.min, axis.max), tickFont.size);
    };
    /** 目盛りの数字の一番広い幅 */
    const widestTick = (ticks: typeof axis.ticks) => Math.max(0, maxOf(ticks.map((t) => measureTextWidth(t.label, specOf(tickFont)))));
    // 縦向きの目盛りの数字の列の幅。描く目盛りの数字の一番広い幅にする（標準と同じ）。描く目盛りは描く範囲の高さで決まり、描く範囲の高さは
    // 列の幅で変わる（描く範囲の幅で名前の回し方・下の帯が変わる）ので、はじめはビューの高さから 80px 引いた長さで描く目盛りの数字の幅で組み、
    // 描く範囲の高さで描く目盛りの数字の幅が違えば、その幅で組み直す（下の nextColumnTry。回数に上限があり、落ち着かないときの決め方も
    // shared/ticks）。1.15.8.39 までは、描かない「いちばん細かい 8 本」の目盛りの数字（0.5 おきの 1.5 など）の幅も入れ、描く数字のほうが
    // 狭い形で列が広がり、Y 軸のタイトルとのすき間が標準より広かった。横向きは列を使わない（値の軸は下か上）
    const tickLabelWidth = (() => {
        if (horizontal || !style.valueAxis.show) return 0;
        if (columnTry) return columnTry.widths[0];
        const length = Math.max(1, H - 80);
        const valueLength = Math.max(0, length - lowZone - highZone);
        const span = axis.max - axis.min;
        return widestTick(ticksShownAt(length, (a, b) => (span > 0 ? (Math.abs(a - b) / span) * valueLength : 0)));
    })();
    // 単位のラベル。Y 軸の値（目盛り）を切っても出す（データ ラベルとツールヒントの数は、既定で軸の表示単位のまま単位の字を付けずに出るので、
    // 単位を消すと数の単位がどこにも出なくなる。1.15.8.29 までは値を切ると単位も出さなかった）。
    // 縦向きの Y 軸の上の単位は、目盛りの列を単位の幅まで広げて置く（単位が軸のタイトルに掛からないように）。広げるのは、描く範囲の幅が
    // ビジュアルの幅の 4 分の 1 を下回らない所まで（shared/unitBadge。目盛りの数字の幅には上限を付けない）。止めた列に入らない単位は、
    // 末尾を「…」で切る（括弧の中を切る。「(千…)」も入らなければ出さない）。描く範囲の右の端にそろえて置く単位
    // （横向き・「プロット エリアの右上」）は左の端で止め、止めても右の端に入らなければ同じく切る
    const unitWidthOf = (text: string) => measureTextWidth(text, specOf(badgeFont));
    const columnBadge = !horizontal && !badgeTopRight;
    /**
     * 縦向きの、目盛りの列の幅 column のときの、値の軸の側の余白（凡例の内側。目盛りの列と描く範囲の間は 6px）。値の軸のタイトルは、
     * 目盛りの列（目盛りの数字と上の単位の広いほう）の字からタイトルの字までを標準の集合縦棒と同じすき間にして置く（shared/valueTitleGap。
     * 左の軸は目盛りが大きいほど少し広い）。タイトルのベースラインは下の valueTitleLayout と同じで、字の外側の端を標準と同じだけ凡例の内側の端
     * （凡例が無ければビューの端）から離す（valueTitleBaselineFromEdge）。1.15.8.29 までは、列の外に 4px、タイトルに文字の大きさと 6px を取った。
     * 1.15.8.31 までは、タイトルのベースラインを凡例の内側の端から、左は 4px と字の上側 0.8 文字、右は 4px と字の下側 0.22 文字に置いた
     */
    // タイトルの字の画素の上下の出は、字・太字・下線から見積もる（titleInkOf。既定の組の字だけなら標準の測りの全角の字と同じ）。既定の外の字
    // （下に伸びる字・高い字・付け記号のある英字など）か下線のあるタイトルは、ビューの端で切れないよう、その分だけ内側に置く。列の側は標準の位置を保ち、
    // 字の画素が目盛りの数字に 2px より寄るときだけ離す（valueTitleBaselineToColumn）
    const valueTitleInk =
        valueTitle && valueTitleFont
            ? titleInkOf(valueTitle.text, valueTitleFont.decoration === "underline", { fontPx: valueTitleFont.size, bold: valueTitleFont.weight === "bold" })
            : undefined;
    // タイトルの側に凡例があれば、凡例の字・印の内側の端から標準と同じだけ離す（右は凡例の分を除いた所の外へ出る）。1.15.8.43 までは、
    // 凡例があっても凡例の分を除いた所の端から凡例が無いときと同じだけ離し、タイトルの大きさで外側の空きを変えなかった
    const valueTitleBaseline =
        valueTitle && valueTitleFont
            ? valueTitleBaselineFromEdge(valueTitleFont.size, switchAxis, valueTitleInk, legendBox?.side === (switchAxis ? "right" : "left"))
            : 0;
    const valueSideOf = (column: number, shown: boolean) =>
        valueTitleFont
            ? valueTitleBaseline + valueTitleBaselineToColumn(valueTitleFont.size, tickFont.size, switchAxis, valueTitleInk) + column + 6
            : Math.max(8, shown ? column + 10 : 4);
    /** 縦向きの、目盛りの列の幅 column のときの描く範囲の幅 */
    const plotWidthWith = (column: number) => W - reserve.left - reserve.right - 8 - valueSideOf(column, true);
    const badgeText = !viewModel.unitBadge
        ? ""
        : columnBadge
          ? (() => {
                const [column] = axisColumnWidths(
                    [{ tick: tickLabelWidth, unit: unitWidthOf(viewModel.unitBadge) }],
                    plotWidthWith(tickLabelWidth),
                    W
                );
                return truncateUnitText(viewModel.unitBadge, column + 1e-6, unitWidthOf);
            })()
          : truncateUnitText(viewModel.unitBadge, W - reserve.left - reserve.right - 4 + 1e-6, unitWidthOf);
    const badgeShown = Boolean(badgeText);
    // 横向きの単位のラベルの行（縦向きは描く範囲の上の行に置き、行を足さない）
    const badgeHeight = badgeShown ? badgeFont.size + 6 : 0;
    const labelTextOf = (b: Bar) => (b.rateText ? `${b.labelText} (${b.rateText})` : b.labelText);
    const valueLabelWidth = style.dataLabels.show
        ? Math.max(
              0,
              maxOf(bars.map((b) => measureTextWidth(horizontal && !verticalText ? labelTextOf(b) : b.labelText, specOf(labelFontOf(b)))))
          )
        : 0;
    /**
     * 縦向きで、値の軸の画面の上の側の端に要るもの：棒の上の外に置くデータ ラベル（位置が自動か外側の先で、棒の先が画面の上を向く棒）と、
     * 目標の名前（線の上に置く）。end はラベルを付ける値、length はそこから描く範囲の上の端までに要る長さ (px)。
     * 長さは、描く物（背景を描くなら背景の四角、描かないなら文字の四角）の遠い側までに OUTSIDE_LABEL_PAD を足す。
     * 横書きは棒の先から字の四角の上まで（LABEL_EDGE の置き方）で、さらに outsideTopHeadroomOf を足す（標準の集合縦棒と同じ）。
     * 縦書きは文字の端までのすき間と文字の幅
     */
    const topNeeds: Array<{ end: number; length: number }> = [];
    if (!horizontal) {
        const position = style.dataLabels.position;
        if (style.dataLabels.show && (position === LABEL_POSITIONS.auto || position === LABEL_POSITIONS.outsideEnd)) {
            for (const bar of bars) {
                if (!bar.labelText || (bar.value >= 0) === invert) continue;
                const font = labelFontOf(bar);
                const lines = [bar.labelText, ...(bar.rateText ? [bar.rateText] : [])];
                const [low, high] = extentOf(bar);
                const length = verticalText
                    ? LABEL_EDGE.textGap + maxOf(lines.map((l) => measureTextWidth(l, specOf(font)))) + labelInkPad + OUTSIDE_LABEL_PAD
                    : attachedBaseline(0, -1, font.size, (lines.length - 1) * (font.size + 2), labelInkPad).reach +
                      labelInkPad +
                      OUTSIDE_LABEL_PAD +
                      outsideTopHeadroomOf(font.size);
                topNeeds.push({ end: bar.value >= 0 ? high : low, length });
            }
        }
        if (viewModel.target?.label) topNeeds.push({ end: viewModel.target.value, length: referenceLabelRoom });
    }
    /** 画面の上の側の値の軸の端を延ばせるか（縦向きで、自動の端で、軸を切っていない側） */
    const extendable = !horizontal && (invert ? axis.autoEnds.min && axis.cut !== "bottom" : axis.autoEnds.max && axis.cut !== "top");
    /**
     * 前の版（1.15.8.25 まで）がラベルの分としていつも空けていた高さ：横書きは、棒の先から（率があれば 2 行目まで）字の四角の遠い側まで
     * ＋ 4 ＋ 0.05 文字、縦書きは文字の端までのすき間 ＋ 文字の幅（図の高さの 4 分の 1 まで）＋ 2、目標の名前は文字の大きさ ＋ 6。
     * 延ばせない端（書式で決めた端・軸を切った側）で topNeeds の分だけ描く範囲の上の端を下げるとき、下げてよい一番下の位置
     * （labelFloor：前の版の描く範囲の上の端）を出すのに使う。延ばせる端は 0
     */
    const labelDrop = (() => {
        if (horizontal || extendable || !topNeeds.length || !(axis.max > axis.min)) return 0;
        const labelSize = Math.max(labelFont.size, totalLabelFont.size);
        const rateLine = bars.some((b) => b.rateText) ? labelSize + 2 : 0;
        const labelRoom = !style.dataLabels.show
            ? 0
            : verticalText
              ? Math.min(valueLabelWidth, H * 0.25) + LABEL_EDGE.textGap + 2
              : attachedBaseline(0, -1, labelSize, rateLine, labelInkPad).reach + 4 + 0.05 * labelSize;
        return Math.max(labelRoom, referenceLabelRoom);
    })();

    // 項目の方向の並び（標準と同じ：外側のパディングとカテゴリ間のスペース）
    const pad = style.columns.categorySpacing / 100;
    const outer = style.columns.outerPadding === null ? pad / 2 : style.columns.outerPadding / 100;
    const minWidth = style.categoryAxis.minCategoryWidth;

    let plot: Rect;
    let rotated = false;
    let step = 1;
    let categoryTexts: Array<{ lines: string[]; title: string }>;
    let scrollAxis: "x" | "y" | null = null;
    let content: number;
    /** スクロール する中身の先頭の余白。斜めの項目名の左側が、中身の左端で切れないように空ける */
    let leadIn = 0;
    /** 縦向きの項目名の高さ（描く範囲の下端＋名前の側の余白から、名前の下端まで） */
    let categoryHeight = 4;
    /** 縦向きの Y 軸の上の単位（目盛りの列）に要る、目盛りを引く範囲の上の端の高さ。出さない・右上に置くなら 0 */
    let columnBadgeNeed = 0;
    /** 縦向きで、ドリルの位置を描く範囲の上の行に出すか */
    let drillRowShown = false;
    /** 縦向きの上の行（単位のラベル・ドリルの位置）の下の線。行の字のベースラインはその badgeRise 上。描く範囲の上の端はここか、それより下 */
    let rowLine = 0;
    /** 縦向きで、延ばせない端のラベルのために描く範囲の上の端を下げてよい一番下の位置（前の版の描く範囲の上の端。下げないなら rowLine） */
    let labelFloor = 0;
    /**
     * 縦向きの単位のラベルのベースラインと、その下の基準の高さ（上の行の下の線か、目盛りを引く範囲の上の端）のあいだ。
     * Y 軸の上の単位は、一番上の目盛りの数字が大きければ TOP_ROW_RISE より高く置く（columnBadgeRiseOf）。
     * ドリルの位置は、同じ行に並べる単位と同じベースラインに置くので、上の行の字はどれもこの高さ
     */
    let badgeRise = TOP_ROW_RISE;

    if (!horizontal) {
        // 描く範囲の上の端と、その上の行（単位のラベル・ドリルの位置）の決まりは標準の集合縦棒と同じ（shared/plotTop）。reserve.top から測る：
        // 一番上の目盛りの数字が大きければ上の端を下げる。「プロット エリアの右上」の単位は描く範囲の上の端の上に、Y 軸の上の単位は
        // 目盛りを引く範囲の上の端の上に置き、足りない分だけ上の端を下げる（下の「値 → 画面の座標」）。ドリルの位置はその行に並べ、
        // 字が大きければ行を広げる（足す高さが図の DRILL_ROW_SHARE_MAX を超えるなら出さない。並べるときは Y 軸の上の単位もこの行に置き、
        // 単位を上げるならドリルの位置も同じベースラインまで上げる）。
        // 棒の上の外に置くデータ ラベルは、上の端を下げずに値の軸の範囲を延ばして入れる（domainFor）。延ばせない端では、
        // labelFloor まで上の端を下げて入れる。下（項目名の側）は、ラベルがはみ出すときだけ nearRoom で空ける
        const base = reserve.top;
        const tickPx = style.valueAxis.show ? tickFont.size : 0;
        const plainTop = base + plainTopOf(tickPx);
        // 「プロット エリアの右上」の単位も、長くて目盛りの列（目盛りの数字は描く範囲から 6px 外）の上に掛かるなら、一番上の目盛りの数字と
        // 重ならないように、Y 軸の上の単位と同じ高さに上げる。単位の置き方は下の badge と同じ（描く範囲の右の端にそろえ、左の端で止める）。
        // 前は TOP_ROW_RISE のままで、目盛りの数字が大きいと単位の下が一番上の目盛りの数字に重なった
        const cornerOverColumn = (() => {
            if (!badgeShown || !badgeTopRight || !style.valueAxis.show) return false;
            const side = valueSideOf(tickLabelWidth, true);
            const plotLeft = reserve.left + (switchAxis ? 8 : side);
            const plotRight = W - reserve.right - (switchAxis ? side : 8);
            const badgeWidth = unitWidthOf(badgeText);
            const right = Math.max(Math.min(plotRight, W - reserve.right - 2), reserve.left + 2 + badgeWidth);
            return switchAxis ? right > plotRight + 6 : right - badgeWidth < plotLeft - 6;
        })();
        if (badgeShown && (!badgeTopRight || cornerOverColumn)) badgeRise = columnBadgeRiseOf(tickPx, badgeFont.size);
        const badgeNeed = badgeShown ? base + topRowNeedOf(badgeFont.size, UNIT_TEXT_ASCENT, badgeRise) : 0;
        const rowTop = Math.max(plainTop, badgeTopRight ? badgeNeed : 0);
        columnBadgeNeed = badgeTopRight ? 0 : badgeNeed;
        const drillTop = Math.ceil(Math.max(rowTop, columnBadgeNeed, base + topRowNeedOf(drillPathFont.size, DRILL_TEXT_ASCENT, badgeRise)));
        drillRowShown = Boolean(viewModel.drillPath) && drillTop - Math.max(rowTop, columnBadgeNeed) <= (H - base - reserve.bottom) * DRILL_ROW_SHARE_MAX;
        const top = drillRowShown ? drillTop : rowTop;
        rowLine = top;
        if (labelDrop > 0) {
            // 前の版（1.15.8.25 まで）の描く範囲の上の端：凡例の下の端（凡例が無ければビューの上の端から TOP_MARGIN）に、
            // ドリルの位置の行（単位を出さないとき）と単位の行（ドリルの位置を並べるなら大きいほうの文字の行）を積み、labelDrop を足した所。
            // 下げるのはここまで。上の行の下の線から labelDrop を測ると、上の行が前の版の行より下の形（上の凡例の下の 10px、
            // ドリルの位置・右上の単位の行）では前の版より描く範囲が低くなり、上の行が前の版より上の形（Y 軸の上の単位を
            // 目盛りの列に移した形）ではラベルの入る高さが前の版より狭くなる。前の版の上の端までにすると、どちらの形も前の版と同じ所まで
            // 下げる。描く範囲の高さ・斜めの名前の長さ・棒の段差は前の版と違うので、入るラベルは前の版と入れ替わることがある。
            // 凡例の下の端は、いまの凡例の箱の下の端（base）。上の凡例の箱は 1.15.8.25 より低い（1.15.8.29 から、行の高さが文字の大きさ ＋ 8px の
            // 四捨五入で、ビューの上の端の 8px を取らない。10pt で 31.3 → 29px）ので、上に凡例を置いた形の下限は 1.15.8.25 の上の端より
            // 上（1 行で 2px ほど）になる
            const shares = Boolean(viewModel.drillPath) && badgeShown;
            const legacyDrillRow = viewModel.drillPath && !shares ? Math.ceil(drillPathFont.size * 1.3) + 6 : 0;
            const legacyRowFont = shares ? Math.max(badgeFont.size, drillPathFont.size) : badgeFont.size;
            const legacyBadgeRow = !badgeShown ? 0 : shares ? legacyRowFont + Math.max(6, Math.ceil(legacyRowFont * 0.22) + 2) : badgeFont.size + 6;
            const legacyTop = Math.max(base, TOP_MARGIN) + legacyDrillRow + legacyBadgeRow + labelDrop;
            labelFloor = Math.max(top, legacyTop);
        } else {
            labelFloor = top;
        }
        /** 上の端は、延ばせない端のラベル・Y 軸の上の単位に合わせて下げてもここまで（斜めの項目名の長さの上限に使う） */
        const lowestTop = Math.max(labelFloor, drillRowShown ? top : columnBadgeNeed);
        // 目盛りの列は、Y 軸の上の単位が目盛りの数字より長ければ、単位の幅まで広げる（上の badgeText。広げる幅に上限があり、入らない単位は切ってある）
        const columnBadgeWidth = badgeShown && columnBadge ? unitWidthOf(badgeText) : 0;
        const valueSide = valueSideOf(Math.max(tickLabelWidth, columnBadgeWidth), style.valueAxis.show || columnBadgeWidth > 0);
        const left = reserve.left + (switchAxis ? 8 : valueSide);
        const right = reserve.right + (switchAxis ? valueSide : 8);
        const width = W - left - right;
        const needed = minWidth > 0 ? n * minWidth : 0;
        if (needed > width + 1) scrollAxis = "x";
        content = scrollAxis ? needed : width;
        const categoryStep = content / (n - pad + 2 * outer);
        const textBand = categoryStep - 4;
        // 項目名：帯に入れば水平（2 行目に上の階層）、入らなければ斜めにして上の階層から出す
        const flatWidth = Math.max(0, maxOf(bars.map((b) => measureTextWidth(b.label, specOf(categoryFontOf(b))))));
        if (!style.categoryAxis.show) {
            categoryTexts = bars.map((b) => ({ lines: [] as string[], title: b.fullLabel }));
            // 項目名を出さずに X 軸のタイトルだけ出すときは、描く範囲の下の端からタイトルの行の上の端まで AXIS_TEXT_BELOW（横向きと同じ。前は 4px）
            if (categoryTitle) categoryHeight = AXIS_TEXT_BELOW;
        } else if (flatWidth <= textBand) {
            // 1 行目のベースラインは、描く範囲の下の端（名前の側の余白の下）から字の大きさ ＋ AXIS_TEXT_DROP（標準の集合縦棒と同じ。
            // shared/axisBands）。最後の行のベースラインの下は、X 軸のタイトルがあれば AXIS_TEXT_BELOW、無ければ下の端まで COLUMN_TEXT_BOTTOM
            // （標準の集合縦棒と同じ。1.15.8.39 までは 4px で、下の帯が標準より 5px ほど低かった）。下へ下りる字・コンマ・下線の下の端がその下へ
            // 出るなら、字の下の端がビューに入るまでとる（axisTextBottomOf。最後の行（2 行なら上の階層の行）に出す字で見積もる。1.15.8.40 の最初の作りでは
            // 字によらず COLUMN_TEXT_BOTTOM で、大きい字の g・p・y が下の端で切れた）。スクロール バーが下にあるときは前のまま 4px。
            // 1.15.8.37 までは、1 行目のベースラインを字の大きさ ＋ 5px に置いた
            const lastRowInk = Math.max(
                0,
                maxOf(bars.map((b) => {
                    const text = hasParent ? b.parentLabel : b.label;
                    const font = categoryFontOf(b);
                    return text ? axisTextDescentOf(text, font.decoration === "underline", { fontPx: font.size, bold: font.weight === "bold" }) * font.size : 0;
                }))
            );
            categoryHeight =
                (hasParent ? 1 : 0) * categoryLineHeight +
                categorySize +
                AXIS_TEXT_DROP +
                (categoryTitle ? AXIS_TEXT_BELOW : scrollAxis ? 4 : axisTextBottomOf(COLUMN_TEXT_BOTTOM, lastRowInk));
            categoryTexts = bars.map((b) => ({
                lines: [
                    truncateToWidth(b.label, textBand, specOf(categoryFontOf(b))),
                    ...(b.parentLabel ? [truncateToWidth(b.parentLabel, textBand, specOf(categoryFontOf(b)))] : []),
                ],
                title: b.fullLabel,
            }));
        } else {
            rotated = true;
            step = Math.max(1, Math.ceil((categorySize * 1.45) / categoryStep));
            // 上限を大きくしても、描く範囲に高さの 4 分の 1 は残す（1.11 までは 100% で描く範囲が無くなり、何も描かなかった）
            const room = H * 0.75 - lowestTop - nearRoom - ROTATED_GAP - categorySize * 0.5 - categoryTitleSize - reserve.bottom - noticeHeight - (scrollAxis ? SCROLLBAR : 0);
            const maxLength = Math.max(0, Math.min(H * style.categoryAxis.maxShare, room)) / SIN45;
            const longest = Math.max(0, maxOf(bars.map((b) => measureTextWidth(pathOf(b), specOf(categoryFontOf(b))))));
            const length = Math.min(longest, maxLength);
            // 斜めの名前は (中心, 下端 + 8) を右端に -45° 回す。縦の広がりは長さ × sin45 と、文字の下側の少し
            // （見積もり：棒との間隔 8 ＋ 長さ × sin45 ＋ 文字の半分）
            categoryHeight = ROTATED_GAP + length * SIN45 + categorySize * 0.5;
            // 名前の長さの上限。欄がいちばん長い名前ちょうどのとき、計算の誤差で省かないよう 0.5px の余裕を持たせる。
            // 入らない名前は末尾を省き、1 文字も入らなければ頭の 1 文字と「…」を残す（元の名前のほうが狭ければ元の名前）。
            // それが上限 ＋ 0.5px を超えるときだけ、その名前を出さない（左の端で省く下の扱いと同じ。1.15.8.42 までは、
            // 1 文字も入らない名前を出さず、上限と上限 ＋ 0.5px のあいだの短い名前もすべて消えた）
            const lengthCap = length + 0.5;
            /** 幅 limit まで末尾を省いた名前。頭の 1 文字と「…」が上限 ＋ 0.5px に入らなければ空 */
            const capName = (path: string, limit: number, font: FontSpec): string => {
                const cut = truncateToWidth(path, limit, font) || headWithEllipsis(path, font);
                return measureTextWidth(cut, font) > lengthCap + 0.5 ? "" : cut;
            };
            categoryTexts = bars.map((b) => ({
                lines: [capName(pathOf(b), lengthCap, specOf(categoryFontOf(b)))],
                title: b.fullLabel,
            }));
            if (scrollAxis) {
                // 斜めの名前は棒の中心から左下へ伸びる。スクロールすると中身の左端より外は見えないので、はみ出す分だけ前を空ける
                leadIn = Math.max(
                    0,
                    maxOf(categoryTexts.map((t, k) =>
                        k % step !== 0 || !t.lines.length
                            ? 0
                            : measureTextWidth(t.lines[0], specOf(categoryFontOf(bars[k]))) * SIN45 +
                              categorySize * 0.5 -
                              categoryStep * (outer + k + (1 - pad) / 2)
                    ))
                );
                content += leadIn;
            } else {
                // スクロールしないときも、斜めの名前はビジュアルの左端より外へ出ると切れる（項目名が長いと、1 本目の名前の頭が隠れた）。
                // 名前の左端が左端の余白に収まるだけ、棒の並びの前を空ける。空けると帯が狭まり中心も動くので、
                // 中心 = left + leadIn + (width − leadIn) × kk として解く（1.15.8.43 から）：
                // - 空けるのは、名前が帯の 0.92 倍より長く、帯が文字の 1.45 倍以上のときだけ（狭い帯は名前を間引くので空けない）
                // - 名前の広がりは「長さ（長さの上限で頭を打つ）× sin45」（文字の半分は足さない）。余白は rotatedLeftMarginOf ＋ 1px
                // - 空けすぎないよう、幅の 30% と、帯がカテゴリの最小幅と文字の 1.45 倍を下回らない所までで止める
                // 左の端は、左に凡例を置けば凡例の分（凡例の箱と、そのあとの余白）の右の端。
                // Y 軸を右に移すと描く範囲が凡例のすぐ右に来て、1 本目の名前が凡例の名前に掛かった。
                // 1.15.8.42 までは、空ける条件を見ず、広がりに文字の半分を足し、余白はビューの左端か凡例の箱の右の端から 4px、上限は幅の 30% だけだった
                const leftEdge = reserve.left;
                const leftMargin = rotatedLeftMarginOf(categorySize);
                const slots = n - pad + 2 * outer;
                const centerRatio = (k: number) => (outer + k + (1 - pad) / 2) / slots;
                const fullWidths = bars.map((b) => measureTextWidth(pathOf(b), specOf(categoryFontOf(b))));
                if (maxOf(fullWidths) > categoryStep * 0.92 && categoryStep >= categorySize * 1.45) {
                    const need = rotatedLabelLeadIn(
                        fullWidths.map((w, k) => ({ extent: Math.min(w, length) * SIN45, ratio: centerRatio(k) })),
                        left - leftEdge,
                        width,
                        leftMargin + 1
                    );
                    const keepStep = Math.max(minWidth, categorySize * 1.45);
                    leadIn = Math.max(0, Math.min(need, width * 0.3, width - keepStep * slots));
                }
                // 名前ごとに、名前の左端が余白に収まる長さ（12px を下回らない）まで末尾を省く。省いても頭の 1 文字と「…」は
                // 残し（元の名前のほうが狭ければ元の名前）、それが長さの上限に入らなければ、その名前だけ出さない。字の左の端（中心 − 長さ × sin45 −
                // 文字の半分）が左の端より外へ出る名前も出さない（切れた字を出さない。余白が文字の半分以上あるので、余白まで省いた名前は外へ
                // 出ない。出るのは、12px の下限か、頭の 1 文字と「…」が左の端の分の長さより長い名前）。
                // 1.15.8.42 までは頭を省き、1 文字も入らなければ出さなかった
                categoryTexts = categoryTexts.map((t, k) => {
                    if (!t.lines.length) return t;
                    const center = left + leadIn + (width - leadIn) * centerRatio(k);
                    const room = Math.max(12, (center - leftEdge - leftMargin) / SIN45);
                    if (room >= lengthCap) return t;
                    const cut = capName(pathOf(bars[k]), room, specOf(categoryFontOf(bars[k])));
                    const outside = center - measureTextWidth(cut, specOf(categoryFontOf(bars[k]))) * SIN45 - categorySize * 0.5 < leftEdge;
                    return { ...t, lines: [outside ? "" : cut] };
                });
            }
        }
        const bottomReserve =
            nearRoom + categoryHeight + categoryTitleSize + reserve.bottom + noticeHeight + (scrollAxis ? SCROLLBAR : 0);
        plot = { x: left, y: top, width, height: H - top - bottomReserve };
    } else {
        const tickHeight = style.valueAxis.show ? tickFont.size + 8 : 4;
        // 値の軸を下に置くときの行は標準の集合横棒と同じ（shared/axisBands）：目盛りの数字のベースラインを描く範囲の下の端から
        // 字の大きさ ＋ AXIS_TEXT_DROP、タイトルのベースラインをその AXIS_TEXT_BELOW ＋ 文字の大きさ × AXIS_TITLE_ASCENT 下に置く
        // （タイトルのベースラインから下の端までは BAR_TITLE_BOTTOM）。タイトルも下の単位の行も無ければ、目盛りの数字のベースラインから
        // 下の端までを BAR_TITLE_BOTTOM と同じく標準の集合横棒にそろえて BAR_TEXT_BOTTOM（1.15.8.39 までは 4px で、下の帯が標準より 5px ほど低かった）。
        // コンマ・下へ下りる字・下線の下の端がその下へ出るなら、字の下の端がビューに入るまでとる（axisTextBottomOf。目盛りの数字は描く範囲の長さで
        // 決まるので、Y 軸の幅の見積もりの tickLabelWidth と同じく、範囲の目盛りといちばん細かい 8 本の目盛りの字で見積もる）。
        // 下に単位の行があれば、目盛りの数字との間は前のまま 4px。
        // 1.15.8.37 までは、目盛りのベースラインを字の大きさ ＋ 4px、タイトルの行を文字の大きさ ＋ 6px にとった。上に置くときは前のまま
        // 目盛りの数字を出さずにタイトルだけ出すときは、描く範囲の下の端からタイトルの行の上の端まで AXIS_TEXT_BELOW（縦向きと同じ。前は 4px）
        const badgeBelow = !(switchAxis || badgeTopRight) && badgeHeight > 0;
        const tickRowBelow = style.valueAxis.show
            ? tickFont.size +
              AXIS_TEXT_DROP +
              (valueTitleFont
                  ? AXIS_TEXT_BELOW
                  : badgeBelow
                    ? 4
                    : axisTextBottomOf(
                          BAR_TEXT_BOTTOM,
                          Math.max(
                              0,
                              maxOf([...axis.ticks, ...(axis.ticksFor?.(Math.max(8, style.valueAxis.tickCount)) ?? [])].map((t) =>
                                  axisTextDescentOf(t.label, tickFont.decoration === "underline", { fontPx: tickFont.size, bold: tickFont.weight === "bold" })
                              ))
                          ) * tickFont.size
                      ))
            : valueTitleFont
              ? AXIS_TEXT_BELOW
              : 4;
        // タイトルのベースラインの下は、下へ下りる字・下線の下の端が入るだけとる（bottomTitleBelowOf。既定の組の字は BAR_TITLE_BOTTOM のまま）
        const titleRowBelow = valueTitleFont ? valueTitleFont.size * AXIS_TITLE_ASCENT + bottomTitleBelowOf(BAR_TITLE_BOTTOM, valueTitleFont.size, valueTitleInk) : 0;
        // 単位のラベルは値の軸の側（下なら目盛りとタイトルの下、上ならその上）に 1 行とる。右上に置くときは上
        const badgeAbove = switchAxis || badgeTopRight;
        const top = reserve.top + (badgeAbove ? badgeHeight : 0) + referenceLabelRoom + 4 + (switchAxis ? tickHeight + valueTitleSize : 0);
        const bottom = reserve.bottom + noticeHeight + (switchAxis ? 4 : tickRowBelow + titleRowBelow) + (badgeAbove ? 0 : badgeHeight);
        const height0 = H - top - bottom;
        const needed = minWidth > 0 ? n * minWidth : 0;
        if (needed > height0 + 1) scrollAxis = "y";
        content = scrollAxis ? needed : height0;
        const band0 = content / (n - pad + 2 * outer);
        const twoLines = hasParent && band0 >= 2 * categoryLineHeight;
        const textOf = (b: Bar) => (twoLines ? b.label : pathOf(b));
        const widest = Math.max(
            0,
            maxOf(bars.map((b) =>
                Math.max(
                    measureTextWidth(textOf(b), specOf(categoryFontOf(b))),
                    twoLines ? measureTextWidth(b.parentLabel, specOf(categoryFontOf(b))) : 0
                )
            ))
        );
        // 棒の外のすき間（LABEL_EDGE の文字の端まで 7.5px）に合わせる。字の送りの右の端からビューの右の端までは H_PLOT_RIGHT
        // （1.15.8.30 までは、ラベルなしで 8px、ラベルの字の右の端から 6px）。背景を描いても字の位置は変えず、背景の四角は右の余白の中へ
        // H_LABEL_BACKGROUND_PAD 広げる（標準の集合横棒と同じ。1.15.8.37 までは、背景の分だけ字を内側へ寄せた）
        const labelRoom = style.dataLabels.show ? Math.min(valueLabelWidth + LABEL_EDGE.textGap + H_PLOT_RIGHT, W * 0.25) : H_PLOT_RIGHT;
        // 上限を大きくしても、描く範囲に幅の 4 分の 1 は残す（縦向きと同じ。データ ラベルの分の右の余白は下の labelMargin で、その残りからとる）。右は H_PLOT_RIGHT とスクロール バーの分だけ除き、
        // データ ラベルの分は除かない（1.15.8.42 までは除いたので、ラベルを出して「幅の最大値」を大きくすると、名前の欄がその分狭かった）
        const room = W * 0.75 - reserve.left - sideTitleSize - nearRoom - reserve.right - H_PLOT_RIGHT - (scrollAxis ? SCROLLBAR : 0);
        // 名前の字の幅の上限は、ビューの幅 × 「最大幅 (%)」（16px を下回らない。名前の左右のすき間 namePads は上限の外）。
        // 名前の欄に使える幅（room から namePads を除いた分）でも頭を打つ。省いても頭の 1 文字と「…」は残し（元の名前のほうが狭ければ元の名前）、
        // それが幅に入らなければ、その名前だけ出さない。どの名前も出さなければ、名前の欄をとらない（1.15.8.43 から。
        // 1.15.8.42 までは、namePads を含めた欄の幅をビューの幅 × 割合で抑え、1 文字も入らない名前は空にした）
        const nameCap = Math.max(16, W * style.categoryAxis.maxShare);
        const textWidth = Math.min(nameCap, Math.max(0, room - namePads));
        /** 幅 textWidth に入るように末尾を省いた名前。頭の 1 文字と「…」も入らなければ空 */
        const fitName = (text: string, font: FontSpec): string => {
            if (!text) return "";
            const cut = truncateToWidth(text, textWidth, font) || headWithEllipsis(text, font);
            return measureTextWidth(cut, font) <= textWidth + 0.5 ? cut : "";
        };
        categoryTexts = bars.map((b) => ({
            lines: !style.categoryAxis.show
                ? []
                : twoLines
                  ? [fitName(b.label, specOf(categoryFontOf(b))), ...(b.parentLabel ? [fitName(b.parentLabel, specOf(categoryFontOf(b)))] : [])]
                  : [fitName(textOf(b), specOf(categoryFontOf(b)))],
            title: b.fullLabel,
        }));
        const namesShown = categoryTexts.some((t) => t.lines.some((line) => line));
        const categoryWidth = namesShown ? Math.min(widest, textWidth) + namePads : 0;
        // 名前の欄を出さなければ、描く範囲をタイトルの字の右の端から SIDE_TITLE_PLOT_GAP に、タイトルも無ければ SIDE_BARE_LEFT（4px）に置く
        // （shared/axisBands の sidePlotLeftOf）。1.15.8.37 までは、名前の欄の代わりにいつも 8px とった。1.15.8.39 までは、タイトルの字の右の端から
        // SIDE_TITLE_NAME_GAP（標準より 2px 狭い）だった
        const left = reserve.left + sidePlotLeftOf(categoryTitleFont ? categoryTitleFont.size : null, categoryWidth, categoryTitleInk) + nearRoom;
        // 目盛りの数字は線の真ん中に置くので、右の余白（凡例の分とスクロール バーを除く）は、広げたあとの幅で出す目盛り（下の
        // shownTicks と同じ選び方）の数字の右の端がビューに入るように決める（広げて目盛りが変われば見直す。shared の endMarginOf）。
        // 目盛りの行はスクロールする窓の外にあるので、数字はスクロール バーの列の下まで使える。広げても描く範囲に幅の 4 分の 1 は残す
        const scrollbar = scrollAxis ? SCROLLBAR : 0;
        const toRight = W - left - reserve.right;
        const widthAt = (margin: number) => toRight - margin - scrollbar;
        const tickNeedAt = (margin: number): number => {
            const w = widthAt(margin);
            const valueLength = w - lowZone - highZone;
            if (!style.valueAxis.show || !(axis.max > axis.min) || !(valueLength > 0)) return -Infinity;
            /** 描く範囲の幅 at のときの値 v の x（下の pos と同じ写し方） */
            const xAt = (v: number, at: number) => {
                const offset = lowZone + ((v - axis.min) / (axis.max - axis.min)) * (at - lowZone - highZone);
                return left + (invert ? at - offset : offset);
            };
            const shown = ticksShownAt(w, (a, b) => Math.abs(xAt(a, w) - xAt(b, w)));
            // 数字の右の端（x ＋ 幅の半分）は幅の 1 次式。ビューの右の端（凡例の手前）に入る幅の上限から、要る余白を出す
            return maxOf(
                shown.map((t) => {
                    const slope = xAt(t.value, 1) - xAt(t.value, 0);
                    if (!(slope > 1e-9)) return -Infinity;
                    const most = (toRight + left - measureTextWidth(t.label, specOf(tickFont)) / 2 - xAt(t.value, 0)) / slope;
                    return toRight - scrollbar - most;
                })
            );
        };
        // データ ラベルの分の右の余白は、描く範囲に幅の 4 分の 1 が残るところまで（名前の欄に使える幅からラベルの分を除かないので、
        // 名前の欄が広いと、ラベルの分をとると描く範囲が右の余白の 10px まで細くなり、MIN_PLOT を割って何も描かない形があった。
        // 入らないラベルは右の余白の外へ出る）
        const labelMargin = Math.min(labelRoom, Math.max(H_PLOT_RIGHT, toRight - scrollbar - W / 4));
        const margin = endMarginOf(labelMargin, toRight - scrollbar - W / 4, tickNeedAt);
        plot = { x: left, y: top, width: widthAt(margin), height: height0 };
        step = Math.max(1, Math.ceil(categoryLineHeight / band0));
    }
    // --- 値 → 画面の座標 --------------------------------------------------------------------------
    // 軸を切った印の余白（lowZone・highZone）は上で決めてある
    /**
     * 値を並べる長さ valueLength (px) のときの、値を並べる範囲。目盛りを引く範囲（axis.min〜axis.max）の画面の上の側の端を、
     * topNeeds が描く範囲に入るまで延ばす（標準の集合縦棒と同じく、描く範囲の上の端は下げない）。延ばすのは extendable のときだけ。
     * 延ばす分は延ばしたあとの範囲の LABEL_ROOM_SHARE_MAX まで（棒が潰れないように）。延ばした端は丸めない（目盛りは目盛りを引く範囲の中に引く）
     */
    const domainFor = (valueLength: number): [number, number] => {
        const { min, max } = axis;
        if (!extendable || !topNeeds.length || !(valueLength > 0) || !(max > min)) return [min, max];
        const most = (max - min) / (1 - LABEL_ROOM_SHARE_MAX);
        // ラベル（値 end、長さの割合 share）が入るのは、延ばす端 e について end + share × (e − 反対の端) が e を越えないとき
        if (!invert) {
            let high = max;
            for (const n of topNeeds) {
                const share = n.length / valueLength;
                high = Math.max(high, share < 1 ? (n.end - share * min) / (1 - share) : Infinity);
            }
            return [min, Math.min(high, min + most)];
        }
        let low = min;
        for (const n of topNeeds) {
            const share = n.length / valueLength;
            low = Math.min(low, share < 1 ? (n.end - share * max) / (1 - share) : -Infinity);
        }
        return [Math.max(low, max - most), max];
    };
    if (labelDrop > 0) {
        // 延ばせない端（書式で決めた端・軸を切った側）では、上の行の下にラベルの分を空けて描く範囲の上の端を下げる。
        // ラベルを付ける値の位置の、描く範囲の中の割合 r は上の端を下げても変わらないので、どのラベルの外の端も上の行の下の線より下になる
        // 一番上の位置を解く。軸の外の値（切った側の外）の棒は、切った印の余白を越えて描く範囲の上の端まで伸びる（pos）ので、
        // ラベルの長さをそのまま空ける。下げるのは labelFloor（前の版の描く範囲の上の端）まで
        const bottom = plot.y + plot.height;
        const zoneTop = invert ? lowZone : highZone;
        const below = bottom - rowLine - lowZone - highZone;
        let top = plot.y;
        for (const n of topNeeds) {
            if (invert ? n.end < axis.min : n.end > axis.max) {
                top = Math.max(top, rowLine + n.length);
                continue;
            }
            const r = Math.max(0, Math.min(1, (invert ? n.end - axis.min : axis.max - n.end) / (axis.max - axis.min)));
            if (r < 1) top = Math.max(top, rowLine + (n.length - zoneTop - r * below) / (1 - r));
        }
        top = Math.min(top, labelFloor);
        // 組み直し（keepPlot）で横にスクロールするなら、描く範囲の高さを MIN_PLOT 残す所までしか下げない（上げるのはスクロール バーの分まで）
        if (keepPlot && scrollAxis) top = Math.min(top, Math.max(labelFloor - SCROLLBAR, bottom - MIN_PLOT));
        if (top > plot.y) plot = { ...plot, y: top, height: bottom - top };
    }
    if (!horizontal && columnBadgeNeed > 0 && !drillRowShown) {
        // Y 軸の上の単位は、目盛りを引く範囲の上の端（ラベルの分だけ範囲を延ばすと、描く範囲の上の端より下）の badgeRise 上に置き、
        // 要る高さに届かなければ描く範囲の上の端を下げる。下げると値を並べる長さが縮み、延ばす分も変わるので、届く一番上の位置を二分法で探す
        // （上の端を下げるほど、目盛りを引く範囲の上の端も下がる）
        const bottom = plot.y + plot.height;
        const frameTopAt = (top: number): number => {
            const length = bottom - top - lowZone - highZone;
            const [low, high] = domainFor(length);
            const beyond = invert ? axis.min - low : high - axis.max;
            return top + (invert ? lowZone : highZone) + (high > low ? (beyond / (high - low)) * Math.max(0, length) : 0);
        };
        if (frameTopAt(plot.y) < columnBadgeNeed) {
            let reach = columnBadgeNeed;
            let short = plot.y;
            for (let k = 0; k < 40; k++) {
                const mid = (reach + short) / 2;
                if (frameTopAt(mid) >= columnBadgeNeed) reach = mid;
                else short = mid;
            }
            plot = { ...plot, y: reach, height: bottom - reach };
        }
    }
    if (plot.width < MIN_PLOT || plot.height < MIN_PLOT) return null;
    const whole = horizontal ? plot.width : plot.height;
    const valueLength = whole - lowZone - highZone;
    // 軸を切った印の余白を引くと値を並べる長さが残らないなら、図を描かない（描く範囲が MIN_PLOT を切るときと同じ）。
    // 長さが負だと、目盛りの線・棒が描く範囲の外に出る（1.15.8.30 までは、描く範囲の高さが 10〜16px の形で線が外に出た）。
    // 印を出さずに描くと、切った軸を切っていない軸と読み違えるので、描かない
    if (valueLength <= 0) return null;
    /** 値を並べる範囲（縦向きは棒の上の外のラベルの分だけ延ばす。目盛りは axis.min〜axis.max の中に引く） */
    const [domainMin, domainMax] = horizontal ? [axis.min, axis.max] : domainFor(valueLength);
    const span = domainMax - domainMin || 1;
    /** 値の小さい端からの距離 → 画面の座標。範囲の反転ではひっくり返す */
    const screen = (offset: number): number => {
        const o = invert ? whole - offset : offset;
        return horizontal ? plot.x + o : plot.y + plot.height - o;
    };
    /** 値 → 値の方向の画面座標。範囲の外は、余白も含めた描く範囲の端で止める（0 が軸の外の合計の棒は端まで伸びる） */
    const pos = (v: number): number =>
        screen(v < domainMin ? 0 : v > domainMax ? whole : lowZone + ((v - domainMin) / span) * valueLength);
    /** 値が増える向きの画面の符号（縦は上が負、横は右が正。反転で逆） */
    const growth = horizontal ? (invert ? -1 : 1) : invert ? 1 : -1;

    // 出す目盛り（選び方は ticksShownAt）。返す tickTarget は、描く範囲の長さで決まる本数の目安（範囲を丸め直すのに使う）
    const tickTarget = tickTargetOf(style.valueAxis.tickCount, horizontal ? plot.width : plot.height, horizontal ? "horizontal" : "vertical");
    const shownTicks = ticksShownAt(horizontal ? plot.width : plot.height, (a, b) => Math.abs(pos(a) - pos(b)));
    if (!horizontal) {
        // 縦向きで、描く目盛りの数字の幅が、この回の列の幅と違えば、その幅で組み直す（落ち着けばこの回のまま描く）
        const nextColumns = nextColumnTry(columnTry, [tickLabelWidth], [style.valueAxis.show ? widestTick(shownTicks) : 0]);
        if (nextColumns) return layoutWith(viewModel, viewport, nearRoom, legendPage, keepPlot, ticksAt, nextColumns);
    }

    // 項目の方向：中身（スクロールするなら中身の長さ）の中で並べる
    const categoryOrigin = (horizontal ? plot.y : plot.x) + leadIn;
    const categoryStepLength = (content - leadIn) / (n - pad + 2 * outer);
    const barThickness = Math.max(1, categoryStepLength * (1 - pad));
    const bandStart = (k: number) => categoryOrigin + categoryStepLength * (outer + k);

    /** 値 a〜b と帯 k を、画面の長方形にする。長さ 0 でも 1px は出す（増減 0 の項目がどこにあるか分かるように） */
    const rectOf = (a: number, b: number, k: number): Rect => {
        let p1 = Math.min(pos(a), pos(b));
        let p2 = Math.max(pos(a), pos(b));
        // 両端とも軸の範囲の同じ側の外なら描かない（1px にすると、見えないはずの棒が端に細い帯で残る）
        const outside = (Math.max(a, b) < axis.min || Math.min(a, b) > axis.max) && a !== b;
        if (outside) p2 = p1;
        else if (p2 - p1 < 1) {
            p1 -= 0.5;
            p2 = p1 + 1;
        }
        return horizontal
            ? { x: p1, y: bandStart(k), width: p2 - p1, height: barThickness }
            : { x: bandStart(k), y: p1, width: barThickness, height: p2 - p1 };
    };

    /**
     * 角を丸める端。棒の端（いちばん上・下）を丸め、値 0 の軸に乗っている端だけ四角にする
     * （四角い端は「軸に乗っている」と読める）。合計は値の端だけ、浮いた増減・小計は両端
     * （2026-09-24 ユーザーと決定。片側だけ丸めると、四角い側が何かに乗っているように見える）。
     * 系列で積むときは、棒の端に届く外側の区画の、その端だけ。軸の範囲の外で切れた端は、そこが本当の端ではないので丸めない
     */
    /** 棒の外側の端（いちばん上・下）で、0 の軸に乗っていない値か */
    const outerEnd = (v: number, low: number, high: number) => style.columns.cornerRadius > 0 && (v === low || v === high) && v !== 0;
    /** 軸の範囲の中か（外なら、描く端は切れた所で本当の端ではない） */
    const inRange = (v: number) => v >= axis.min && v <= axis.max;
    /** 値 a〜b の長方形のうち、a の端が画面のどちらの辺か */
    const sideOfEnd = (a: number, b: number): CornerSide =>
        horizontal ? (pos(a) > pos(b) ? "right" : "left") : pos(a) < pos(b) ? "top" : "bottom";

    const laid: BarLayout[] = bars.map((bar, k) => {
        const [low, high] = extentOf(bar);
        const rect = rectOf(low, high, k);
        const segments = bar.segments.map((segment): SegmentLayout => {
            const { from, to, highlightTo } = segment;
            const segmentRect = rectOf(from, to, k);
            const roundFrom = from !== to && outerEnd(from, low, high) && inRange(from);
            const roundTo = from !== to && outerEnd(to, low, high) && inRange(to);
            // 該当分が 0（highlightTo = from）のハイライトは 1px の帯になる。丸めた端に四角い帯を重ねないよう、そのときは描かない
            const highlight = highlightTo === null || (highlightTo === from && roundFrom) ? null : rectOf(from, highlightTo, k);
            // ハイライトの該当分は、積み始めの端は区画と同じに、該当分の終わりは、区画の外側の端に当たる区画のときに丸める
            // （値の端を丸める）。終わりが 0・軸の範囲の外なら四角
            const roundHighlightEnd =
                highlight !== null && highlightTo !== null && highlightTo !== from && outerEnd(to, low, high) && highlightTo !== 0 && inRange(highlightTo);
            const cornerOf = (sides: CornerSide[]): Corner | null =>
                sides.length ? { sides, radius: style.columns.cornerRadius } : null;
            return {
                segment,
                rect: segmentRect,
                highlight,
                corner: cornerOf([...(roundFrom ? [sideOfEnd(from, to)] : []), ...(roundTo ? [sideOfEnd(to, from)] : [])]),
                highlightCorner:
                    highlight && highlightTo !== null && highlightTo !== from
                        ? cornerOf([
                              ...(roundFrom ? [sideOfEnd(from, highlightTo)] : []),
                              ...(roundHighlightEnd ? [sideOfEnd(highlightTo, from)] : []),
                          ])
                        : null,
            };
        });

        let category: BarLayout["category"] = null;
        const texts = categoryTexts[k];
        if (texts.lines.length && k % step === 0) {
            if (!horizontal) {
                const cx = rect.x + rect.width / 2;
                // 斜めの名前は、文字の高さの中ほど（ベースラインから 0.35 文字）を通る線が棒の中心で終わるように、
                // ベースラインを文字の下側（右下）へずらす。ベースラインの端を中心に置くと、文字の本体が左上に寄って見える
                const lift = categoryFontOf(bar).size * 0.35 * SIN45;
                category = rotated
                    ? {
                          x: cx + lift,
                          y: plot.y + plot.height + nearRoom + ROTATED_GAP + lift,
                          anchor: "end",
                          lines: texts.lines,
                          rotate: -45,
                          font: categoryFontOf(bar),
                          lineHeight: categoryLineHeight,
                          title: texts.title,
                      }
                    : {
                          x: cx,
                          y: plot.y + plot.height + nearRoom + categorySize + AXIS_TEXT_DROP,
                          anchor: "middle",
                          lines: texts.lines,
                          rotate: 0,
                          font: categoryFontOf(bar),
                          lineHeight: categoryLineHeight,
                          title: texts.title,
                      };
            } else {
                const cy = rect.y + rect.height / 2;
                const firstY = cy + categoryFontOf(bar).size * 0.35 - ((texts.lines.length - 1) * categoryLineHeight) / 2;
                category = {
                    x: plot.x - nearRoom - SIDE_NAME_PLOT_GAP,
                    y: firstY,
                    anchor: "end",
                    lines: texts.lines,
                    rotate: 0,
                    font: categoryFontOf(bar),
                    lineHeight: categoryLineHeight,
                    title: texts.title,
                };
            }
        }
        // ラベルは、ほかの物（軸・凡例・線）をそろえてから最後に置く（下の「ラベルの置き方」）
        return { bar, rect, segments, label: null as TextLayout | null, category };
    });

    // --- 軸を切った印（合計の棒の根元） ------------------------------------------------
    // 一番下の目盛り（軸の端）の外に余白（BREAK_ZONE）をとり、印はその真ん中に置く。
    // 目盛りの線のすぐ上に印があると、その目盛りのすぐ上で切ったように見えるため（実際は 0 と端の目盛りのあいだを省いている）
    const breaks: BreakMark[] = [];
    if (marked && axis.cut !== "none") {
        const center = screen(axis.cut === "bottom" ? BREAK_ZONE / 2 : whole - BREAK_ZONE / 2);
        for (const b of laid) {
            if (b.bar.kind !== BAR_KINDS.total) continue;
            // 根元が余白まで届いている（0 が軸の外にある）合計の棒だけ
            const [start, end] = horizontal ? [b.rect.x, b.rect.x + b.rect.width] : [b.rect.y, b.rect.y + b.rect.height];
            if (center < start + 0.5 || center > end - 0.5) continue;
            const [u0, u1] = horizontal ? [b.rect.y, b.rect.y + b.rect.height] : [b.rect.x, b.rect.x + b.rect.width];
            breaks.push(breakMark(style.breakStyle, horizontal, center, u0 - 1, u1 + 1));
        }
    }

    // --- 接続線（棒の終わりの累計の高さで、前の棒の左端から次の棒の右端まで） ---------------------
    // 標準のウォーターフォールと同じ引き方（2026-09-19 に撮った第 26 ページ ① で、どの線も棒 2 本ぶんの幅にまたがっていた）。
    // 棒の上（下）の辺にも線が重なる。横向きは、前の棒の上端から次の棒の下端まで。
    // 小計（浮いた棒）は累計を進めないので、前後の線は同じ高さになる。同じ高さの線は 1 本にまとめる
    // （二重に引くと、点線・破線の模様がずれて濃く見えるため）
    // 棒の端を丸めていれば、線はその端の角で止める（丸めて削った角の上に直線を残さない）
    const connectors: Line[] = [];
    /** 棒 b の、画面の位置 level にある辺を丸めた半径。その辺を丸めていなければ 0 */
    const roundedAt = (b: BarLayout, level: number): number =>
        Math.max(
            0,
            maxOf(b.segments.map((s) => {
                if (!s.corner) return 0;
                const edgeOf: Record<CornerSide, number> = {
                    top: s.rect.y,
                    bottom: s.rect.y + s.rect.height,
                    left: s.rect.x,
                    right: s.rect.x + s.rect.width,
                };
                return s.corner.sides.some((side) => Math.abs(edgeOf[side] - level) < 0.5) ? cornerRadiusOf(s.rect, s.corner) : 0;
            }))
        );
    for (let k = 0; k + 1 < laid.length; k++) {
        // 累計が軸の範囲の外なら引かない（描く範囲で切れて見えない線。端に貼り付けて引くと、無い高さに線が出る）
        if (!inRange(laid[k].bar.to)) continue;
        const level = pos(laid[k].bar.to);
        const a = laid[k].rect;
        const b = laid[k + 1].rect;
        const trimStart = roundedAt(laid[k], level);
        const trimEnd = roundedAt(laid[k + 1], level);
        const [u0, u1] = horizontal ? [a.y + trimStart, b.y + b.height - trimEnd] : [a.x + trimStart, b.x + b.width - trimEnd];
        const line = horizontal ? { x1: level, y1: u0, x2: level, y2: u1 } : { x1: u0, y1: level, x2: u1, y2: level };
        const last = connectors[connectors.length - 1];
        const sameLevel = last && Math.abs((horizontal ? last.x1 : last.y1) - level) < 0.5;
        if (sameLevel) {
            // 同じ高さの線は、端を伸ばして 1 本にする
            if (horizontal) last.y2 = Math.max(last.y2, u1);
            else last.x2 = Math.max(last.x2, u1);
        } else {
            connectors.push(line);
        }
    }

    // --- 目盛り・グリッド線（グリッド線は中身の長さいっぱい） --------------------------------------
    const plotEnd = horizontal ? plot.y + plot.height : plot.x + plot.width;
    // categoryOrigin は先頭の余白（leadIn）を足した後。中身の端は余白ごと数える
    const contentEnd = scrollAxis ? (horizontal ? plot.y : plot.x) + content : plotEnd;
    const ticks: TextLayout[] = style.valueAxis.show
        ? shownTicks.map(
              (t): TextLayout =>
                  horizontal
                      ? {
                            x: pos(t.value),
                            y: switchAxis ? plot.y - 6 : plot.y + plot.height + tickFont.size + AXIS_TEXT_DROP,
                            anchor: "middle",
                            lines: [t.label],
                            rotate: 0,
                            font: tickFont,
                            lineHeight: tickFont.size + 2,
                        }
                      : {
                            x: switchAxis ? plot.x + plot.width + 6 : plot.x - 6,
                            y: pos(t.value) + tickFont.size * 0.35,
                            anchor: switchAxis ? "start" : "end",
                            lines: [t.label],
                            rotate: 0,
                            font: tickFont,
                            lineHeight: tickFont.size + 2,
                        }
          )
        : [];
    const gridlines = shownTicks.map((t) =>
        horizontal
            ? { x1: pos(t.value), y1: plot.y, x2: pos(t.value), y2: contentEnd }
            : { x1: plot.x, y1: pos(t.value), x2: contentEnd, y2: pos(t.value) }
    );

    // --- 目標の線・定数線（名前は下の「ラベルの置き方」で置く） ------------------------------------------------
    const referenceOf = (ref: ViewModel["target"]): ReferenceLine | null => {
        if (!ref) return null;
        // 定数線は軸を広げない。値を並べる範囲の外なら描かない（目標は軸に入れてある）
        if (ref.value < domainMin || ref.value > domainMax) return null;
        const p = pos(ref.value);
        const line = horizontal ? { x1: p, y1: plot.y, x2: p, y2: contentEnd } : { x1: plot.x, y1: p, x2: contentEnd, y2: p };
        return { line, color: ref.color, width: ref.width, dash: ref.dash, cap: ref.cap, label: null, tooltip: ref.tooltip };
    };
    const target = referenceOf(viewModel.target);
    const constantLine = referenceOf(viewModel.constantLine);
    if (constantLine && viewModel.constantLine) {
        // 定数線の透過性・前後・網掛け領域（標準の Y 軸の定数線と同じ。網掛けは値の小さい側か大きい側の、値を並べる範囲の端まで）
        const cs = viewModel.constantLine.style;
        constantLine.opacity = Math.max(0, Math.min(1, 1 - cs.transparency / 100));
        constantLine.layer = cs.layer;
        if (cs.shadeShow) {
            const p = horizontal ? constantLine.line.x1 : constantLine.line.y1;
            const edge = pos(viewModel.constantLine.shadeRegion === "above" ? domainMax : domainMin);
            const [from, to] = [Math.min(p, edge), Math.max(p, edge)];
            const fill = { color: cs.shadeColor, opacity: Math.max(0, Math.min(1, 1 - cs.shadeTransparency / 100)) };
            constantLine.shade = horizontal
                ? { x: from, y: plot.y, width: to - from, height: contentEnd - plot.y, ...fill }
                : { x: plot.x, y: from, width: contentEnd - plot.x, height: to - from, ...fill };
        }
    }

    // --- 軸のタイトル ---------------------------------------------------------------------------------
    // タイトルは沿う辺（描く範囲の長さ）に収まるように後ろを「…」で省く。縦向きの Y 軸の上の単位は、目盛りの列を単位の幅まで
    // 広げてあるので、タイトル（列の外）には掛からない
    const titleText = (text: string, font: RenderFont, x: number, y: number, rotate: number, room: number): TextLayout => ({
        x,
        y,
        anchor: "middle",
        lines: [truncateToWidth(text, Math.max(0, room), specOf(font))],
        rotate,
        font,
        lineHeight: font.size + 2,
    });
    let valueTitleLayout: TextLayout | null = null;
    if (valueTitle && valueTitleFont) {
        if (!horizontal) {
            // 左右どちらも -90°（下から上へ読む）。-90° の文字は、ベースラインから左へ文字の上側、右へ下側が出る。字の外側の端を、
            // 凡例の内側の端（凡例が無ければビューの端）から標準と同じだけ離す（上の valueTitleBaseline）
            // （1.3.0.0 までは右を 90° にし、文字の上側が右へ出るのにベースラインを 0.3 文字しか内側にせず、右端で切れていた）
            const x = switchAxis ? W - reserve.right - valueTitleBaseline : reserve.left + valueTitleBaseline;
            valueTitleLayout = titleText(valueTitle.text, valueTitleFont, x, plot.y + plot.height / 2, -90, plot.height);
        } else {
            const y = switchAxis
                ? plot.y - (style.valueAxis.show ? tickFont.size + 8 : 4) - 4
                : H - reserve.bottom - noticeHeight - bottomTitleBelowOf(BAR_TITLE_BOTTOM, valueTitleFont.size, valueTitleInk) - (badgeTopRight ? 0 : badgeHeight);
            valueTitleLayout = titleText(valueTitle.text, valueTitleFont, plot.x + plot.width / 2, y, 0, plot.width);
        }
    }
    let categoryTitleLayout: TextLayout | null = null;
    if (categoryTitle && categoryTitleFont) {
        // 縦向きは項目名のすぐ下（斜めの名前でも、名前の下端から少しだけ空ける）
        categoryTitleLayout = !horizontal
            ? titleText(
                  categoryTitle.text,
                  categoryTitleFont,
                  plot.x + plot.width / 2,
                  plot.y + plot.height + nearRoom + categoryHeight + categoryTitleFont.size * AXIS_TITLE_ASCENT,
                  0,
                  plot.width
              )
            : titleText(
                  categoryTitle.text,
                  categoryTitleFont,
                  reserve.left + sideTitleBaselineOf(categoryTitleFont.size, categoryTitleInk),
                  plot.y + plot.height / 2,
                  -90,
                  plot.height
              );
    }

    // --- 凡例 -------------------------------------------------------------------------------------------
    // 共通の組み方（箱の中の座標）を、ビューの座標にする。印は行の縦の真ん中にそろえ、字のベースラインは legendBaselineOf
    let legend: LegendLayout | null = null;
    if (legendBox) {
        const { box } = legendBox;
        const baselineOf = (centerY: number) => legendBaselineOf(box.y + centerY, legendFont.size);
        const title: TextLayout | null = legendBox.title && {
            x: box.x + legendBox.title.x,
            y: baselineOf(legendBox.title.y),
            anchor: "start",
            lines: [legendBox.title.text],
            rotate: 0,
            font: { ...legendFont, weight: "bold" },
            lineHeight: legendFont.size + 2,
        };
        const items: LegendLayout["items"] = legendBox.items.map((i) => {
            const item = viewModel.legend[i.index];
            const x = box.x + i.x;
            // 目標の線は、見える左の端が x に来るように引く（どこも見えない線は x から）
            const ink = targetInkOf(item);
            const lineX1 = item.line ? x - (ink?.start ?? 0) : undefined;
            return {
                item,
                index: i.index,
                x,
                swatchWidth: i.glyphWidth,
                y: box.y + i.y - swatch / 2,
                textX: box.x + i.textX,
                textY: baselineOf(i.y),
                text: i.text,
                ...(lineX1 === undefined ? {} : { lineX1, lineX2: lineX1 + lineGlyphLength }),
            };
        });
        const shift = (arrow: LegendArrow): LegendArrow => ({ ...arrow, x: box.x + arrow.x, y: box.y + arrow.y });
        const arrows = legendBox.arrows && { prev: shift(legendBox.arrows.prev), next: shift(legendBox.arrows.next) };
        legend = { title, items, font: legendFont, swatch, box: { ...box }, pageCount: legendBox.pageCount, page: legendBox.page, arrows };
    }

    // 単位のラベル：縦向きは値の軸の目盛りの列にそろえて、目盛りを引く範囲の上の端の badgeRise 上に置く（ドリルの位置を並べるなら、
    // 「プロット エリアの右上」と同じく上の行）。軸のタイトルとは、目盛りの列を単位の幅まで広げてあるので重ならない。
    // 横向きは値の軸の端（右）にそろえ、値の軸が下なら目盛りとタイトルの下、上ならその上に置く（1.11 までは左上）。
    // 「プロット エリアの右上」は、どちらの向きでも描く範囲の右端にそろえて上に置く
    let badge: TextLayout | null = null;
    if (badgeHeight) {
        const badgeWidth = unitWidthOf(badgeText);
        const [x, anchor]: [number, TextLayout["anchor"]] = horizontal || badgeTopRight
            ? // 単位が長くても左の凡例・ビューの外に出ないように、左端で止める
              [Math.max(Math.min(plot.x + plot.width, W - reserve.right - 2), reserve.left + 2 + badgeWidth), "end"]
            : switchAxis
              ? [Math.min(plot.x + plot.width + 6, W - reserve.right - 2 - badgeWidth), "start"]
              : [Math.max(plot.x - 6, reserve.left + 2 + badgeWidth), "end"];
        badge = {
            x,
            // 横向きで下に置くときは、文字の下側（0.22 文字）までがとった行に収まるように（文字が大きくても凡例・知らせに掛けない）
            y: !horizontal
                ? (badgeTopRight || drillRowShown ? rowLine : Math.min(pos(axis.min), pos(axis.max))) - badgeRise
                : !switchAxis && !badgeTopRight
                  ? H - reserve.bottom - noticeHeight - 1 - badgeFont.size * 0.22
                  : reserve.top + badgeFont.size,
            anchor,
            lines: [badgeText],
            rotate: 0,
            font: badgeFont,
            lineHeight: badgeFont.size + 2,
        };
    }
    // ドリルダウンした位置。縦向きは描く範囲の上の行に、描く範囲の左端から右端まで置く。左の列に単位のラベル（Y 軸の上の単位）があれば
    // その右端（描く範囲の左端の 6px 外）から 10px 空ける（6px だと単位と位置が 1 つの文字に見えた）。ベースラインは単位と同じ。
    // 「プロット エリアの右上」の単位、Y 軸を右に移したときの単位（右端の外から左へ伸びる）があれば、その手前 8px まで。
    // 横向きは自分の行（上の凡例の下）の左端から。入りきらなければ末尾を「…」で省く（ツールヒントに全部）
    const rowDrillPathSpot = () => {
        const badgeBox = badge ? boxOf(badge, 0) : null;
        const onRight = switchAxis || badgeTopRight;
        const x = badgeBox && !onRight ? Math.max(plot.x, badgeBox.x + badgeBox.width + 10) : plot.x;
        const right = badgeBox && onRight ? Math.min(plot.x + plot.width, badgeBox.x - 8) : plot.x + plot.width;
        return { x, y: rowLine - badgeRise, width: right - x };
    };
    const drillPathSpot: { x: number; y: number; width: number } | null = !viewModel.drillPath
        ? null
        : !horizontal
          ? drillRowShown
              ? rowDrillPathSpot()
              : null
          : { x: reserve.left + 2, y: reserve.top - drillPathRow + 3 + drillPathFont.size, width: W - reserve.left - reserve.right };
    const drillPathText = drillPathSpot ? truncateToWidth(viewModel.drillPath, Math.max(0, drillPathSpot.width - 4), specOf(drillPathFont)) : "";
    const drillPath: (TextLayout & { title: string }) | null = drillPathSpot && drillPathText
        ? {
              x: drillPathSpot.x,
              y: drillPathSpot.y,
              anchor: "start",
              lines: [drillPathText],
              rotate: 0,
              font: drillPathFont,
              lineHeight: drillPathFont.size + 2,
              title: viewModel.drillPath,
          }
        : null;
    // --- ラベルの置き方（docs/waterfall.md「見せ方」） --------------------------------------------------------
    // 1. ぶつかってはいけない物を集める：棒、接続線、目標の線、定数線、軸（目盛り・タイトル）、項目名の側、凡例、単位のラベル
    // 2. ラベルごとに、棒のすぐそばの候補を好ましい順に作る（書式の「位置」が自動でなければ、その 1 つだけ）
    // 3. 大事なラベルから置く（合計 → 小計 → 増減の絶対値の大きい順。そのあと定数線・目標の名前）。
    //    候補を上から試し、何ともぶつからない最初の場所に置く。どこにも置けなければ出さない（値はツールヒントで見られる）
    // 帯を「どこまでも」伸ばす長さ。スクロールする中身より長くとる（100 万 px 固定だと、項目がとても多いと先の方に届かない）
    const BIG = 1e6 + content;
    /** 中身（スクロールするときはスクロールする側）の範囲。ラベルはこの中に出す */
    const contentBox: Rect = !scrollAxis
        ? { x: 0, y: 0, width: W, height: H }
        : horizontal
          ? { x: 0, y: plot.y, width: W, height: content }
          : { x: plot.x, y: 0, width: content, height: H };
    /**
     * 窓（ビューの座標）。スクロールしないときはビュー全体。スクロール バーは窓の端に出るので、凡例・知らせの手前で止める
     * （スクロール バーの幅は描く範囲の外に取ってある）
     */
    const regionBox: Rect = !scrollAxis
        ? { x: 0, y: 0, width: W, height: H }
        : horizontal
          ? { x: 0, y: plot.y, width: W - reserve.right, height: plot.height }
          : { x: plot.x, y: 0, width: plot.width, height: H - reserve.bottom - noticeHeight };
    // 斜めの項目名は、右端の文字の上側が (下端 + ROTATED_GAP) より 0.75 文字ほど上に出る
    const rotatedRise = rotated ? Math.min(0, ROTATED_GAP - categorySize * 0.75) : 0;
    /** 項目名の側の端（縦は下端の y、横は左端の x）。ここから外は項目名・X 軸のタイトル */
    const nearEdge = horizontal ? plot.x - nearRoom : plot.y + plot.height + nearRoom + rotatedRise;
    const nearBand: Obstacle = horizontal
        ? { kind: "box", box: { x: -BIG, y: -BIG, width: BIG + nearEdge, height: 2 * BIG } }
        : { kind: "box", box: { x: -BIG, y: nearEdge, width: 2 * BIG, height: BIG } };

    // スクロールしない側（ビューの座標）に描く物：目盛り・軸のタイトル・単位のラベル・凡例・知らせ
    const overlayBoxes: Rect[] = [
        ...ticks.map((t) => boxOf(t, 0)),
        ...(valueTitleLayout ? [boxOf(valueTitleLayout, 0)] : []),
        ...(categoryTitleLayout ? [boxOf(categoryTitleLayout, 0)] : []),
        ...(badge ? [boxOf(badge, 0)] : []),
        // 描く範囲の上の行に並べたドリルの位置（横向きで自分の行をとるときは、下の上の余白に含まれる）
        ...(drillPath ? [boxOf(drillPath, 0)] : []),
        // 凡例の場所（上は凡例が無くても題名の下の余白。縦向きで上の凡例が無ければ、上の行の字と同じく TOP_TEXT_GAP 空ける）
        { x: 0, y: 0, width: W, height: Math.max(reserve.top, TOP_TEXT_GAP) },
        ...(reserve.bottom ? [{ x: 0, y: H - noticeHeight - reserve.bottom, width: W, height: reserve.bottom }] : []),
        ...(reserve.left ? [{ x: 0, y: 0, width: reserve.left, height: H }] : []),
        ...(reserve.right ? [{ x: W - reserve.right, y: 0, width: reserve.right, height: H }] : []),
        ...(noticeHeight ? [{ x: 0, y: H - noticeHeight, width: W, height: noticeHeight }] : []),
    ];
    // スクロールするときは、窓に掛かる物だけを、スクロールの方向いっぱいに伸ばして中身の座標にする（どこまでスクロールしても重ならないように）
    const overlayInContent: Rect[] = !scrollAxis
        ? overlayBoxes
        : overlayBoxes
              .filter((r) => hitsBox(r, regionBox))
              .map((r) => (scrollAxis === "x" ? { x: -BIG, y: r.y, width: 2 * BIG, height: r.height } : { x: r.x, y: -BIG, width: r.width, height: 2 * BIG }));
    const barOwner = (b: Bar) => `bar:${b.key}`;
    const lineObstacle = (ref: ReferenceLine, owner: string): Obstacle => ({ kind: "segment", ...ref.line, width: ref.width, owner });
    const contentObstacles: Obstacle[] = [
        ...laid.flatMap((b): Obstacle[] => [
            { kind: "box", box: b.rect, owner: barOwner(b.bar) },
            // ハイライトの該当分は棒からはみ出すことがある
            ...b.segments.flatMap((s): Obstacle[] => (s.highlight ? [{ kind: "box", box: s.highlight, owner: barOwner(b.bar) }] : [])),
        ]),
        ...(style.connectors.show ? connectors.map((c): Obstacle => ({ kind: "segment", ...c, width: style.connectors.width })) : []),
        ...(target ? [lineObstacle(target, "ref:target")] : []),
        ...(constantLine ? [lineObstacle(constantLine, "ref:constant")] : []),
        // 項目名の側（項目名・X 軸のタイトル・スクロール バー）
        nearBand,
        ...(scrollAxis === "y" ? [asBox({ x: W - reserve.right - SCROLLBAR, y: -BIG, width: SCROLLBAR, height: 2 * BIG })] : []),
        ...overlayInContent.map(asBox),
    ];

    // データ ラベルの候補：外の先（増えた棒・正の合計は値の大きい側、減った棒・負の合計は小さい側）→ 中の先
    const labelFill = style.dataLabels.background;
    const labelPad = labelFill ? LABEL_PADDING : HIT_PAD;
    // 横向きの棒の横書きのラベルの背景は、字の送りの左右に H_LABEL_BACKGROUND_PAD（標準の集合横棒と同じ。1.15.8.37 までは上下と同じ LABEL_PADDING）
    const labelPadX = labelFill && horizontal && !verticalText ? H_LABEL_BACKGROUND_PAD : labelPad;
    /** 棒の中に入るかを見るときの、字の端から背景の四角の端まで（値の向き。横向きの棒の横書きのラベルに背景を描くときだけ） */
    const boxPadAlong = labelFill && horizontal && !verticalText ? H_LABEL_BACKGROUND_PAD : 0;
    type Spot = "outEnd" | "inEnd" | "inHigh" | "inCenter" | "inLow";
    const position = style.dataLabels.position;
    const spots: Spot[] =
        position === LABEL_POSITIONS.outsideEnd
            ? ["outEnd"]
            : position === LABEL_POSITIONS.insideTop
              ? ["inHigh"]
              : position === LABEL_POSITIONS.insideCenter
                ? ["inCenter"]
                : position === LABEL_POSITIONS.insideBottom
                  ? ["inLow"]
                  : // 自動は、棒の先の外、入らなければ棒の中の先（標準と同じ）。反対の端へずらして置くことはしない。
                    // どこにも置けない（ほかのラベル・線・棒と重なる）ラベルは出さない
                    ["outEnd", "inEnd"];
    const labelRequestOf = (b: BarLayout): LabelRequest<TextLayout> => {
        const { bar, rect } = b;
        const lines = horizontal && !verticalText ? [labelTextOf(bar)] : [bar.labelText, ...(bar.rateText ? [bar.rateText] : [])];
        const baseFont = labelFontOf(bar);
        const lineHeight = baseFont.size + 2;
        const textLength = maxOf(lines.map((l) => measureTextWidth(l, specOf(baseFont))));
        const size = baseFont.size;
        /** 1 行目から最後の行までのベースラインのあいだ */
        const span = (lines.length - 1) * lineHeight;
        /** 値の向きに文字の高さが並ぶか（縦向きの横書き・横向きの縦書き）。並ばなければ文字の幅が並ぶ */
        const stacked = horizontal === verticalText;
        /** 文字を置く真ん中から、1 行目のベースラインまで（値の向きの座標。下の text の置き方と同じ） */
        const firstBaselineOffset = size * 0.35 - span / 2;
        const [low, high] = extentOf(bar);
        const highEnd = pos(high);
        const lowEnd = pos(low);
        /** 棒の値の向きの長さ */
        const barLength = horizontal ? rect.width : rect.height;
        // 棒の先（増えた棒・正の合計は値の大きい側、減った棒・負の合計は小さい側）と、そこから外へ向かう画面の向き
        const positive = bar.value >= 0;
        const endAt = positive ? highEnd : lowEnd;
        const outward = positive ? growth : -growth;
        /** 端 edge に付けて、dir の側に置くときの真ん中と、棒の中に入るか（LABEL_EDGE の置き方） */
        const attach = (edge: number, dir: number): { center: number; fits: boolean } => {
            if (stacked) {
                const { baseline, reach } = attachedBaseline(edge, dir, size, span, labelInkPad);
                return { center: baseline - firstBaselineOffset, fits: reach + LABEL_EDGE.fitSlack <= barLength };
            }
            // 背景を描くなら、棒の中に入るかは背景の四角（字の左右に boxPadAlong）で見る（棒の外へ出さない）。1.15.8.38 の最初の作りまでは
            // 字の四角で見ていて、背景が棒の端から最大 2.5px 出た
            return {
                center: edge + dir * (LABEL_EDGE.textGap + textLength / 2),
                fits: LABEL_EDGE.textGap + textLength + Math.max(LABEL_EDGE.textFitSlack, boxPadAlong) <= barLength,
            };
        };
        const spotOf = (spot: Spot): { center: number; fits: boolean } => {
            switch (spot) {
                case "outEnd":
                    return attach(endAt, outward);
                case "inEnd":
                    return attach(endAt, -outward);
                case "inHigh":
                    return attach(highEnd, -growth);
                case "inLow":
                    return attach(lowEnd, growth);
                default: {
                    // 内側中央は真ん中に置き、数字の四角（文字の幅が並ぶなら文字の幅。背景を描くなら背景の四角の幅）が棒の長さに収まれば入る
                    const digits = stacked ? span + (LABEL_EDGE.digitAscent + LABEL_EDGE.digitDescent) * size : textLength + boxPadAlong * 2;
                    return { center: (highEnd + lowEnd) / 2, fits: digits <= barLength };
                }
            }
        };
        const across = horizontal ? rect.y + rect.height / 2 : rect.x + rect.width / 2;
        /** 四角を、値の向きに棒の端から端までで切る */
        const clipAlongValue = (box: Rect, bar: Rect): Rect => {
            if (horizontal) {
                const x0 = Math.max(box.x, bar.x);
                return { ...box, x: x0, width: Math.max(0, Math.min(box.x + box.width, bar.x + bar.width) - x0) };
            }
            const y0 = Math.max(box.y, bar.y);
            return { ...box, y: y0, height: Math.max(0, Math.min(box.y + box.height, bar.y + bar.height) - y0) };
        };
        // 比べる形の小計は、見せ方によって棒の中の色が違う（透かす・枠だけ・斜線・点線の枠）。その色で文字色を決める
        const rawColor = bar.segments[0]?.color ?? bar.color;
        const segmentColor = bar.kind === BAR_KINDS.subtotal ? subtotalPaint(style, rawColor).under : rawColor;
        const explicit = isTotalBar(bar) ? style.dataLabels.totalFont.color : style.dataLabels.font.color;
        /**
         * 棒の外の候補が中身（contentBox）から項目の向き（縦向きは左右、横向きは上下）にはみ出すなら、収まるだけ内側へ寄せた候補。
         * はみ出さない・中身より長いなら無し。値の向きは変えない（はみ出せば、寄せた候補も置けない）
         */
        const shiftedInto = (c: Candidate<TextLayout>): Candidate<TextLayout> | null => {
            const [from, to] = horizontal ? [c.box.y, c.box.y + c.box.height] : [c.box.x, c.box.x + c.box.width];
            const [lo, hi] = horizontal ? [contentBox.y, contentBox.y + contentBox.height] : [contentBox.x, contentBox.x + contentBox.width];
            if (to - from > hi - lo) return null;
            const d = from < lo ? lo - from : to > hi ? hi - to : 0;
            if (d === 0) return null;
            // 文字・枠・ほかのラベルと比べる四角を同じだけ動かす
            return horizontal
                ? { item: { ...c.item, y: c.item.y + d }, box: { ...c.box, y: c.box.y + d }, hit: c.hit && { ...c.hit, y: c.hit.y + d } }
                : { item: { ...c.item, x: c.item.x + d }, box: { ...c.box, x: c.box.x + d }, hit: c.hit && { ...c.hit, x: c.hit.x + d } };
        };
        const candidates = spots.flatMap((spot): Array<Candidate<TextLayout>> => {
            const inside = spot !== "outEnd";
            const { center: at, fits } = spotOf(spot);
            // 棒の中の候補は、値の向きに入るときだけ
            if (inside && !fits) return [];
            const cx = horizontal ? at : across;
            const cy = horizontal ? across : at;
            // 自動の文字色は、文字のすぐ下の色で決める：背景を出すなら背景（透過性があれば下の色と混ぜた色）、
            // 出さないなら置いた場所の色（棒の中は棒の色、外はビジュアルの背景）。
            // 1.7.0.0 までは背景を出していても棒の中なら棒の色で決め、白い背景に白い文字を描いていた（第 32 ページ ③）
            const underneath = inside ? segmentColor : style.background;
            const base = labelFill ? blend(labelFill.color, underneath, 1 - labelFill.opacity) : underneath;
            // 棒の中は白を、外は灰色を優先する（標準と同じ。背景の既定は黒・透過性 90% なので、棒の中は白、外は灰色になる）
            // 「符号の色」は棒の外か背景を付けたラベルだけ（合計には付けない）。棒の中は今までどおり読める色（標準も中のラベルは白）
            const tone = bar.labelToneColor && (!inside || labelFill) ? bar.labelToneColor : "";
            const color =
                tone ||
                explicit ||
                autoLabelColor(inside, base);
            const font = { ...baseFont, color };
            // 真ん中 (cx, cy) に置く。複数行は 1 行目のベースラインを上へずらす
            const text: TextLayout = verticalText
                ? { x: cx + firstBaselineOffset, y: cy, anchor: "middle", lines, rotate: -90, font, lineHeight }
                : { x: cx, y: cy + firstBaselineOffset, anchor: "middle", lines, rotate: 0, font, lineHeight };
            text.background = backgroundOf(lines, "middle", font, lineHeight, labelFill, labelPadX);
            text.halo = style.dataLabels.halo;
            // 棒の中の候補は、自分の棒の中に収まるときだけ。ほかのラベルとは文字の縁も入れて見る（hit。切らない）。
            // 棒・線などとのぶつかりを見る四角（box）は、棒の中の候補なら値の向きに棒の中で切る。内側中央は数字の四角が棒に収まれば出すので、文字の四角の下側
            // （数字の無い所）と余白が棒の端から少し出る。そこで棒の端の接続線・項目名の側とぶつけて出さなくすると、上の「入る」で
            // 決めたことにならない。切った四角を自分の棒の中に収める（within）ので、within が見るのは棒の幅（太さ）の向きだけになる
            const box = boxOf(text, labelPad, labelPadX);
            const candidate: Candidate<TextLayout> = {
                item: text,
                box: inside ? clipAlongValue(box, rect) : box,
                hit: hitBoxOf(text),
                within: inside ? rect : undefined,
            };
            // 端の棒の外の候補が中身からはみ出すなら、その直後に、内側へ寄せた候補を試す（棒の中の候補は、自分の棒から離れるので寄せない）
            const shifted = inside ? null : shiftedInto(candidate);
            return shifted ? [candidate, shifted] : [candidate];
        });
        return { owner: barOwner(bar), candidates };
    };
    const rank = (b: BarLayout) => (b.bar.kind === BAR_KINDS.total ? 0 : b.bar.kind === BAR_KINDS.subtotal ? 1 : 2);
    const labelled = style.dataLabels.show
        ? laid
              .map((b, index) => ({ b, index }))
              .filter(({ b }) => b.bar.labelText)
              .sort((p, q) => rank(p.b) - rank(q.b) || Math.abs(q.b.bar.value) - Math.abs(p.b.bar.value) || p.index - q.index)
              .map(({ b }) => b)
        : [];
    const labelRequests = labelled.map(labelRequestOf);

    // 目標・定数線の名前の候補：縦向きは線の右端の上 → 右端の下 → 左端の上・下 → 線に沿って右から左へずらした所の上・下。
    // 横向きは描く範囲の上（線の真上 → 左 → 右。Y 軸を上に切り替えたときは目盛り・タイトルのさらに上）
    const referenceTitleTop = horizontal && switchAxis ? plot.y - (style.valueAxis.show ? tickFont.size + 8 : 4) - valueTitleSize : plot.y;
    const nameRequestOf = (source: ViewModel["target"], ref: ReferenceLine, owner: string): LabelRequest<TextLayout> | null => {
        if (!source || !source.label) return null;
        const font: RenderFont = { ...labelFont, color: source.color };
        const text = (x: number, y: number, anchor: TextLayout["anchor"]): TextLayout => ({
            x,
            y,
            anchor,
            lines: [source.label],
            rotate: 0,
            font,
            lineHeight: font.size + 2,
        });
        const texts: TextLayout[] = [];
        const p = horizontal ? ref.line.x1 : ref.line.y1;
        if (horizontal) {
            const y = referenceTitleTop - 4;
            texts.push(text(p, y, "middle"), text(p - 3, y, "end"), text(p + 3, y, "start"));
        } else {
            const right = plot.x + plot.width;
            const above = p - 4;
            const below = p + 4 + font.size * 0.8;
            texts.push(text(right, above, "end"), text(right, below, "end"), text(plot.x + 2, above, "start"), text(plot.x + 2, below, "start"));
            if (!scrollAxis) {
                const width = measureTextWidth(source.label, specOf(font));
                const stride = Math.max(8, categoryStepLength / 2);
                for (let r = right - stride; r - width >= plot.x; r -= stride) texts.push(text(r, above, "end"), text(r, below, "end"));
            }
        }
        return { owner, candidates: texts.map((item) => ({ item, box: boxOf(item, HIT_PAD) })) };
    };
    const nameRequests: Array<{ ref: ReferenceLine; request: LabelRequest<TextLayout> }> = [];
    const pairs: Array<[ViewModel["target"], ReferenceLine | null, string]> = [[viewModel.target, target, "ref:target"]];
    // 定数線のデータ ラベルは標準の「水平方向の位置」「縦位置」で決まった場所に置く（スクロールしない側。ほかのラベルは避けない。
    // 文字の縁で読めるようにする）。1.15.5.16 までは目標の名前と同じく、ラベルを避けて置いていた
    let constantBox: Rect | null = null;
    const cl = viewModel.constantLine;
    if (constantLine && cl?.label) {
        const cs = cl.style;
        const fontSize = cs.labelFontSize * PT_TO_PX;
        const font: RenderFont = { ...categoryFont, size: fontSize, weight: "normal", style: "normal", decoration: "none", color: cs.labelColor };
        const at = horizontal ? constantLine.line.x1 : constantLine.line.y1;
        const valueRange: [number, number] = horizontal ? [plot.x, plot.x + plot.width] : [plot.y, plot.y + plot.height];
        const categoryRange: [number, number] = horizontal ? [plot.y, plot.y + plot.height] : [plot.x, plot.x + plot.width];
        const [spot] = placeReferenceLabels(
            [{ label: cl.label, at }],
            !horizontal,
            valueRange,
            categoryRange,
            fontSize,
            (text) => measureTextWidth(text, specOf(font)),
            cs.labelAvoidOverlap,
            cs.labelHorizontal,
            cs.labelVertical
        );
        constantLine.label = {
            x: spot.x,
            y: spot.y,
            anchor: spot.anchor,
            lines: [cl.label],
            rotate: 0,
            font,
            lineHeight: fontSize + 2,
            halo: { color: style.highContrast?.background ?? style.background, width: 3 },
        };
        constantBox = hitBoxOf(constantLine.label);
    }
    for (const [source, ref, owner] of pairs) {
        const request = ref ? nameRequestOf(source, ref, owner) : null;
        if (ref && request) nameRequests.push({ ref, request });
    }

    if (!scrollAxis) {
        // スクロールしないときは、ラベルと線の名前を 1 度に置く（名前はラベルのあと）
        // データ ラベルは定数線のデータ ラベルを避けない（標準と同じ）。目標の名前は両方を避ける
        const placed = placeLabels(labelRequests, contentObstacles, { bounds: contentBox });
        labelled.forEach((b, i) => (b.label = placed[i]?.item ?? null));
        const labelBoxes = labelled.flatMap((b) => (b.label ? [asBox(hitBoxOf(b.label))] : []));
        const names = placeLabels(
            nameRequests.map((n) => n.request),
            [...contentObstacles, ...labelBoxes, ...(constantBox ? [asBox(constantBox)] : [])],
            { bounds: contentBox }
        );
        nameRequests.forEach((n, i) => (n.ref.label = names[i]?.item ?? null));
    } else {
        // スクロールするときは、ラベルは中身の座標で置く。線の名前は動かない側（ビューの座標）に描くので、
        // 動かない物（目盛り・凡例など）と項目名の側、もう 1 本の名前だけを避ける（棒とラベルはスクロールで動く）
        const placed = placeLabels(labelRequests, contentObstacles, { bounds: contentBox });
        labelled.forEach((b, i) => (b.label = placed[i]?.item ?? null));
        const names = placeLabels(
            nameRequests.map((n) => n.request),
            [...overlayBoxes.map(asBox), nearBand, ...(constantBox ? [asBox(constantBox)] : [])],
            { bounds: { x: 0, y: 0, width: W, height: H } }
        );
        nameRequests.forEach((n, i) => (n.ref.label = names[i]?.item ?? null));
    }

    // 好ましい置き場所（最初の候補）が項目名の側へはみ出して置けなかったラベルがあれば、はみ出した分を返す。
    // layoutOf が描く範囲と項目名のあいだを空けて置き直す。1px 以下のはみ出しも返す（項目名の側とは少しでも重なれば置けない。
    // 1.15.8.30 までは 1px を超えるはみ出しだけを返し、空け直して 1px 以下が残ると、そこで止まってラベルを出さなかった）
    let nearOverflow = 0;
    /** 好ましい置き場所（最初の候補）に置けたラベルの数 */
    let firstPlaced = 0;
    labelRequests.forEach((request, i) => {
        if (request.candidates[0] && labelled[i].label === request.candidates[0].item) firstPlaced++;
        const first = request.candidates[0];
        if (!first || labelled[i].label === first.item) return;
        const over = horizontal ? nearEdge - first.box.x : first.box.y + first.box.height - nearEdge;
        if (over > 0) nearOverflow = Math.max(nearOverflow, over);
    });

    const layout: ChartLayout = {
        horizontal,
        plot,
        scroll: {
            axis: scrollAxis,
            // スクロールしないときは、ビュー全体を窓にする（斜めの項目名などが描く範囲の外にはみ出しても切らない）
            region: regionBox,
            content,
        },
        bars: laid,
        ticks,
        gridlines,
        breaks,
        connectors,
        target,
        constantLine,
        legend,
        valueTitle: valueTitleLayout,
        categoryTitle: categoryTitleLayout,
        badge,
        drillPath,
        notice: viewModel.notice ? { x: 4, y: H - 4 } : null,
        rotatedCategories: rotated,
        categoryStep: step,
    };
    return { layout, nearOverflow, tickTarget, firstPlaced };
}
