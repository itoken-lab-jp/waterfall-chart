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

import { BAR_KINDS, Bar, LegendItem, Segment, TextStyle, TooltipItem, ViewModel, blend, extentOf } from "./viewModel";
import { FontSpec, measureTextWidth, readableText, truncateStartToWidth, truncateToWidth } from "./unitUtils";
import { BREAK_STYLES, LABEL_POSITIONS, ORIENTATIONS } from "./settings";
import { LEVEL_SEPARATOR } from "./data";
import { Candidate, LabelRequest, Obstacle, placeLabels } from "./placement";

import IViewport = powerbi.IViewport;

/** 書式の文字サイズはポイント。SVG にはピクセルで渡す */
export const PT_TO_PX = 4 / 3;
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
const NOTICE_HEIGHT = 16;
/**
 * ビジュアルの上端に空ける余白。題名が折り返したときに、上の凡例・単位のラベルが題名の最後の行に
 * かぶって見えないように空ける（1.3.0.0 までは凡例の文字が上端から 4px、単位は 2px ほどだった）。
 * ビジュアルはビューの外（上）には描いていない（.wf-container が viewport の大きさで overflow: hidden）
 */
const TOP_MARGIN = 8;
/** 下の凡例の上の隙間（px）。項目名・X 軸のタイトルと凡例をくっつけない（標準・棒グラフと同じくらい。1.11 まではなし） */
const LEGEND_GAP_BOTTOM = 10;
const legendGapBottom = (height: number) => Math.max(0, Math.min(LEGEND_GAP_BOTTOM, (height - 150) / 5));
const LEGEND_SWATCH = 10;
/** 凡例の線の印（目標の線）の長さ。破線・点線の模様が見える長さ */
const LEGEND_LINE = 18;
/** 凡例の印の幅（棒は四角、目標は線） */
const swatchWidthOf = (item: LegendItem) => (item.kind === "target" ? LEGEND_LINE : LEGEND_SWATCH);
/** スクロール バーの太さぶん、描く範囲を空ける */
export const SCROLLBAR = 12;
/** データ ラベルの背景の余白 */
const LABEL_PADDING = 2;
/** 自動の色で、棒の外に出すラベル・背景を出すラベルの色（足りなければ白か濃い灰色） */
const LABEL_OUTSIDE_COLOR = "#605E5C";
/**
 * 自動の文字色に求めるコントラスト比。棒の中は白を優先し 3 以上（標準と同じく、テーマの青 #118DFF の上は白）、
 * 外・背景の上は灰色を優先し 4.5 以上
 */
const INSIDE_TEXT_CONTRAST = 3;
const OUTSIDE_TEXT_CONTRAST = 4.5;

export interface Rect {
    x: number;
    y: number;
    width: number;
    height: number;
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
export type CornerSide = "top" | "bottom" | "left" | "right";

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

export interface Corner {
    sides: CornerSide[];
    radius: number;
}

/**
 * 棒（区画）の path。corner があれば、その端（sides）の角だけを丸める（棒グラフと同じ二次曲線）。
 * corner が無い・半径が 0 なら長方形
 */
/** 実際に描く角の半径。棒の太さの半分と、長さ（両端を丸めるなら長さの半分）まで（棒グラフと同じ考え方）。丸めなければ 0 */
export function cornerRadiusOf(r: Rect, corner: Corner | null): number {
    if (!corner || !corner.sides.length) return 0;
    const has = (side: CornerSide) => corner.sides.includes(side);
    const alongHeight = (has("top") ? 1 : 0) + (has("bottom") ? 1 : 0);
    const alongWidth = (has("left") ? 1 : 0) + (has("right") ? 1 : 0);
    return Math.max(
        0,
        Math.min(
            corner.radius,
            alongHeight ? Math.min(r.width / 2, r.height / alongHeight) : Infinity,
            alongWidth ? Math.min(r.height / 2, r.width / alongWidth) : Infinity
        )
    );
}

export function barPath(r: Rect, corner: Corner | null): string {
    const { x, y, width: w, height: h } = r;
    const has = (side: CornerSide) => corner?.sides.includes(side) ?? false;
    const k = cornerRadiusOf(r, corner);
    if (!corner || k <= 0) return `M ${x},${y} h ${w} v ${h} h ${-w} Z`;
    const x2 = x + w;
    const y2 = y + h;
    // 角ごとの半径（左上から時計回り）。丸めない角は 0（二次曲線が点に縮む）
    const tl = has("top") || has("left") ? k : 0;
    const tr = has("top") || has("right") ? k : 0;
    const br = has("bottom") || has("right") ? k : 0;
    const bl = has("bottom") || has("left") ? k : 0;
    return (
        `M ${x + tl},${y} L ${x2 - tr},${y} Q ${x2},${y} ${x2},${y + tr} ` +
        `L ${x2},${y2 - br} Q ${x2},${y2} ${x2 - br},${y2} ` +
        `L ${x + bl},${y2} Q ${x},${y2} ${x},${y2 - bl} ` +
        `L ${x},${y + tl} Q ${x},${y} ${x + tl},${y} Z`
    );
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

export interface LegendLayout {
    title: TextLayout | null;
    /** y は色の四角の上端（線の印は y + 5 に引く）、textY は名前のベースライン、swatchWidth は印の幅（名前は x + swatchWidth + 4 から） */
    items: Array<{ item: LegendItem; index: number; x: number; swatchWidth: number; y: number; textY: number; text: string }>;
    font: RenderFont;
}

export interface ReferenceLine {
    line: Line;
    color: string;
    width: number;
    dash: string | null;
    /** 名前と値（スクロールしない側に置く）。出さなければ null */
    label: TextLayout | null;
    /** 線にマウスを当てたときのツールヒント。空なら出さない（定数線） */
    tooltip: TooltipItem[];
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

type Side = "top" | "bottom" | "left" | "right";

/** 凡例の位置（topLeft など）を、置く辺と寄せ方に分ける */
function legendPlacement(position: string): { side: Side; align: "start" | "center" | "end" } {
    const side: Side = position.startsWith("bottom")
        ? "bottom"
        : position.startsWith("left")
          ? "left"
          : position.startsWith("right")
            ? "right"
            : "top";
    const rest = position.replace(/^(top|bottom|left|right)/, "").toLowerCase();
    const align = rest === "center" ? "center" : rest === "right" || rest === "bottom" ? "end" : "start";
    return { side, align };
}

/** 背景の長方形（(x, y) を原点に、回す前の座標）。文字の寄せ方と行数から出す */
function backgroundOf(
    lines: string[],
    anchor: TextLayout["anchor"],
    font: RenderFont,
    lineHeight: number,
    fill: { color: string; opacity: number } | null
): TextLayout["background"] {
    if (!fill) return null;
    const width = Math.max(0, ...lines.map((l) => measureTextWidth(l, specOf(font)))) + LABEL_PADDING * 2;
    const x = anchor === "start" ? -LABEL_PADDING : anchor === "end" ? -width + LABEL_PADDING : -width / 2;
    const top = -font.size * 0.8 - LABEL_PADDING;
    const height = font.size * 1.02 + (lines.length - 1) * lineHeight + LABEL_PADDING * 2;
    return { x, y: top, width, height, color: fill.color, opacity: fill.opacity };
}

/** 長方形どうしが重なるか（辺が接するだけなら重ならない） */
const hitsBox = (a: Rect, b: Rect): boolean =>
    a.x < b.x + b.width && a.x + a.width > b.x && a.y < b.y + b.height && a.y + a.height > b.y;

/** 四角を障害物にする */
const asBox = (box: Rect): Obstacle => ({ kind: "box", box });

/**
 * 値の軸の目盛りの本数の上限。標準（powerbi-visuals-utils-chartutils の getRecommendedNumberOfTicksForYAxis・ForXAxis）と同じ。
 * 縦の軸は高さ、横の軸は幅で決める
 */
export function recommendedTickCount(length: number, axis: "vertical" | "horizontal"): number {
    const [small, medium] = axis === "vertical" ? [150, 300] : [300, 500];
    return length < small ? 3 : length < medium ? 5 : 8;
}

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
 * 文字の外枠（画面の座標）。pad は四方に足す余白。回転は 0° と -90° だけ見る（ラベル・目標と定数線の名前）。
 * 高さは backgroundOf と同じ見積もり（上に 0.8 文字、下に 0.22 文字）
 */
export function boxOf(text: TextLayout, pad: number): Rect {
    const spec = specOf(text.font);
    const width = Math.max(0, ...text.lines.map((l) => measureTextWidth(l, spec))) + pad * 2;
    const x0 = text.anchor === "start" ? -pad : text.anchor === "end" ? -width + pad : -width / 2;
    const y0 = -text.font.size * 0.8 - pad;
    const height = text.font.size * 1.02 + (text.lines.length - 1) * text.lineHeight + pad * 2;
    // -90° は (dx, dy) → (dy, -dx)
    if (text.rotate === -90) return { x: text.x + y0, y: text.y - x0 - width, width: height, height: width };
    return { x: text.x + x0, y: text.y + y0, width, height };
}

/** ラベルの四角に足す余白（背景の無いラベル）。ほかの物とくっつきすぎないように */
const HIT_PAD = 1;
/** 名前の側の余白を取り直す回数の上限 */
const NEAR_PASSES = 3;

/**
 * 座標の計算。描けないほど狭いとき・棒が無いときは null。
 *
 * 1 回目は項目名の側（縦は下、横は左）に余白をとらずに置く。値の端にある棒のラベルが項目名の側へはみ出すなら、
 * はみ出した分だけ描く範囲と項目名のあいだを空けて置き直す（ラベルが項目名に重ならないように。空けるのは要るときだけ）
 */
export function layoutOf(viewModel: ViewModel, viewport: IViewport): ChartLayout | null {
    // 空けるのはビューの 1/4 まで。それでも入らないラベルは出さない
    const cap = (viewModel.style.orientation === ORIENTATIONS.horizontal ? viewport.width : viewport.height) * 0.25;
    let room = 0;
    let result = layoutWith(viewModel, viewport, room);
    for (let pass = 0; result && result.nearOverflow > 0 && pass < NEAR_PASSES; pass++) {
        const wider = Math.min(cap, room + Math.ceil(result.nearOverflow) + 1);
        if (wider <= room) break;
        room = wider;
        const next = layoutWith(viewModel, viewport, room);
        // 空けると描く範囲が狭くなりすぎるときは、はみ出すラベルを出さない版のまま
        if (!next) break;
        result = next;
    }
    // 描く範囲が足りない・位置の文字が 1 文字も入らないときは、ドリルの位置をとらずに組み直す
    // （位置の行のせいで図ごと消えない・空の位置で行の高さだけとらないように）
    if (viewModel.drillPath && !result?.layout.drillPath) return layoutOf({ ...viewModel, drillPath: "" }, viewport);
    return result?.layout ?? null;
}

function layoutWith(
    viewModel: ViewModel,
    viewport: IViewport,
    nearRoom: number
): { layout: ChartLayout; nearOverflow: number } | null {
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
    /** 上に空けるラベルの行の高さ（大きいほうの文字） */
    const labelLineHeight = Math.max(labelFont.size, totalLabelFont.size) + 2;
    const valueTitle = style.valueAxis.title;
    const categoryTitle = style.categoryAxis.title;
    const valueTitleFont = valueTitle ? renderFontOf(valueTitle.font) : null;
    const categoryTitleFont = categoryTitle ? renderFontOf(categoryTitle.font) : null;
    const valueTitleSize = valueTitleFont ? valueTitleFont.size + 6 : 0;
    const categoryTitleSize = categoryTitleFont ? categoryTitleFont.size + 6 : 0;
    const hasRate = bars.some((b) => b.rateText);
    const verticalText = style.dataLabels.orientation === "vertical";
    const hasParent = bars.some((b) => b.parentLabel);
    const invert = style.valueAxis.invert;
    const switchAxis = style.valueAxis.switchPosition;

    // --- 凡例 ------------------------------------------------------------------------------------
    const legendShown = style.legend.show && viewModel.legend.length > 0;
    const placement = legendPlacement(style.legend.position);
    const legendRow = legendFont.size + 10;
    const legendTitleWidth = style.legend.title
        ? measureTextWidth(style.legend.title, specOf({ ...legendFont, weight: "bold" })) + 8
        : 0;
    const itemWidth = (item: LegendItem) => swatchWidthOf(item) + 4 + measureTextWidth(item.name, specOf(legendFont)) + 12;
    const legendColumnWidth = legendShown
        ? Math.min(W * 0.3, Math.max(legendTitleWidth, ...viewModel.legend.map(itemWidth)) + 8)
        : 0;
    // ドリルダウンした位置。縦向きで単位のラベルを出すときは、その行（描く範囲の上）に並べて行を足さない
    // （2026-09-25 ユーザー「縦に積むとグラフエリアが狭くなる」）。横向き（単位は右下）・単位を出さないときは、上の凡例の下に 1 行とる
    const drillPathFont = renderFontOf(style.categoryAxis.font);
    const badgeShown = Boolean(viewModel.unitBadge) && style.valueAxis.show;
    const pathSharesBadgeRow = Boolean(viewModel.drillPath) && !horizontal && badgeShown;
    // 行の高さは、文字の下側（0.22 文字、boxOf と同じ見積もり）まで収まるようにとる（大きいフォントでも下の物に掛けない）
    const drillPathRow = viewModel.drillPath && !pathSharesBadgeRow ? Math.ceil(drillPathFont.size * 1.3) + 6 : 0;
    const reserve = {
        top: TOP_MARGIN + (legendShown && placement.side === "top" ? legendRow : 0) + drillPathRow,
        // 低いビジュアルでは隙間を縮める（高さ 200px 以上で 10px、150px 以下で 0）。描く範囲が 10px を切ると図ごと消えるため
        bottom: legendShown && placement.side === "bottom" ? legendRow + legendGapBottom(H) : 0,
        left: legendShown && placement.side === "left" ? legendColumnWidth : 0,
        right: legendShown && placement.side === "right" ? legendColumnWidth : 0,
    };
    const noticeHeight = viewModel.notice ? NOTICE_HEIGHT : 0;
    // 単位のラベルの行。ドリルの位置を並べるときは、大きいほうの文字に合わせる（文字の下側まで収める）
    const badgeRowFont = pathSharesBadgeRow ? Math.max(tickFont.size, drillPathFont.size) : tickFont.size;
    // 並べないときは 1.11.18 までと同じ高さ（ドリルしていないときの見た目を変えない）
    const badgeHeight = !badgeShown ? 0 : pathSharesBadgeRow ? badgeRowFont + Math.max(6, Math.ceil(badgeRowFont * 0.22) + 2) : tickFont.size + 6;
    const referenceLabelRoom = viewModel.target?.label || viewModel.constantLine?.label ? labelFont.size + 6 : 0;

    const tickLabelWidth = style.valueAxis.show
        ? // 目盛りの本数は描く範囲の長さで決まる（下の axisTicks）。いちばん細かい 8 本の数字の幅も見積もりに入れる
          Math.max(0, ...[...axis.ticks, ...(axis.ticksFor?.(Math.max(8, style.valueAxis.tickCount)) ?? [])].map((t) => measureTextWidth(t.label, specOf(tickFont))))
        : 0;
    const labelTextOf = (b: Bar) => (b.rateText ? `${b.labelText} (${b.rateText})` : b.labelText);
    const valueLabelWidth = style.dataLabels.show
        ? Math.max(
              0,
              ...bars.map((b) => measureTextWidth(horizontal && !verticalText ? labelTextOf(b) : b.labelText, specOf(labelFontOf(b))))
          )
        : 0;

    // 項目の方向の並び（標準・棒グラフと同じ：外側のパディングとカテゴリ間のスペース）
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

    if (!horizontal) {
        // 値のラベルは縦書きなら文字の幅ぶん、横書きなら行数ぶん上に空ける。
        // 下（項目名の側）は、ラベルがはみ出すときだけ nearRoom で空ける
        const labelRoom = style.dataLabels.show
            ? verticalText
                ? Math.min(valueLabelWidth, H * 0.25) + 6
                : (hasRate ? 2 : 1) * labelLineHeight + 6
            : 4;
        const axisWidth = style.valueAxis.show ? tickLabelWidth + 10 : 4;
        const left = reserve.left + (switchAxis ? 8 : Math.max(8, axisWidth + valueTitleSize));
        const right = reserve.right + (switchAxis ? Math.max(8, axisWidth + valueTitleSize) : 8);
        const width = W - left - right;
        const top = reserve.top + badgeHeight + Math.max(labelRoom, referenceLabelRoom);
        const needed = minWidth > 0 ? n * minWidth : 0;
        if (needed > width + 1) scrollAxis = "x";
        content = scrollAxis ? needed : width;
        const categoryStep = content / (n - pad + 2 * outer);
        const textBand = categoryStep - 4;
        // 項目名：帯に入れば水平（2 行目に上の階層）、入らなければ斜めにして上の階層から出す
        const flatWidth = Math.max(0, ...bars.map((b) => measureTextWidth(b.label, specOf(categoryFontOf(b)))));
        if (!style.categoryAxis.show) {
            categoryTexts = bars.map((b) => ({ lines: [] as string[], title: b.fullLabel }));
        } else if (flatWidth <= textBand) {
            categoryHeight = (hasParent ? 2 : 1) * categoryLineHeight + 6;
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
            // 上限を大きくしても、描く範囲に高さの 4 分の 1 は残す（棒グラフと同じ。1.11 までは 100% で描く範囲が無くなり、何も描かなかった）
            const room = H * 0.75 - top - nearRoom - ROTATED_GAP - categorySize * 0.5 - categoryTitleSize - reserve.bottom - noticeHeight - (scrollAxis ? SCROLLBAR : 0);
            const maxLength = Math.max(0, Math.min(H * style.categoryAxis.maxShare, room)) / SIN45;
            const longest = Math.max(0, ...bars.map((b) => measureTextWidth(pathOf(b), specOf(categoryFontOf(b)))));
            const length = Math.min(longest, maxLength);
            // 斜めの名前は (中心, 下端 + 8) を右端に -45° 回す。縦の広がりは長さ × sin45 と、文字の下側の少し
            // （棒グラフと同じ見積もり：棒との間隔 8 ＋ 長さ × sin45 ＋ 文字の半分）
            categoryHeight = ROTATED_GAP + length * SIN45 + categorySize * 0.5;
            categoryTexts = bars.map((b) => ({
                lines: [truncateStartToWidth(pathOf(b), length, specOf(categoryFontOf(b)))],
                title: b.fullLabel,
            }));
            if (scrollAxis) {
                // 斜めの名前は棒の中心から左下へ伸びる。スクロールすると中身の左端より外は見えないので、はみ出す分だけ前を空ける
                leadIn = Math.max(
                    0,
                    ...categoryTexts.map((t, k) =>
                        k % step !== 0 || !t.lines.length
                            ? 0
                            : measureTextWidth(t.lines[0], specOf(categoryFontOf(bars[k]))) * SIN45 +
                              categorySize * 0.5 -
                              categoryStep * (outer + k + (1 - pad) / 2)
                    )
                );
                content += leadIn;
            }
        }
        const bottomReserve =
            nearRoom + categoryHeight + categoryTitleSize + reserve.bottom + noticeHeight + (scrollAxis ? SCROLLBAR : 0);
        plot = { x: left, y: top, width, height: H - top - bottomReserve };
    } else {
        const tickHeight = style.valueAxis.show ? tickFont.size + 8 : 4;
        // 単位のラベルは値の軸の側（下なら目盛りとタイトルの下、上ならその上）に 1 行とる（棒グラフの横棒と同じ）
        const top = reserve.top + (switchAxis ? badgeHeight : 0) + referenceLabelRoom + 4 + (switchAxis ? tickHeight + valueTitleSize : 0);
        const bottom = reserve.bottom + noticeHeight + (switchAxis ? 4 : tickHeight + valueTitleSize + badgeHeight);
        const height0 = H - top - bottom;
        const needed = minWidth > 0 ? n * minWidth : 0;
        if (needed > height0 + 1) scrollAxis = "y";
        content = scrollAxis ? needed : height0;
        const band0 = content / (n - pad + 2 * outer);
        const twoLines = hasParent && band0 >= 2 * categoryLineHeight;
        const textOf = (b: Bar) => (twoLines ? b.label : pathOf(b));
        const widest = Math.max(
            0,
            ...bars.map((b) =>
                Math.max(
                    measureTextWidth(textOf(b), specOf(categoryFontOf(b))),
                    twoLines ? measureTextWidth(b.parentLabel, specOf(categoryFontOf(b))) : 0
                )
            )
        );
        const labelRoom = style.dataLabels.show ? Math.min(valueLabelWidth + 10, W * 0.25) : 8;
        // 上限を大きくしても、描く範囲に幅の 4 分の 1 は残す（縦向きと同じ）
        const room = W * 0.75 - reserve.left - categoryTitleSize - nearRoom - reserve.right - labelRoom - (scrollAxis ? SCROLLBAR : 0);
        const categoryWidth = style.categoryAxis.show ? Math.max(0, Math.min(widest + 12, W * style.categoryAxis.maxShare, room)) : 0;
        const textWidth = Math.max(0, categoryWidth - 12);
        categoryTexts = bars.map((b) => ({
            lines: !style.categoryAxis.show
                ? []
                : twoLines
                  ? [
                        truncateToWidth(b.label, textWidth, specOf(categoryFontOf(b))),
                        ...(b.parentLabel ? [truncateToWidth(b.parentLabel, textWidth, specOf(categoryFontOf(b)))] : []),
                    ]
                  : [truncateStartToWidth(textOf(b), textWidth, specOf(categoryFontOf(b)))],
            title: b.fullLabel,
        }));
        const left = reserve.left + categoryTitleSize + Math.max(8, categoryWidth) + nearRoom;
        const right = reserve.right + labelRoom + (scrollAxis ? SCROLLBAR : 0);
        plot = { x: left, y: top, width: W - left - right, height: height0 };
        step = Math.max(1, Math.ceil(categoryLineHeight / band0));
    }
    if (plot.width < 10 || plot.height < 10) return null;

    // --- 値 → 画面の座標 --------------------------------------------------------------------------
    const span = axis.max - axis.min || 1;
    // 軸を切ったときは、切った側の端に印の余白をとる（目盛りはその内側に並ぶ）。
    // 「切った印」をなしにしたときは、標準と同じく印も余白もとらず、軸の外の分をそのまま省く
    const marked = style.breakStyle !== BREAK_STYLES.none;
    const lowZone = marked && axis.cut === "bottom" ? BREAK_ZONE : 0;
    const highZone = marked && axis.cut === "top" ? BREAK_ZONE : 0;
    const whole = horizontal ? plot.width : plot.height;
    const valueLength = whole - lowZone - highZone;
    /** 値の小さい端からの距離 → 画面の座標。範囲の反転ではひっくり返す */
    const screen = (offset: number): number => {
        const o = invert ? whole - offset : offset;
        return horizontal ? plot.x + o : plot.y + plot.height - o;
    };
    /** 値 → 値の方向の画面座標。軸の範囲の外は、余白も含めた描く範囲の端で止める（0 が軸の外の合計の棒は端まで伸びる） */
    const pos = (v: number): number =>
        screen(v < axis.min ? 0 : v > axis.max ? whole : lowZone + ((v - axis.min) / span) * valueLength);
    /** 値が増える向きの画面の符号（縦は上が負、横は右が正。反転で逆） */
    const growth = horizontal ? (invert ? -1 : 1) : invert ? 1 : -1;

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
     * （棒グラフと同じく、四角い端は「軸に乗っている」と読める）。合計は値の端だけ、浮いた増減・小計は両端
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
            // （棒グラフと同じく値の端を丸める）。終わりが 0・軸の範囲の外なら四角
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
                          y: plot.y + plot.height + nearRoom + categoryLineHeight + 2,
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
                    x: plot.x - nearRoom - 6,
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
            ...b.segments.map((s) => {
                if (!s.corner) return 0;
                const edgeOf: Record<CornerSide, number> = {
                    top: s.rect.y,
                    bottom: s.rect.y + s.rect.height,
                    left: s.rect.x,
                    right: s.rect.x + s.rect.width,
                };
                return s.corner.sides.some((side) => Math.abs(edgeOf[side] - level) < 0.5) ? cornerRadiusOf(s.rect, s.corner) : 0;
            })
        );
    for (let k = 0; k + 1 < laid.length; k++) {
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
    // 目盛りの本数は、標準と同じく描く範囲の長さで決める（縦は高さ 150・300px、横は幅 300・500px で 3・5・8 まで）
    // 「目盛りの本数 (目安)」があればその本数
    const axisTicks = axis.ticksFor
        ? axis.ticksFor(style.valueAxis.tickCount || recommendedTickCount(horizontal ? plot.width : plot.height, horizontal ? "horizontal" : "vertical"))
        : axis.ticks;
    // それでも目盛りの数字が重なるときは、間引いて出す（グリッド線も同じ目盛りだけ。0 があれば 0 を残す）
    const tickSpacing = axisTicks.length > 1 ? Math.abs(pos(axisTicks[1].value) - pos(axisTicks[0].value)) : Infinity;
    // 縦は文字の行の高さ（和文も入るフォントの上下の広がり 1.4 文字）ぶん、横は一番長い数字 + 8px
    const tickNeed = horizontal ? tickLabelWidth + 8 : tickFont.size * 1.4;
    const tickStride = tickSpacing > 0 ? Math.max(1, Math.ceil(tickNeed / tickSpacing)) : 1;
    const zeroTick = axisTicks.findIndex((t) => t.value === 0);
    const tickPhase = zeroTick >= 0 ? zeroTick % tickStride : 0;
    const shownTicks = axisTicks.filter((_, i) => i % tickStride === tickPhase);
    const ticks: TextLayout[] = style.valueAxis.show
        ? shownTicks.map(
              (t): TextLayout =>
                  horizontal
                      ? {
                            x: pos(t.value),
                            y: switchAxis ? plot.y - 6 : plot.y + plot.height + tickFont.size + 4,
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
        // 定数線は軸を広げない。範囲の外なら描かない（目標は軸に入れてある）
        if (ref.value < axis.min || ref.value > axis.max) return null;
        const p = pos(ref.value);
        const line = horizontal ? { x1: p, y1: plot.y, x2: p, y2: contentEnd } : { x1: plot.x, y1: p, x2: contentEnd, y2: p };
        return { line, color: ref.color, width: ref.width, dash: ref.dash, label: null, tooltip: ref.tooltip };
    };
    const target = referenceOf(viewModel.target);
    const constantLine = referenceOf(viewModel.constantLine);

    // --- 軸のタイトル ---------------------------------------------------------------------------------
    // タイトルは沿う辺（描く範囲の長さ）に収まるように後ろを「…」で省く。単位のラベル（描く範囲の上）に届かないように
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
            // 左右どちらも -90°（下から上へ読む。棒グラフと同じ）。-90° の文字は、ベースラインから左へ文字の上側（0.8 文字）、
            // 右へ下側（0.22 文字）が出る。右に置くときは下側が右端に届くように置く
            // （1.3.0.0 までは右を 90° にし、文字の上側が右へ出るのにベースラインを 0.3 文字しか内側にせず、右端で切れていた）
            const x = switchAxis ? W - reserve.right - 4 - valueTitleFont.size * 0.22 : reserve.left + 4 + valueTitleFont.size * 0.8;
            valueTitleLayout = titleText(valueTitle.text, valueTitleFont, x, plot.y + plot.height / 2, -90, plot.height);
        } else {
            const y = switchAxis
                ? plot.y - (style.valueAxis.show ? tickFont.size + 8 : 4) - 4
                : H - reserve.bottom - noticeHeight - 4 - badgeHeight;
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
                  plot.y + plot.height + nearRoom + categoryHeight + 2 + categoryTitleFont.size * 0.8,
                  0,
                  plot.width
              )
            : titleText(
                  categoryTitle.text,
                  categoryTitleFont,
                  reserve.left + 4 + categoryTitleFont.size * 0.8,
                  plot.y + plot.height / 2,
                  -90,
                  plot.height
              );
    }

    // --- 凡例 -------------------------------------------------------------------------------------------
    let legend: LegendLayout | null = null;
    /**
     * 凡例の 1 行（上端 top）の、色の四角の上端と文字のベースライン。文字が大きくても行の上端より上に出ないように
     * ベースラインを文字の高さに合わせて下げ、四角は文字の真ん中にそろえる（1.3.0.0 までは文字の大きさによらず上端 + 9）
     */
    const rowOf = (top: number) => {
        const textY = top + Math.max(9, legendFont.size * 0.8 + 1);
        return { y: textY - legendFont.size * 0.35 - LEGEND_SWATCH / 2, textY };
    };
    if (legendShown) {
        const items: LegendLayout["items"] = [];
        let title: TextLayout | null = null;
        const titleFont: RenderFont = { ...legendFont, weight: "bold" };
        const titleLayout = (x: number, y: number): TextLayout => ({
            x,
            y: y + legendFont.size * 0.85,
            anchor: "start",
            lines: [style.legend.title],
            rotate: 0,
            font: titleFont,
            lineHeight: legendFont.size + 2,
        });
        if (placement.side === "top" || placement.side === "bottom") {
            const y = placement.side === "top" ? TOP_MARGIN + 4 : H - noticeHeight - legendRow + 4;
            const available = W - 8;
            const widths = viewModel.legend.map(itemWidth);
            let used = legendTitleWidth;
            let count = 0;
            for (const w of widths) {
                if (used + w > available && count > 0) break;
                used += w;
                count++;
            }
            const startX = placement.align === "center" ? (W - used) / 2 : placement.align === "end" ? W - 4 - used : 4;
            if (style.legend.title) title = titleLayout(startX, y);
            let x = startX + legendTitleWidth;
            for (let i = 0; i < count; i++) {
                const dropped = i === count - 1 && count < viewModel.legend.length;
                items.push({
                    item: viewModel.legend[i],
                    index: i,
                    x,
                    swatchWidth: swatchWidthOf(viewModel.legend[i]),
                    ...rowOf(y),
                    text: dropped ? `${viewModel.legend[i].name} …` : viewModel.legend[i].name,
                });
                x += widths[i];
            }
        } else {
            const x = placement.side === "left" ? 4 : W - legendColumnWidth + 4;
            const rows = viewModel.legend.length + (style.legend.title ? 1 : 0);
            const areaTop = TOP_MARGIN + 4;
            const areaHeight = H - 8 - TOP_MARGIN - noticeHeight;
            let y =
                placement.align === "center"
                    ? areaTop + (areaHeight - rows * legendRow) / 2
                    : placement.align === "end"
                      ? areaTop + areaHeight - rows * legendRow
                      : areaTop;
            if (style.legend.title) {
                title = titleLayout(x, y);
                y += legendRow;
            }
            viewModel.legend.forEach((item, i) => {
                if (y + legendRow > areaTop + areaHeight + 1) return;
                const swatchWidth = swatchWidthOf(item);
                const textWidth = legendColumnWidth - swatchWidth - 16;
                items.push({ item, index: i, x, swatchWidth, ...rowOf(y), text: truncateToWidth(item.name, textWidth, specOf(legendFont)) });
                y += legendRow;
            });
        }
        legend = { title, items, font: legendFont };
    }

    // 単位のラベル：縦向きは値の軸の目盛りの列にそろえて、描く範囲の上に置く（棒グラフと同じ）。
    // 軸のタイトルは描く範囲の高さに収めてあるので、上の単位とは重ならない。
    // 横向きは値の軸の端（右）にそろえ、値の軸が下なら目盛りとタイトルの下、上ならその上に置く（棒グラフの横棒と同じ。1.11 までは左上）
    let badge: TextLayout | null = null;
    if (badgeHeight) {
        const badgeWidth = measureTextWidth(viewModel.unitBadge, specOf(tickFont));
        const [x, anchor]: [number, TextLayout["anchor"]] = horizontal
            ? // 単位が長くても左の凡例・ビューの外に出ないように、左端で止める
              [Math.max(Math.min(plot.x + plot.width, W - reserve.right - 2), reserve.left + 2 + badgeWidth), "end"]
            : switchAxis
              ? [Math.min(plot.x + plot.width + 6, W - reserve.right - 2 - badgeWidth), "start"]
              : [Math.max(plot.x - 6, reserve.left + 2 + badgeWidth), "end"];
        badge = {
            x,
            // 下に置くときは、文字の下側（0.22 文字）までがとった行に収まるように（文字が大きくても凡例・知らせに掛けない）
            y: horizontal && !switchAxis ? H - reserve.bottom - noticeHeight - 1 - tickFont.size * 0.22 : reserve.top + badgeRowFont,
            anchor,
            lines: [viewModel.unitBadge],
            rotate: 0,
            font: tickFont,
            lineHeight: tickFont.size + 2,
        };
    }
    // ドリルダウンした位置。単位のラベルの行に並べるときは描く範囲の左端から（単位は目盛りの列の上、軸を右に移せば右）、
    // 自分の行をとるときはその左端から。入りきらなければ末尾を「…」で省く（ツールヒントに全部）
    // 単位のラベルは、目盛りの列より長いと描く範囲の左端を越えて伸びる（軸を右に移せば、右端の外から左へ）。その端から 8px 空ける
    const sharedDrillPathSpot = () => {
        const badgeBox = badge ? boxOf(badge, 0) : null;
        if (switchAxis) {
            const right = badgeBox ? Math.min(plot.x + plot.width, badgeBox.x - 8) : plot.x + plot.width;
            return { x: plot.x, y: reserve.top + badgeRowFont, width: right - plot.x };
        }
        const x = badgeBox ? Math.max(plot.x, badgeBox.x + badgeBox.width + 8) : plot.x;
        return { x, y: reserve.top + badgeRowFont, width: W - reserve.right - x };
    };
    const drillPathSpot: { x: number; y: number; width: number } | null = !viewModel.drillPath
        ? null
        : pathSharesBadgeRow
          ? sharedDrillPathSpot()
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
        // 単位のラベルの行に並べたドリルの位置（自分の行をとるときは、下の上の余白に含まれる）
        ...(drillPath ? [boxOf(drillPath, 0)] : []),
        // 凡例の場所（上は凡例が無くても題名の下の余白）
        { x: 0, y: 0, width: W, height: reserve.top },
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

    // データ ラベルの候補：外の先（増えた棒・正の合計は値の大きい側、減った棒・負の合計は小さい側）→ 中の先 → 外の根元
    const labelFill = style.dataLabels.background;
    const labelPad = labelFill ? LABEL_PADDING : HIT_PAD;
    /** 背景の四角が文字の外に出る分 */
    const fillPad = labelFill ? LABEL_PADDING : 0;
    type Spot = "outEnd" | "inEnd" | "outStart" | "inHigh" | "inCenter" | "inLow";
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
                  : ["outEnd", "inEnd", "outStart"];
    const labelRequestOf = (b: BarLayout): LabelRequest<TextLayout> => {
        const { bar, rect } = b;
        const lines = horizontal && !verticalText ? [labelTextOf(bar)] : [bar.labelText, ...(bar.rateText ? [bar.rateText] : [])];
        const baseFont = labelFontOf(bar);
        const lineHeight = baseFont.size + 2;
        const textLength = Math.max(...lines.map((l) => measureTextWidth(l, specOf(baseFont))));
        const textThickness = baseFont.size + (lines.length - 1) * lineHeight;
        // 文字が値の方向に占める長さ
        const alongValue = horizontal === verticalText ? textThickness : textLength;
        const [low, high] = extentOf(bar);
        const highEnd = pos(high);
        const lowEnd = pos(low);
        // 棒の先（増えた棒・正の合計は値の大きい側、減った棒・負の合計は小さい側）と、そこから外へ向かう画面の向き
        const positive = bar.value >= 0;
        const endAt = positive ? highEnd : lowEnd;
        const startAt = positive ? lowEnd : highEnd;
        const outward = positive ? growth : -growth;
        const centerOf = (spot: Spot): number => {
            const half = alongValue / 2;
            switch (spot) {
                case "outEnd":
                    return endAt + outward * (half + 4);
                case "outStart":
                    return startAt - outward * (half + 4);
                // 棒の中は、端から 2px（背景を出すなら背景の四角が端から 2px）空ける
                case "inEnd":
                    return endAt - outward * (half + 2 + fillPad);
                case "inHigh":
                    return highEnd - growth * (half + 2 + fillPad);
                case "inLow":
                    return lowEnd + growth * (half + 2 + fillPad);
                default:
                    return (highEnd + lowEnd) / 2;
            }
        };
        const across = horizontal ? rect.y + rect.height / 2 : rect.x + rect.width / 2;
        // 比べる形の小計は、見せ方によって棒の中の色が違う（透かす・枠だけ・斜線・点線の枠）。その色で文字色を決める
        const rawColor = bar.segments[0]?.color ?? bar.color;
        const segmentColor = bar.kind === BAR_KINDS.subtotal ? subtotalPaint(style, rawColor).under : rawColor;
        const explicit = isTotalBar(bar) ? style.dataLabels.totalFont.color : style.dataLabels.font.color;
        const candidates = spots.map((spot): Candidate<TextLayout> => {
            const inside = spot !== "outEnd" && spot !== "outStart";
            const at = centerOf(spot);
            const cx = horizontal ? at : across;
            const cy = horizontal ? across : at;
            // 自動の文字色は、文字のすぐ下の色で決める：背景を出すなら背景（透過性があれば下の色と混ぜた色）、
            // 出さないなら置いた場所の色（棒の中は棒の色、外はビジュアルの背景）。
            // 1.7.0.0 までは背景を出していても棒の中なら棒の色で決め、白い背景に白い文字を描いていた（第 32 ページ ③）
            const underneath = inside ? segmentColor : style.background;
            const base = labelFill ? blend(labelFill.color, underneath, 1 - labelFill.opacity) : underneath;
            // 棒の中は白を、外は灰色を優先する（標準と同じ。背景の既定は黒・透過性 90% なので、棒の中は白、外は灰色になる）
            const color =
                explicit ||
                (inside
                    ? readableText(base, "#FFFFFF", INSIDE_TEXT_CONTRAST)
                    : readableText(base, LABEL_OUTSIDE_COLOR, OUTSIDE_TEXT_CONTRAST));
            const font = { ...baseFont, color };
            // 真ん中 (cx, cy) に置く。複数行は 1 行目のベースラインを上へずらす
            const firstBaseline = font.size * 0.35 - ((lines.length - 1) * lineHeight) / 2;
            const text: TextLayout = verticalText
                ? { x: cx + firstBaseline, y: cy, anchor: "middle", lines, rotate: -90, font, lineHeight }
                : { x: cx, y: cy + firstBaseline, anchor: "middle", lines, rotate: 0, font, lineHeight };
            text.background = backgroundOf(lines, "middle", font, lineHeight, labelFill);
            // 棒の中の候補は、自分の棒の中に収まるときだけ
            return { item: text, box: boxOf(text, labelPad), within: inside ? rect : undefined };
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
    const pairs: Array<[ViewModel["target"], ReferenceLine | null, string]> = [
        [viewModel.constantLine, constantLine, "ref:constant"],
        [viewModel.target, target, "ref:target"],
    ];
    for (const [source, ref, owner] of pairs) {
        const request = ref ? nameRequestOf(source, ref, owner) : null;
        if (ref && request) nameRequests.push({ ref, request });
    }

    if (!scrollAxis) {
        // スクロールしないときは、ラベルと線の名前を 1 度に置く（名前はラベルのあと）
        const placed = placeLabels([...labelRequests, ...nameRequests.map((n) => n.request)], contentObstacles, { bounds: contentBox });
        labelled.forEach((b, i) => (b.label = placed[i]?.item ?? null));
        nameRequests.forEach((n, i) => (n.ref.label = placed[labelRequests.length + i]?.item ?? null));
    } else {
        // スクロールするときは、ラベルは中身の座標で置く。線の名前は動かない側（ビューの座標）に描くので、
        // 動かない物（目盛り・凡例など）と項目名の側、もう 1 本の名前だけを避ける（棒とラベルはスクロールで動く）
        const placed = placeLabels(labelRequests, contentObstacles, { bounds: contentBox });
        labelled.forEach((b, i) => (b.label = placed[i]?.item ?? null));
        const names = placeLabels(
            nameRequests.map((n) => n.request),
            [...overlayBoxes.map(asBox), nearBand],
            { bounds: { x: 0, y: 0, width: W, height: H } }
        );
        nameRequests.forEach((n, i) => (n.ref.label = names[i]?.item ?? null));
    }

    // 好ましい置き場所（最初の候補）が項目名の側へはみ出して置けなかったラベルがあれば、はみ出した分を返す。
    // layoutOf が描く範囲と項目名のあいだを空けて置き直す
    let nearOverflow = 0;
    labelRequests.forEach((request, i) => {
        const first = request.candidates[0];
        if (!first || labelled[i].label === first.item) return;
        const over = horizontal ? nearEdge - first.box.x : first.box.y + first.box.height - nearEdge;
        if (over > 1) nearOverflow = Math.max(nearOverflow, over);
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
    return { layout, nearOverflow };
}
