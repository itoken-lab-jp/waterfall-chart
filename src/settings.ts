"use strict";

/**
 * 書式ペイン。カードの name とスライスの name は capabilities.json の objects と完全に一致させる
 * （ずれてもエラーは出ず、書式ペインの変更が値に反映されないだけになる。format-pane Skill）。
 *
 * データ次第で中身が変わるもの（イベントの選択肢、系列の色、形で意味の無い設定）は、
 * update() のたびに viewModel の結果から流し込む（applyData）。
 */

import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import SimpleCard = formattingSettings.SimpleCard;
import CompositeCard = formattingSettings.CompositeCard;
import Group = formattingSettings.Group;
import Model = formattingSettings.Model;

import { UNIT_TYPES, UNIT_NOTATIONS, PRECISIONS } from "./shared/units";
import { LEGEND_POSITIONS, LEGEND_POSITION_ITEMS, LegendPosition, standardLegendPosition, legendPlacementValue } from "./shared/legend";
import { AutoNumUpDown, itemOf, dropdownValue, TITLE_STYLES, TITLE_STYLE_ITEMS, LABEL_ORIENTATIONS, LABEL_ORIENTATION_ITEMS } from "./shared/formatting";
import { AXIS_SCROLL_START_ITEMS } from "./shared/scrollStart";
import { NEGATIVE_STYLE_ITEMS, ZERO_STYLE_ITEMS, SIGN_TONE_MODE_ITEMS, DEFAULT_GOOD_COLOR, DEFAULT_BAD_COLOR } from "./shared/numberFormat";
import {
    applyNumberRanges,
    FONT_SIZE_RANGE,
    legendParts,
    gridlineParts,
    categoryAxisParts,
    valueAxisParts,
    labelValueParts,
    labelBackgroundParts,
    labelHaloParts,
    fontControl as sharedFontControl,
    NumberRange,
    DEFAULT_FONT_FAMILY,
} from "./shared/formatCards";
import { applyThemeBasics } from "./shared/theme";
import { referenceLineParts, syncReferenceLineVisibility, VALUE_SHADE_REGION_ITEMS } from "./shared/referenceLine";
import { LINE_STYLES, LINE_STYLE_ITEMS } from "./shared/gridlines";

export { UNIT_TYPES, UNIT_NOTATIONS, PRECISIONS };

/** 向き。横は項目名が長いとき */
export const ORIENTATIONS = { vertical: "vertical", horizontal: "horizontal" } as const;
export type Orientation = (typeof ORIENTATIONS)[keyof typeof ORIENTATIONS];

export const ORIENTATION_ITEMS: powerbi.IEnumMember[] = [
    { value: ORIENTATIONS.vertical, displayName: "縦" },
    { value: ORIENTATIONS.horizontal, displayName: "横" },
];

/** イベントのつなぎ方（形 1）。ends = 両端だけ、all = 左端から右端までのイベントを順に */
export const CONNECT_MODES = { ends: "ends", all: "all" } as const;
export type ConnectMode = (typeof CONNECT_MODES)[keyof typeof CONNECT_MODES];

export const CONNECT_MODE_ITEMS: powerbi.IEnumMember[] = [
    { value: CONNECT_MODES.ends, displayName: "両端だけ" },
    { value: CONNECT_MODES.all, displayName: "すべてを順に" },
];

/** 増減の棒の並び。小計があるときは、小計の区切りの中で並べ替える */
export const ORDERS = { data: "data", delta: "delta", absDelta: "absDelta" } as const;
export type Order = (typeof ORDERS)[keyof typeof ORDERS];

export const ORDER_ITEMS: powerbi.IEnumMember[] = [
    { value: ORDERS.data, displayName: "データの順" },
    { value: ORDERS.delta, displayName: "増減の大きい順" },
    { value: ORDERS.absDelta, displayName: "増減の絶対値の大きい順" },
];

/**
 * 起点の取り方（形 3・5）。起点のメジャーは項目ごとに評価されて届く。
 * first = 最初の項目の値（期首の残高のように、どの項目でも同じ値になるもの・月ごとの期首）、
 * sum = 項目の合計（事業ごとに期首を持つモデルを事業で分けて出すとき）
 */
export const ORIGIN_MODES = { first: "first", sum: "sum" } as const;
export type OriginMode = (typeof ORIGIN_MODES)[keyof typeof ORIGIN_MODES];

export const ORIGIN_MODE_ITEMS: powerbi.IEnumMember[] = [
    { value: ORIGIN_MODES.first, displayName: "最初の項目の値" },
    { value: ORIGIN_MODES.sum, displayName: "項目の合計" },
];

/** 率。伸び率 = 差 ÷ その項目の左端の値、寄与率 = 差 ÷ 左端の合計、構成比 = 値 ÷ 最後の合計（形 2） */
export const RATE_TYPES = { none: "none", growth: "growth", contribution: "contribution", share: "share" } as const;
export type RateType = (typeof RATE_TYPES)[keyof typeof RATE_TYPES];

export const RATE_TYPE_ITEMS: powerbi.IEnumMember[] = [
    { value: RATE_TYPES.none, displayName: "なし" },
    { value: RATE_TYPES.growth, displayName: "伸び率（差 ÷ その項目の左端の値）" },
    { value: RATE_TYPES.contribution, displayName: "寄与率（差 ÷ 左端の合計）" },
    { value: RATE_TYPES.share, displayName: "構成比（値 ÷ 最後の合計）" },
];

export { LINE_STYLES, LINE_STYLE_ITEMS };
export type LineStyle = (typeof LINE_STYLES)[keyof typeof LINE_STYLES];

/** 軸を切った印。slash = 斜線（既定）、wave = 波線、none = 描かない（標準と同じく省くだけ） */
export const BREAK_STYLES = { slash: "slash", wave: "wave", none: "none" } as const;
export type BreakStyle = (typeof BREAK_STYLES)[keyof typeof BREAK_STYLES];

export const BREAK_STYLE_ITEMS: powerbi.IEnumMember[] = [
    { value: BREAK_STYLES.slash, displayName: "斜線" },
    { value: BREAK_STYLES.wave, displayName: "波線" },
    { value: BREAK_STYLES.none, displayName: "なし" },
];

/**
 * 比べる形の小計（計画からその段階までの差）の見せ方。色はどれも差の色（増加・減少）。
 * ふつうの増減の棒と見分けられるように、既定は枠だけ（2026-09-24 ユーザー決定。1.11 までは透かす形だけ）
 */
export const SUBTOTAL_STYLES = { outline: "outline", fill: "fill", hatch: "hatch", dashed: "dashed" } as const;
export type SubtotalStyle = (typeof SUBTOTAL_STYLES)[keyof typeof SUBTOTAL_STYLES];

export const SUBTOTAL_STYLE_ITEMS: powerbi.IEnumMember[] = [
    { value: SUBTOTAL_STYLES.outline, displayName: "枠だけ" },
    { value: SUBTOTAL_STYLES.fill, displayName: "透かす" },
    { value: SUBTOTAL_STYLES.hatch, displayName: "斜線" },
    { value: SUBTOTAL_STYLES.dashed, displayName: "点線の枠" },
];

export { LEGEND_POSITIONS, LEGEND_POSITION_ITEMS, standardLegendPosition, legendPlacementValue };
export type { LegendPosition };

/** データ ラベルの位置。自動と外側の上は、増えた棒の先（減った棒は下・左） */
export const LABEL_POSITIONS = {
    auto: "auto",
    outsideEnd: "outsideEnd",
    insideTop: "insideTop",
    insideCenter: "insideCenter",
    insideBottom: "insideBottom",
} as const;

export const LABEL_POSITION_ITEMS: powerbi.IEnumMember[] = [
    // 名前と並びは標準のウォーターフォールと同じ
    { value: LABEL_POSITIONS.auto, displayName: "自動" },
    { value: LABEL_POSITIONS.insideTop, displayName: "内側上" },
    { value: LABEL_POSITIONS.outsideEnd, displayName: "外側上" },
    { value: LABEL_POSITIONS.insideCenter, displayName: "内側中央" },
    { value: LABEL_POSITIONS.insideBottom, displayName: "内側下" },
];

/** データ ラベルの表示単位。auto は Y 軸と同じ */
export const LABEL_UNIT_TYPES: powerbi.IEnumMember[] = [{ value: "auto", displayName: "Y 軸と同じ" }, ...UNIT_TYPES.slice(1)];

/**
 * 既定のフォント（Segoe UI）。書式ペインのフォントの欄に出る名前。
 * 1.15 までは代わりのフォントまで並べた指定（"Segoe UI", wf_segoe-ui_normal, …）で、保存した値はそのまま読む
 */
export { DEFAULT_FONT_FAMILY };

/**
 * 軸のタイトル・凡例の文字の既定。標準のウォーターフォールと同じ DIN 12・10（2026-09-23 に Desktop で並べて確認）。
 * 1.8.0.0 までは 9・8 で、標準より小さかった
 */
export const AXIS_TITLE_FONT_FAMILY = "DIN";
export const AXIS_TITLE_FONT_SIZE = 12;
export const LEGEND_FONT_SIZE = 10;

/**
 * 比較の列の値（イベント）がまだ届いていないときの選択肢。update() のたびに applyData で差し替える。
 * 書式では「比較」「イベント」という言葉を使わない（2026-09-19 のユーザーの決め）
 */
const NO_EVENT: powerbi.IEnumMember = { value: "", displayName: "（なし）" };

/** Y 軸の単位ラベルの位置。valueAxisTop = 値の軸の目盛りの列の端（縦向きは上、横向きは右下）、plotTopRight = 描く範囲の右上 */
export const UNIT_POSITIONS: powerbi.IEnumMember[] = [
    { value: "valueAxisTop", displayName: "Y 軸の上 (左上)" },
    { value: "plotTopRight", displayName: "プロット エリアの右上" },
    { value: "none", displayName: "非表示" },
];

/** 単位ラベルの書き方。表示される形そのものを選択肢の名前にする */
export const UNIT_STYLES: powerbi.IEnumMember[] = [
    { value: "parentheses", displayName: "(百万円)" },
    { value: "withPrefix", displayName: "(単位: 百万円)" },
];

/** データ ラベルの背景の既定（標準のウォーターフォールと同じ黒・透過性 90%） */
export const DEFAULT_LABEL_BACKGROUND = "#000000";
export const DEFAULT_LABEL_BACKGROUND_TRANSPARENCY = 90;

/** 「最後の合計」の名前の既定 */
export const DEFAULT_TOTAL_LABEL = "合計";
/** 「その他」の名前の既定 */
export const DEFAULT_OTHERS_LABEL = "その他";
/** 「その他」で残す数の既定。標準のウォーターフォールの「最大の内訳」と同じ 5 */
export const DEFAULT_OTHERS_COUNT = 5;
/** 系列の色を書式ペインに並べる上限。多すぎると書式ペインが重くなる */
export const MAX_SERIES_TARGETS = 30;

export { AutoNumUpDown, dropdownValue, TITLE_STYLES, TITLE_STYLE_ITEMS, LABEL_ORIENTATIONS, LABEL_ORIENTATION_ITEMS };

/**
 * フォント（種類・サイズ・B/I/U）。子のスライスの name が capabilities のプロパティになる。
 * prefix を付けると titleFontFamily のように前に付く（同じカードに 2 つ置くとき）
 */
function fontControl(prefix: string, displayName: string, size: number, bold = false, family = DEFAULT_FONT_FAMILY): formattingSettings.FontControl {
    return sharedFontControl({ controlName: `${prefix || "font"}Control`, prefix, displayName, size, bold, family });
}

/**
 * 棒の組み立て。イベントの選択肢はデータ次第なので、update() のたびに applyEvents で組み直す
 * （保存するのはイベントの名前）。
 */
export class LayoutCardSettings extends SimpleCard {
    name = "layout";
    displayName = "ウォーターフォール";

    orientation = new formattingSettings.ItemDropdown({
        name: "orientation",
        displayName: "向き",
        items: ORIENTATION_ITEMS,
        value: ORIENTATION_ITEMS[0],
    });

    connectMode = new formattingSettings.ItemDropdown({
        name: "connectMode",
        displayName: "つなぎ方",
        items: CONNECT_MODE_ITEMS,
        value: CONNECT_MODE_ITEMS[0],
    });

    leftEvent = new formattingSettings.ItemDropdown({
        name: "leftEvent",
        displayName: "左端",
        items: [NO_EVENT],
        value: NO_EVENT,
    });

    rightEvent = new formattingSettings.ItemDropdown({
        name: "rightEvent",
        displayName: "右端",
        items: [NO_EVENT],
        value: NO_EVENT,
    });

    order = new formattingSettings.ItemDropdown({
        name: "order",
        displayName: "並び",
        items: ORDER_ITEMS,
        value: ORDER_ITEMS[0],
    });

    originMode = new formattingSettings.ItemDropdown({
        name: "originMode",
        displayName: "起点の取り方",
        description: "起点は項目ごとに評価されて届く。どの項目でも同じ値（期首の残高など）なら最初の項目の値、項目ごとに期首を持つなら項目の合計",
        items: ORIGIN_MODE_ITEMS,
        value: ORIGIN_MODE_ITEMS[0],
    });

    /** 比べる形で、左端の合計の前に左端の値を項目ごとに積む（売上の科目 → 売上 → 差。2026-10-13 ユーザー）。標準に無い項目 */
    leftBreakdown = new formattingSettings.ToggleSwitch({
        name: "leftBreakdown",
        displayName: "左端の内訳を積む",
        description: "左端の合計の前に、左端の値を項目ごとに 0 から積み上げる（売上の科目を積んでから売上の合計へ、など）。値が無い項目は積まない",
        value: false,
    });

    totalShow = new formattingSettings.ToggleSwitch({
        name: "totalShow",
        displayName: "最後の合計",
        value: true,
    });

    totalLabel = new formattingSettings.TextInput({
        name: "totalLabel",
        displayName: "最後の合計の名前",
        value: "",
        placeholder: DEFAULT_TOTAL_LABEL,
    });

    /** 増減が 0 の項目の棒と名前を出さない（凡例で積んで打ち消し合う項目は残す）。標準に無い項目 */
    hideZero = new formattingSettings.ToggleSwitch({
        name: "hideZero",
        displayName: "0 の項目を隠す",
        description: "増減が 0 の項目の棒と名前を出さない。凡例の値ごとの増減が打ち消し合って 0 になる項目は残す。左端・右端・小計の値は隠した項目も含めて出す",
        value: false,
    });

    /** ドリルダウンしたとき、今いる位置（事業A ＞ 製品A1 など）を左上に出す（2026-09-25 ユーザー）。標準に無い項目 */
    drillPathShow = new formattingSettings.ToggleSwitch({
        name: "drillPathShow",
        displayName: "ドリルの位置",
        description: "ドリルダウンや絞り込みで、表示している項目の上の階層が 1 つに決まるとき、その位置（事業A ＞ 製品A1 など）を左上に出す",
        value: true,
    });

    /** マウスを乗せたとき、右下に「画像としてコピー」のボタンを出す。標準に無い項目 */
    copyButton = new formattingSettings.ToggleSwitch({
        name: "copyButton",
        displayName: "画像のコピー",
        description: "マウスを乗せたとき、右下に「画像としてコピー」のボタンを出す。押すと、見えているグラフを画像としてクリップボードに入れ、PowerPoint などに貼れる",
        value: true,
    });

    slices = [
        this.orientation,
        this.connectMode,
        this.leftEvent,
        this.rightEvent,
        this.leftBreakdown,
        this.order,
        this.originMode,
        this.totalShow,
        this.totalLabel,
        this.hideZero,
        this.drillPathShow,
        this.copyButton,
    ];

    /**
     * イベントの選択肢を流し込み、形で意味の無い設定を隠す。
     * events が 2 つ未満（比べる形でない）ときは、イベントの設定を隠す。
     * left / right は viewModel が生の objects から読んで解決した名前（保存値が今のデータに無ければ既定の最初と最後）。
     * 形 1 では右端のイベントが最後の合計の代わりになるので、最後の合計の設定を隠す。起点の取り方は起点があるときだけ。
     * つなぎ方はイベントが 3 つ以上のときだけ（2 つなら両端だけとすべてを順にが同じ）
     */
    applyEvents(events: string[], left: string, right: string, hasOrigin = false): void {
        const compare = events.length >= 2;
        // 比べる段が 2 つなら、両端だけとすべてを順には同じ棒になる
        this.connectMode.visible = events.length > 2;
        this.leftEvent.visible = compare;
        this.rightEvent.visible = compare;
        this.leftBreakdown.visible = compare;
        this.totalShow.visible = !compare;
        // 名前は最後の合計を出すときだけ使う
        this.totalLabel.visible = !compare && (this.totalShow.value ?? true);
        this.originMode.visible = hasOrigin;
        if (!compare) return;
        const items = events.map((name) => ({ value: name, displayName: name }));
        this.leftEvent.items = items;
        this.rightEvent.items = items;
        this.leftEvent.value = items.find((item) => item.value === left) ?? items[0];
        this.rightEvent.value = items.find((item) => item.value === right) ?? items[items.length - 1];
    }
}

/**
 * 小さい増減を 1 本にまとめる。既定はオフ。カードの名前と数の項目は標準と同じ「詳細」「最大の内訳」
 * （ユーザー、2026-09-20）。保存先は others・count のまま
 */
export class OthersCardSettings extends SimpleCard {
    name = "others";
    displayName = "詳細";

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "その他にまとめる",
        value: false,
    });

    topLevelSlice = this.show;

    /** 欄の範囲は NUMBER_RANGES（描くときも viewModel で同じ範囲に丸める） */
    count = new formattingSettings.NumUpDown({
        name: "count",
        displayName: "最大の内訳",
        description: "区切りごとに、増減の絶対値の大きいものをこの数だけ残し、残りを 1 本にまとめる",
        value: DEFAULT_OTHERS_COUNT,
    });

    label = new formattingSettings.TextInput({
        name: "label",
        displayName: "まとめた棒の名前",
        value: "",
        placeholder: DEFAULT_OTHERS_LABEL,
    });

    slices = [this.count, this.label];
}

/** 凡例（系列）の色の「設定の適用先」の 1 項目 */
class SeriesColorItem extends SimpleCard {
    name = "seriesColorsTarget";

    constructor(displayName: string, slices: formattingSettings.Slice[]) {
        super();
        this.displayName = displayName;
        this.slices = slices;
    }
}

export interface SeriesColorTarget {
    name: string;
    color: string;
    selector: powerbi.data.Selector;
}


/**
 * 列（標準のウォーターフォールの「列」カードと同じ形）。色は空のままならテーマの色を使う
 * （viewModel が解決した色を書き戻して、ペインにも出す）
 */
export class ColumnsCardSettings extends CompositeCard {
    name = "columns";
    // 標準の日本語の表示は「列」だが、縦向き・横向きのどちらでも読めるよう「棒」にする（保存先は columns のまま）
    displayName = "棒";

    increaseFill = new formattingSettings.ColorPicker({
        name: "increaseFill",
        displayName: "増加",
        value: { value: "" },
    });

    decreaseFill = new formattingSettings.ColorPicker({
        name: "decreaseFill",
        displayName: "減少",
        value: { value: "" },
    });

    totalFill = new formattingSettings.ColorPicker({
        name: "totalFill",
        displayName: "合計",
        description: "右端・小計・最後の合計の棒（起点の色を決めなければ、左端の合計も）",
        value: { value: "" },
    });

    /**
     * 左端の合計（比べる形の左端のイベント、期首などの起点）だけの色。空なら合計の色（既存のレポートの見た目を変えない）。
     * 前年・計画を灰、今年・実績を濃い色にして両端を見分ける作りが多い（2026-09-25 ユーザー）。標準に無い項目
     */
    startFill = new formattingSettings.ColorPicker({
        name: "startFill",
        displayName: "起点",
        description: "左端の合計（計画・前年・期首など）の棒。空なら合計の色",
        value: { value: "" },
    });

    /**
     * 比べる形の小計の透過性。色は増加・減少の色（差が 0 以上なら増加、マイナスなら減少）を使い、
     * これだけ透かして描く（ユーザー、2026-09-19「小計の色はほかの上下の色と合わせて、透明度だけ落とす」）。
     * 保存先の名前は subtotalTransparency のまま。欄の範囲は NUMBER_RANGES（viewModel も 0〜100 に丸める）
     */
    subtotalTransparency = new formattingSettings.NumUpDown({
        name: "subtotalTransparency",
        displayName: "小計の透過性 (%)",
        description: "比べる形で小計を挟んだときの、浮いた小計の棒。色は増加・減少の色",
        value: 50,
    });

    /** 比べる形の小計の見せ方。透過性は「透かす」のときだけ出す */
    subtotalStyle = new formattingSettings.ItemDropdown({
        name: "subtotalStyle",
        displayName: "小計の見せ方",
        description: "比べる形で小計を挟んだときの、浮いた小計の棒。ふつうの増減の棒と見分けやすいように描き分ける",
        items: SUBTOTAL_STYLE_ITEMS,
        value: SUBTOTAL_STYLE_ITEMS[0],
    });

    othersFill = new formattingSettings.ColorPicker({
        name: "othersFill",
        displayName: "その他",
        value: { value: "" },
    });

    /** 空 = 自動（カテゴリ間のスペースの半分） */
    outerPadding = new AutoNumUpDown({
        name: "outerPadding",
        displayName: "外側のパディング (%)",
        value: undefined,
    });

    categorySpacing = new formattingSettings.NumUpDown({
        name: "categorySpacing",
        displayName: "カテゴリ間のスペース (%)",
        value: 20,
    });

    /**
     * 棒の角丸（既定 0）。棒の端を丸め、値 0 の軸に乗っている端だけ四角にする（合計は値の端だけ、浮いた増減・小計は両端）。
     * 欄の範囲は NUMBER_RANGES（viewModel も 0〜30 に丸める）
     */
    cornerRadius = new formattingSettings.NumUpDown({
        name: "cornerRadius",
        displayName: "角丸 (px)",
        value: 0,
    });

    colorGroup = new Group({
        name: "columnsColor",
        displayName: "色",
        slices: [this.increaseFill, this.decreaseFill, this.totalFill, this.startFill, this.othersFill, this.subtotalStyle, this.subtotalTransparency],
    });

    layoutGroup = new Group({
        name: "columnsLayout",
        displayName: "レイアウト",
        slices: [this.outerPadding, this.categorySpacing, this.cornerRadius],
    });

    /**
     * 凡例（系列）の値ごとの色。「設定の適用先」で選ぶ。保存先は columns の fill（selector 付き）。
     * 凡例に列が無いときは隠す
     */
    seriesGroup = new Group({
        name: "columnsSeries",
        displayName: "凡例の色",
        slices: [],
        visible: false,
    });

    groups = [this.colorGroup, this.seriesGroup, this.layoutGroup];

    applySeries(targets: SeriesColorTarget[]): void {
        this.seriesGroup.visible = targets.length > 0;
        this.seriesGroup.container = targets.length
            ? new formattingSettings.Container({
                  displayName: "設定の適用先",
                  containerItems: targets.slice(0, MAX_SERIES_TARGETS).map(
                      (target) =>
                          new SeriesColorItem(target.name, [
                              new formattingSettings.ColorPicker({
                                  name: "fill",
                                  displayName: "カラー",
                                  value: { value: target.color },
                                  selector: target.selector,
                              }),
                          ])
                  ),
              })
            : undefined;
    }
}

/**
 * 凡例。標準と同じく「増加・減少・合計（・その他）」を出す。小計は出さない。系列があるときは系列と合計
 * （系列で積むと増減の棒は系列の色になり、増加・減少の色を使わないため）
 */
export class LegendCardSettings extends CompositeCard {
    name = "legend";
    displayName = "凡例";

    // タイトルの既定はオン（標準と同じ）。テーマに値があればテーマに従う（applyThemeDefaults）
    private parts = legendParts({ showDisplayName: "凡例", titleShow: true, font: { controlName: "fontControl", displayName: "フォント", size: LEGEND_FONT_SIZE, family: DEFAULT_FONT_FAMILY } });
    show = this.parts.show;
    topLevelSlice = this.show;
    position = this.parts.position;
    /** 上下に置いた凡例の行数（左右に置いたときは出さない。syncRows） */
    rows = this.parts.rows;
    font = this.parts.font;
    labelColor = this.parts.labelColor;
    titleShow = this.parts.titleShow;
    titleText = this.parts.titleText;
    optionsGroup = this.parts.optionsGroup;
    textGroup = this.parts.textGroup;
    titleGroup = this.parts.titleGroup;

    groups = [this.optionsGroup, this.textGroup, this.titleGroup];

    /** 行数は、位置が上・下のときだけ出す */
    syncRows(): void {
        this.parts.syncRows();
    }
}

/** 棒と棒をつなぐ線（標準と同じ「接続線」。ユーザーの言葉、2026-09-20） */
export class ConnectorsCardSettings extends SimpleCard {
    name = "connectors";
    displayName = "接続線";

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "接続線",
        value: true,
    });

    topLevelSlice = this.show;

    /** 空 = テーマの補助の線の色 */
    color = new formattingSettings.ColorPicker({
        name: "color",
        displayName: "カラー",
        value: { value: "" },
    });

    /** 欄の範囲は NUMBER_RANGES（viewModel も 0.5〜10 に丸める） */
    width = new formattingSettings.NumUpDown({
        name: "width",
        displayName: "幅 (px)",
        value: 1,
    });

    lineStyle = new formattingSettings.ItemDropdown({
        name: "lineStyle",
        displayName: "線のスタイル",
        items: LINE_STYLE_ITEMS,
        value: itemOf(LINE_STYLE_ITEMS, LINE_STYLES.solid),
    });

    slices = [this.color, this.width, this.lineStyle];
}

/**
 * データ ラベル（保存先の名前：position・orientation・fontFamily・bold・italic・color・precision・
 * backgroundShow・backgroundColor・backgroundTransparency）。表示単位は Y 軸と別に持てる（unitType、既定は Y 軸と同じ）
 */
export class DataLabelsCardSettings extends CompositeCard {
    name = "dataLabels";
    displayName = "データ ラベル";

    private values = labelValueParts({
        precisions: PRECISIONS,
        negativeStyles: NEGATIVE_STYLE_ITEMS,
        zeroStyles: ZERO_STYLE_ITEMS,
        toneModes: SIGN_TONE_MODE_ITEMS,
        goodColor: DEFAULT_GOOD_COLOR,
        badColor: DEFAULT_BAD_COLOR,
    });
    private background = labelBackgroundParts({
        showDisplayName: "背景",
        colorDisplayName: "カラー",
        color: DEFAULT_LABEL_BACKGROUND,
        transparency: DEFAULT_LABEL_BACKGROUND_TRANSPARENCY,
    });

    /** 既定はオフ（標準のウォーターフォールと同じ）。保存したレポートは保存値どおり */
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "データ ラベル",
        value: false,
    });

    topLevelSlice = this.show;

    orientation = new formattingSettings.ItemDropdown({
        name: "orientation",
        displayName: "方向",
        items: LABEL_ORIENTATION_ITEMS,
        value: LABEL_ORIENTATION_ITEMS[0],
    });

    position = new formattingSettings.ItemDropdown({
        name: "position",
        displayName: "位置",
        items: LABEL_POSITION_ITEMS,
        value: LABEL_POSITION_ITEMS[0],
    });

    font = fontControl("", "フォント", 9);

    color = this.values.color;

    unitType = new formattingSettings.ItemDropdown({
        name: "unitType",
        displayName: "表示単位",
        items: LABEL_UNIT_TYPES,
        value: LABEL_UNIT_TYPES[0],
    });

    precision = this.values.precision;

    plusSign = new formattingSettings.ToggleSwitch({
        name: "plusSign",
        displayName: "増加に + を付ける",
        value: true,
    });

    negativeStyle = this.values.negativeStyle;
    zeroStyle = this.values.zeroStyle;
    negativeZero = this.values.negativeZero;
    toneMode = this.values.toneMode;
    positiveColor = this.values.positiveColor;
    negativeColor = this.values.negativeColor;

    rateType = new formattingSettings.ItemDropdown({
        name: "rateType",
        displayName: "率",
        description: "増減の棒のラベルとツールヒントに率を足す。伸び率・寄与率は比べる形、構成比は合流の形",
        items: RATE_TYPE_ITEMS,
        value: RATE_TYPE_ITEMS[0],
    });

    backgroundShow = this.background.backgroundShow;
    backgroundColor = this.background.backgroundColor;
    backgroundTransparency = this.background.backgroundTransparency;

    optionsGroup = new Group({
        name: "labelOptions",
        displayName: "オプション",
        slices: [this.orientation, this.position],
    });

    valuesGroup = new Group({
        name: "labelValues",
        displayName: "値",
        slices: [this.font, this.color, this.unitType, this.precision, this.plusSign, this.negativeStyle, this.zeroStyle, this.negativeZero, this.toneMode, this.positiveColor, this.negativeColor, this.rateType],
    });

    /** 合計・小計の棒のラベル。既定で太字にして、増減のラベルと見分ける */
    totalFont = fontControl("total", "フォント", 9, true);

    /** 空 = 値のラベルと同じ（それも空なら自動） */
    totalColor = new formattingSettings.ColorPicker({
        name: "totalColor",
        displayName: "カラー",
        value: { value: "" },
    });

    totalsGroup = new Group({
        name: "labelTotals",
        displayName: "合計のラベル",
        description: "合計・小計の棒のラベル",
        slices: [this.totalFont, this.totalColor],
    });

    backgroundGroup = new Group({
        name: "labelBackground",
        displayName: "背景",
        topLevelSlice: this.backgroundShow,
        slices: [this.backgroundColor, this.backgroundTransparency],
    });

    /** 文字のまわりに縁を付ける（標準に無い項目。隣の棒や線の上でも読める） */
    private halo = labelHaloParts();
    haloShow = this.halo.haloShow;
    haloColor = this.halo.haloColor;
    haloWidth = this.halo.haloWidth;
    haloGroup = this.halo.haloGroup;

    groups = [this.optionsGroup, this.valuesGroup, this.totalsGroup, this.backgroundGroup, this.haloGroup];
}

/** 目標の線。「目標」にフィールドがあるときだけ出す。名前は凡例に出し、線にマウスを当てるとツールヒントで名前と値を出す */
export class TargetCardSettings extends SimpleCard {
    name = "target";
    displayName = "目標";

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "目標の線",
        value: true,
    });

    topLevelSlice = this.show;

    color = new formattingSettings.ColorPicker({
        name: "color",
        displayName: "カラー",
        value: { value: "#252423" },
    });

    width = new formattingSettings.NumUpDown({
        name: "width",
        displayName: "幅 (px)",
        value: 1.5,
    });

    lineStyle = new formattingSettings.ItemDropdown({
        name: "lineStyle",
        displayName: "線のスタイル",
        items: LINE_STYLE_ITEMS,
        value: itemOf(LINE_STYLE_ITEMS, LINE_STYLES.dashed),
    });

    /** 名前は凡例に出す（印は線）。グラフの中の線の上にも名前と値を出すときにオン。既定はオフ */
    labelShow = new formattingSettings.ToggleSwitch({
        name: "labelShow",
        displayName: "グラフの中の名前と値",
        description: "名前は凡例に出す。オンにすると、グラフの中の線の上にも名前と値を出す",
        value: false,
    });

    slices = [this.color, this.width, this.lineStyle, this.labelShow];
}

/**
 * 定数線（固定の値）。標準と同じく分析ペインに置く（capabilities の objectCategory: 2 と、カードの analyticsPane）。
 * 線・網掛け領域・データ ラベルの項目は標準の Y 軸の定数線と同じ（作りは shared/referenceLine.ts）。
 * 1.15.5.16 までの保存先（value は数、名前は labelText）はそのまま読む。ラベルのスタイルは labelStyle に保存する。
 * 軸は広げない（標準と同じ。範囲の外なら描かない）
 */
export class ConstantLineCardSettings extends CompositeCard {
    name = "constantLine";
    displayName = "定数線";
    analyticsPane = true;

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "定数線",
        value: false,
    });

    topLevelSlice = this.show;

    value = new formattingSettings.NumUpDown({
        name: "value",
        displayName: "値",
        value: 0,
    });

    /** 線の名前（データ ラベルのスタイルが名前・両方のとき）。保存先の名前は前のまま */
    labelText = new formattingSettings.TextInput({
        name: "labelText",
        displayName: "名前",
        value: "",
        placeholder: "定数線",
    });

    /** 線・網掛け領域・データ ラベルの部品（標準の Y 軸の定数線と同じく、ラベルの既定は左） */
    readonly parts = referenceLineParts({ labelHorizontal: "left", labelTextName: "labelStyle" });
    color = this.parts.color;
    transparency = this.parts.transparency;
    lineStyle = this.parts.lineStyle;
    dashArray = this.parts.dashArray;
    scaleWithWidth = this.parts.scaleWithWidth;
    dashCap = this.parts.dashCap;
    width = this.parts.width;
    layer = this.parts.layer;
    shadeShow = this.parts.shadeShow;

    shadeRegion = new formattingSettings.ItemDropdown({
        name: "shadeRegion",
        displayName: "位置",
        items: VALUE_SHADE_REGION_ITEMS,
        value: VALUE_SHADE_REGION_ITEMS[0],
    });

    shadeMatchLine = this.parts.shadeMatchLine;
    shadeColor = this.parts.shadeColor;
    shadeTransparency = this.parts.shadeTransparency;
    labelShow = this.parts.labelShow;
    labelHorizontal = this.parts.labelHorizontal;
    labelVertical = this.parts.labelVertical;
    labelStyle = this.parts.labelText;
    labelColor = this.parts.labelColor;
    labelFontSize = this.parts.labelFontSize;

    /** 値の表示単位はデータ ラベルと同じ。小数点以下の桁数だけ選べる（自動はデータ ラベルと同じ） */
    labelPrecision = new formattingSettings.ItemDropdown({ name: "labelPrecision", displayName: "小数点以下桁数の値", items: PRECISIONS, value: PRECISIONS[0] });

    labelAvoidOverlap = this.parts.labelAvoidOverlap;

    lineGroup = new Group({
        name: "constantLineLine",
        displayName: "線",
        slices: [this.value, this.labelText, this.color, this.transparency, this.lineStyle, this.dashArray, this.scaleWithWidth, this.dashCap, this.width, this.layer],
    });

    shadeGroup = new Group({
        name: "constantLineShade",
        displayName: "網掛け領域",
        topLevelSlice: this.shadeShow,
        slices: [this.shadeRegion, this.shadeMatchLine, this.shadeColor, this.shadeTransparency],
    });

    labelGroup = new Group({
        name: "constantLineLabel",
        displayName: "データ ラベル",
        topLevelSlice: this.labelShow,
        slices: [this.labelHorizontal, this.labelVertical, this.labelStyle, this.labelColor, this.labelFontSize, this.labelPrecision, this.labelAvoidOverlap],
    });

    groups = [this.lineGroup, this.shadeGroup, this.labelGroup];

    /** カスタムの線の項目はカスタムのときだけ。網掛けの色は線の色に合わせないときだけ */
    syncVisibility(): void {
        syncReferenceLineVisibility(this.parts);
        // 線は 1 本だけなので、ラベルどうしがぶつかることは無い
        this.labelAvoidOverlap.visible = false;
    }
}

/** X 軸（横向きでは Y 軸と呼ぶ） */
export class CategoryAxisCardSettings extends CompositeCard {
    name = "categoryAxis";
    displayName = "X 軸";

    private parts = categoryAxisParts({
        font: { controlName: "fontControl", displayName: "フォント", size: 9, family: DEFAULT_FONT_FAMILY },
        titleFont: { controlName: "titleControl", prefix: "title", displayName: "フォント", size: AXIS_TITLE_FONT_SIZE, family: AXIS_TITLE_FONT_FAMILY },
        titleStyleItems: TITLE_STYLE_ITEMS,
        scrollStartItems: AXIS_SCROLL_START_ITEMS,
        // 高さの最大値・カテゴリの最小幅は、標準と同じく数値の欄＋スライダー
        sliders: true,
    });
    show = this.parts.show;
    font = this.parts.font;
    labelColor = this.parts.labelColor;
    maxHeight = this.parts.maxHeight;
    titleShow = this.parts.titleShow;
    titleText = this.parts.titleText;
    titleStyle = this.parts.titleStyle;
    titleFont = this.parts.titleFont;
    titleColor = this.parts.titleColor;
    minCategoryWidth = this.parts.minCategoryWidth;
    scrollStart = this.parts.scrollStart;

    valuesGroup = new Group({
        name: "categoryValues",
        displayName: "値",
        topLevelSlice: this.show,
        slices: this.parts.valueSlices,
    });

    /**
     * 合計・小計の棒の項目名（起点・左端と右端・小計・最後の合計）。既定で太字にして、増減の項目と見分ける。
     * データ ラベルの「合計のラベル」と同じ形（名前は total を前に付ける）
     */
    totalFont = fontControl("total", "フォント", 9, true);

    /** 空 = 値（項目名）のカラーと同じ */
    totalColor = new formattingSettings.ColorPicker({
        name: "totalColor",
        displayName: "カラー",
        value: { value: "" },
    });

    totalsGroup = new Group({
        name: "categoryTotals",
        displayName: "合計のラベル",
        description: "起点・左端と右端・小計・最後の合計の棒の名前",
        slices: [this.totalFont, this.totalColor],
    });

    titleGroup = this.parts.titleGroup;
    layoutGroup = this.parts.layoutGroup;

    groups = [this.valuesGroup, this.totalsGroup, this.titleGroup, this.layoutGroup];
}

/** Y 軸（横向きでは X 軸と呼ぶ）（保存先の名前：start・end・invertRange・roundRange・switchPosition など） */
export class ValueAxisCardSettings extends CompositeCard {
    name = "valueAxis";
    displayName = "Y 軸";

    private parts = valueAxisParts({
        font: { controlName: "fontControl", displayName: "フォント", size: 9, family: DEFAULT_FONT_FAMILY },
        titleFont: { controlName: "titleControl", prefix: "title", displayName: "フォント", size: AXIS_TITLE_FONT_SIZE, family: AXIS_TITLE_FONT_FAMILY },
        titleStyleItems: TITLE_STYLE_ITEMS,
        unitTypeItems: UNIT_TYPES,
        unitNotationItems: UNIT_NOTATIONS,
        precisionItems: PRECISIONS,
        startDescription: "入れると軸を切る・0 から描くより優先する。0 より大きい値なら、合計の棒に切った印を付ける",
        tickCountDescription: "空なら自動。数を入れると、その本数以内で切りのいい目盛りにする（数字が重なるなら本数を減らす）",
        unitTypeDescription: "自動は増減の大きさで決める。データ ラベルは既定でこの単位",
    });

    start = this.parts.start;
    end = this.parts.end;

    startAtZero = new formattingSettings.ToggleSwitch({
        name: "startAtZero",
        displayName: "0 から描く",
        description: "オフ（既定）では増減が見やすいように軸を切り、合計の棒に切った印を付ける",
        value: false,
    });

    breakStyle = new formattingSettings.ItemDropdown({
        name: "breakStyle",
        displayName: "切った印",
        description: "軸を切ったとき、合計の棒の根元に入れる切れ目",
        items: BREAK_STYLE_ITEMS,
        value: BREAK_STYLE_ITEMS[0],
    });

    invertRange = this.parts.invertRange;
    roundRange = this.parts.roundRange;
    tickCount = this.parts.tickCount;
    show = this.parts.show;
    font = this.parts.font;
    labelColor = this.parts.labelColor;
    unitType = this.parts.unitType;
    unitNotation = this.parts.unitNotation;
    precision = this.parts.precision;
    switchPosition = this.parts.switchPosition;
    titleShow = this.parts.titleShow;
    titleText = this.parts.titleText;
    titleStyle = this.parts.titleStyle;
    titleFont = this.parts.titleFont;
    titleColor = this.parts.titleColor;
    unitShow = this.parts.unitShow;
    unitText = this.parts.unitText;

    unitPosition = new formattingSettings.ItemDropdown({
        name: "unitPosition",
        displayName: "位置",
        items: UNIT_POSITIONS,
        value: UNIT_POSITIONS[0],
    });

    unitStyle = new formattingSettings.ItemDropdown({
        name: "unitStyle",
        displayName: "スタイル",
        items: UNIT_STYLES,
        value: UNIT_STYLES[0],
    });

    /** オフにすると、単位の語（百万など）を付けず、単位の追加文字だけを出す（軸のタイトルの単位も同じ） */
    unitIncludeDisplayUnit = new formattingSettings.ToggleSwitch({
        name: "unitIncludeDisplayUnit",
        displayName: "表示単位（百万など）を付ける",
        value: true,
    });

    /** 空 = 値（目盛り）の文字サイズ。保存していなければ、目盛りの大きさを書式ペインにも出す（applyData） */
    unitFontSize = new AutoNumUpDown({
        name: "unitFontSize",
        displayName: "単位文字サイズ",
        value: undefined,
    });

    /** 空 = 値（目盛り）のカラー。保存していなければ、目盛りの色を書式ペインにも出す（applyData） */
    unitColor = new formattingSettings.ColorPicker({
        name: "unitColor",
        displayName: "単位カラー",
        value: { value: "" },
    });

    rangeGroup = new Group({
        name: "valueRange",
        displayName: "範囲",
        slices: [this.start, this.end, this.startAtZero, this.breakStyle, this.invertRange, this.roundRange, this.tickCount],
    });

    valuesGroup = new Group({
        name: "valueValues",
        displayName: "値",
        topLevelSlice: this.show,
        slices: [this.font, this.labelColor, this.unitType, this.unitNotation, this.precision, this.switchPosition],
    });

    titleGroup = this.parts.titleGroup;

    unitGroup = new Group({
        name: "valueUnit",
        displayName: "単位ラベル",
        topLevelSlice: this.unitShow,
        slices: [this.unitText, this.unitPosition, this.unitStyle, this.unitIncludeDisplayUnit, this.unitFontSize, this.unitColor],
    });

    groups = [this.rangeGroup, this.valuesGroup, this.titleGroup, this.unitGroup];
}

/**
 * グリッド線（値の軸の目盛りの線）。既定は標準と同じ点線。
 * 表示のトグルは「横」のまとまりの見出しに置く（保存先は horizontalShow のまま。1.15 まではカードの見出し）
 */
export class GridlinesCardSettings extends CompositeCard {
    name = "gridlines";
    displayName = "グリッド線";

    private horizontal = gridlineParts({
        prefix: "horizontal",
        showDisplayName: "横",
        show: true,
        styleItems: LINE_STYLE_ITEMS,
        style: itemOf(LINE_STYLE_ITEMS, LINE_STYLES.dotted),
    });

    horizontalShow = this.horizontal.show;
    horizontalColor = this.horizontal.color;
    horizontalTransparency = this.horizontal.transparency;
    horizontalStyle = this.horizontal.style;
    horizontalScaleWithWidth = this.horizontal.scaleWithWidth;
    horizontalDashArray = this.horizontal.dashArray;
    horizontalDashCap = this.horizontal.dashCap;
    horizontalWidth = this.horizontal.width;

    horizontalGroup = new Group({
        name: "horizontalGridlines",
        displayName: "横",
        topLevelSlice: this.horizontalShow,
        slices: this.horizontal.slices,
    });

    groups = [this.horizontalGroup];

    /** 線のスタイルがカスタムのときだけ、ダッシュ配列とダッシュ キャップを出す */
    syncCustomDash(): void {
        this.horizontal.sync();
    }
}

/** データから決まる書式の中身。viewModel が組み、update() のたびに流し込む */
export interface DataDrivenFormat {
    events: string[];
    leftEvent: string;
    rightEvent: string;
    hasOrigin: boolean;
    hasTarget: boolean;
    seriesTargets: SeriesColorTarget[];
    horizontal: boolean;
    /** 比べる形で小計を挟んでいる（浮いた小計の棒がある） */
    subtotalDiff: boolean;
    /** 左端の合計の棒（比べる形の左端のイベント、起点）がある。起点の色を使う */
    hasStart: boolean;
}

const PERCENT_RANGE: NumberRange = [0, 100];

/**
 * 数値の欄の下限・上限（オブジェクト名.プロパティ名）。描くときに丸める範囲と同じにする（viewModel.ts・shared の referenceLine.ts と
 * halo.ts）。フォントの部品の文字サイズは FONT_SIZE_RANGE。範囲がずれていないかは test/numberRanges.test.tsx が描画と突き合わせる
 * 素の numeric の欄にも付ける。以前は書式ペインが丸ごと空になった記録があり、2026-09 の Desktop では再現しなかった（format-pane Skill）。
 * 書式ペインが空になったら、まずここを疑う
 */
export const NUMBER_RANGES: Readonly<Record<string, NumberRange>> = {
    "others.count": [1, 1000],
    "columns.subtotalTransparency": PERCENT_RANGE,
    "columns.outerPadding": PERCENT_RANGE,
    "columns.categorySpacing": [0, 50],
    "columns.cornerRadius": [0, 30],
    "connectors.width": [0.5, 10],
    "dataLabels.backgroundTransparency": PERCENT_RANGE,
    "dataLabels.haloWidth": [0, 10],
    "target.width": [0.5, 10],
    "constantLine.transparency": PERCENT_RANGE,
    "constantLine.width": [1, 10],
    "constantLine.shadeTransparency": PERCENT_RANGE,
    "constantLine.labelFontSize": FONT_SIZE_RANGE,
    "valueAxis.unitFontSize": FONT_SIZE_RANGE,
    "categoryAxis.maxHeight": [5, 100],
    // スライダーは上限が要る。500 px（描くときも 500 で丸める）
    "categoryAxis.minCategoryWidth": [0, 500],
    "gridlines.horizontalTransparency": PERCENT_RANGE,
    "gridlines.horizontalWidth": [0.5, 10],
};

export class VisualFormattingSettingsModel extends Model {
    layout = new LayoutCardSettings();
    others = new OthersCardSettings();
    columns = new ColumnsCardSettings();
    legend = new LegendCardSettings();
    connectors = new ConnectorsCardSettings();
    dataLabels = new DataLabelsCardSettings();
    target = new TargetCardSettings();
    constantLine = new ConstantLineCardSettings();
    categoryAxis = new CategoryAxisCardSettings();
    valueAxis = new ValueAxisCardSettings();
    gridlines = new GridlinesCardSettings();

    /**
     * 並びは標準にそろえる：形の設定 → X 軸 → Y 軸 → 凡例 → グリッド線 → 列 → その他 → つなぎの線 →
     * データ ラベル → 目標の線。定数線は分析ペイン（analyticsPane）
     */
    cards = [
        this.layout,
        this.categoryAxis,
        this.valueAxis,
        this.legend,
        this.gridlines,
        this.columns,
        this.others,
        this.connectors,
        this.dataLabels,
        this.target,
        this.constantLine,
    ];

    constructor() {
        super();
        applyNumberRanges(this.cards, NUMBER_RANGES);
    }

    /**
     * 基本テーマ・カスタムテーマに合わせる。テーマは標準のビジュアルの名前（showAxisTitle・showTitle）で
     * 値を持つので、capabilities にその名前も置いて受け取り、作り手が自作の設定（titleShow）を保存していないときの既定にする。
     * 書式ペインにも同じ値を出す。populate の直後に呼ぶ。凡例の位置は 1.9 までの保存値（topLeft など）を標準の値に読み替える
     */
    applyThemeDefaults(objects: powerbi.DataViewObjects | undefined): void {
        const raw = (card: string, prop: string): unknown => objects?.[card]?.[prop];
        // 項目名の欄の上限は描画と同じ 5〜100、グリッド線の幅は 0.5〜10 に丸める。
        // 自作の「グリッド線」カードは数値の軸の線だけ（horizontal*、横向きでも数値の軸の線）
        const g = this.gridlines;
        applyThemeBasics(objects, {
            categoryTitleShow: this.categoryAxis.titleShow,
            valueTitleShow: this.valueAxis.titleShow,
            legendTitleShow: this.legend.titleShow,
            legendPosition: this.legend.position,
            categoryMaxHeight: this.categoryAxis.maxHeight,
            maxHeightMin: 5,
            gridlines: [{ axis: "valueAxis", prefix: "horizontal", show: g.horizontalShow, color: g.horizontalColor, style: g.horizontalStyle, width: g.horizontalWidth }],
            lineStyleItems: LINE_STYLE_ITEMS,
            gridlineWidthMin: 0.5,
        });
        // 凡例の行数は、位置（読み替えたあと）が上・下のときだけ出す
        this.legend.syncRows();

        // 合計・小計の文字（太字）は標準に項目が無く、テーマの文字の大きさ・フォントが届かない。変えていなければ、ふつうの文字と
        // 同じにする（Fluent 2 は項目名を 10.5 にするので、1.11.1.0 までは合計・小計の名前だけ 9 のまま小さく見えた）
        for (const [card, fonts] of [
            ["categoryAxis", this.categoryAxis],
            ["dataLabels", this.dataLabels],
        ] as const) {
            if (raw(card, "totalFontSize") == null) fonts.totalFont.fontSize.value = fonts.font.fontSize.value;
            if (raw(card, "totalFontFamily") == null) fonts.totalFont.fontFamily.value = fonts.font.fontFamily.value;
        }
    }

    /**
     * データ次第の中身を流し込む。凡例に列が無ければ凡例の色を、目標が無ければ目標のカードを隠す。
     * 横向きでは、項目の軸が縦（Y 軸）、値の軸が横（X 軸）になる（標準の横棒と同じ呼び方）
     */
    applyData(data: DataDrivenFormat): void {
        this.layout.applyEvents(data.events, data.leftEvent, data.rightEvent, data.hasOrigin);
        this.columns.applySeries(data.seriesTargets);
        this.target.visible = data.hasTarget;
        this.gridlines.syncCustomDash();
        this.columns.subtotalStyle.visible = data.subtotalDiff;
        this.columns.subtotalTransparency.visible = data.subtotalDiff && this.columns.subtotalStyle.value?.value === SUBTOTAL_STYLES.fill;
        // 起点の色は左端の合計の棒があるときだけ、その他の色はまとめるときだけ使う。凡例で積むと、その他の棒も凡例の色で塗る
        this.columns.startFill.visible = data.hasStart;
        this.columns.othersFill.visible = (this.others.show.value ?? false) && data.seriesTargets.length === 0;
        // 凡例で積むと増減の棒は凡例の色で塗る。増加・減少の色を使うのは比べる形の小計の棒だけ
        const deltaColors = data.seriesTargets.length === 0 || data.subtotalDiff;
        this.columns.increaseFill.visible = deltaColors;
        this.columns.decreaseFill.visible = deltaColors;
        // 単位ラベルの文字サイズ・カラーは、保存していなければ値（目盛り）と同じものを書式ペインにも出す（保存はしない）
        const axis = this.valueAxis;
        if (axis.unitFontSize.value == null) axis.unitFontSize.value = axis.font.fontSize.value;
        if (!axis.unitColor.value?.value) axis.unitColor.value = { value: axis.labelColor.value?.value ?? "" };
        this.categoryAxis.displayName = data.horizontal ? "Y 軸" : "X 軸";
        this.valueAxis.displayName = data.horizontal ? "X 軸" : "Y 軸";
        this.categoryAxis.maxHeight.displayName = data.horizontal ? "幅の最大値 (%)" : "高さの最大値 (%)";
    }
}
