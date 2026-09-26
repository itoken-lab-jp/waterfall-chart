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
import { AutoNumUpDown, itemOf } from "./shared/formatting";
import { NEGATIVE_STYLE_ITEMS, ZERO_STYLE_ITEMS } from "./shared/numberFormat";

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

export const LINE_STYLES = { solid: "solid", dashed: "dashed", dotted: "dotted" } as const;
export type LineStyle = (typeof LINE_STYLES)[keyof typeof LINE_STYLES];

export const LINE_STYLE_ITEMS: powerbi.IEnumMember[] = [
    { value: LINE_STYLES.solid, displayName: "実線" },
    { value: LINE_STYLES.dashed, displayName: "破線" },
    { value: LINE_STYLES.dotted, displayName: "点線" },
];

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

/** 軸のタイトルの出し方 */
export const TITLE_STYLES = { showTitleOnly: "showTitleOnly", showUnitOnly: "showUnitOnly", showBoth: "showBoth" } as const;

export const TITLE_STYLE_ITEMS: powerbi.IEnumMember[] = [
    { value: TITLE_STYLES.showTitleOnly, displayName: "タイトルのみを表示" },
    { value: TITLE_STYLES.showUnitOnly, displayName: "単位のみを表示" },
    { value: TITLE_STYLES.showBoth, displayName: "両方を表示" },
];

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

export const LABEL_ORIENTATIONS = { horizontal: "horizontal", vertical: "vertical" } as const;

export const LABEL_ORIENTATION_ITEMS: powerbi.IEnumMember[] = [
    { value: LABEL_ORIENTATIONS.horizontal, displayName: "横" },
    { value: LABEL_ORIENTATIONS.vertical, displayName: "縦" },
];

/** データ ラベルの表示単位。auto は Y 軸と同じ */
export const LABEL_UNIT_TYPES: powerbi.IEnumMember[] = [{ value: "auto", displayName: "Y 軸と同じ" }, ...UNIT_TYPES.slice(1)];

/** 既定のフォント。レポートのテーマに合わせるのが基本なので Power BI 標準と同じ並びを初期値にする */
export const DEFAULT_FONT_FAMILY = '"Segoe UI", wf_segoe-ui_normal, helvetica, arial, sans-serif';

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

/** ItemDropdown の値の文字列。保存値が items に無いと value は undefined のままになる */
export function dropdownValue(slice: formattingSettings.ItemDropdown, fallback: string): string {
    const raw = slice.value as powerbi.IEnumMember | string | undefined;
    if (raw === undefined || raw === null) return fallback;
    if (typeof raw === "object") return raw.value === undefined || raw.value === null ? fallback : String(raw.value);
    return String(raw);
}

export { AutoNumUpDown };

/**
 * フォント（種類・サイズ・B/I/U）。子のスライスの name が capabilities のプロパティになる。
 * prefix を付けると titleFontFamily のように前に付く（同じカードに 2 つ置くとき）
 */
function fontControl(prefix: string, displayName: string, size: number, bold = false, family = DEFAULT_FONT_FAMILY): formattingSettings.FontControl {
    const n = (suffix: string) => (prefix ? `${prefix}${suffix}` : suffix[0].toLowerCase() + suffix.slice(1));
    return new formattingSettings.FontControl({
        name: `${prefix || "font"}Control`,
        displayName,
        fontFamily: new formattingSettings.FontPicker({ name: n("FontFamily"), displayName: "フォント", value: family }),
        fontSize: new formattingSettings.NumUpDown({ name: n("FontSize"), displayName: "文字サイズ", value: size }),
        bold: new formattingSettings.ToggleSwitch({ name: n("Bold"), displayName: "太字", value: bold }),
        italic: new formattingSettings.ToggleSwitch({ name: n("Italic"), displayName: "斜体", value: false }),
        underline: new formattingSettings.ToggleSwitch({ name: n("Underline"), displayName: "下線", value: false }),
    });
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

    /** ドリルダウンしたとき、今いる位置（事業A ＞ 製品A1 など）を左上に出す（2026-09-25 ユーザー）。標準に無い項目 */
    drillPathShow = new formattingSettings.ToggleSwitch({
        name: "drillPathShow",
        displayName: "ドリルの位置",
        description: "ドリルダウンや絞り込みで、表示している項目の上の階層が 1 つに決まるとき、その位置（事業A ＞ 製品A1 など）を左上に出す",
        value: true,
    });

    slices = [
        this.orientation,
        this.connectMode,
        this.leftEvent,
        this.rightEvent,
        this.order,
        this.originMode,
        this.totalShow,
        this.totalLabel,
        this.drillPathShow,
    ];

    /**
     * イベントの選択肢を流し込み、形で意味の無い設定を隠す。
     * events が 2 つ未満（比べる形でない）ときは、イベントの設定を隠す。
     * left / right は viewModel が生の objects から読んで解決した名前（保存値が今のデータに無ければ既定の最初と最後）。
     * 形 1 では右端のイベントが最後の合計の代わりになるので、最後の合計の設定を隠す。起点の取り方は起点があるときだけ
     */
    applyEvents(events: string[], left: string, right: string, hasOrigin = false): void {
        const compare = events.length >= 2;
        this.connectMode.visible = compare;
        this.leftEvent.visible = compare;
        this.rightEvent.visible = compare;
        this.totalShow.visible = !compare;
        this.totalLabel.visible = !compare;
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

    /** 素の numeric に options を付けない（format-pane Skill）。範囲は viewModel でクランプする */
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
     * 保存先の名前は subtotalTransparency のまま。素の numeric なので options を付けない。viewModel で 0〜100 にクランプする
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
     * 素の numeric なので options を付けない。viewModel で 0〜30 にクランプする
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

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "凡例",
        value: true,
    });

    topLevelSlice = this.show;

    position = new formattingSettings.ItemDropdown({
        name: "position",
        displayName: "位置",
        items: LEGEND_POSITION_ITEMS,
        value: LEGEND_POSITION_ITEMS[0],
    });

    font = fontControl("", "フォント", LEGEND_FONT_SIZE);

    labelColor = new formattingSettings.ColorPicker({
        name: "labelColor",
        displayName: "カラー",
        value: { value: "#605E5C" },
    });

    titleShow = new formattingSettings.ToggleSwitch({
        name: "titleShow",
        displayName: "タイトル",
        value: false,
    });

    titleText = new formattingSettings.TextInput({
        name: "titleText",
        displayName: "タイトル テキスト",
        value: "",
        placeholder: "自動",
    });

    optionsGroup = new Group({ name: "legendOptions", displayName: "オプション", slices: [this.position] });
    textGroup = new Group({ name: "legendText", displayName: "テキスト", slices: [this.font, this.labelColor] });
    titleGroup = new Group({
        name: "legendTitle",
        displayName: "タイトル",
        topLevelSlice: this.titleShow,
        slices: [this.titleText],
    });

    groups = [this.optionsGroup, this.textGroup, this.titleGroup];
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

    /** 素の numeric なので options を付けない。viewModel で 0.5〜10 にクランプする */
    width = new formattingSettings.NumUpDown({
        name: "width",
        displayName: "幅 (px)",
        value: 1,
    });

    lineStyle = new formattingSettings.ItemDropdown({
        name: "lineStyle",
        displayName: "線のスタイル",
        items: LINE_STYLE_ITEMS,
        value: LINE_STYLE_ITEMS[0],
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

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "データ ラベル",
        value: true,
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

    /** 空 = 自動（棒の外は灰色、棒の中は棒の色に合わせて白か黒） */
    color = new formattingSettings.ColorPicker({
        name: "color",
        displayName: "カラー",
        value: { value: "" },
    });

    unitType = new formattingSettings.ItemDropdown({
        name: "unitType",
        displayName: "表示単位",
        items: LABEL_UNIT_TYPES,
        value: LABEL_UNIT_TYPES[0],
    });

    precision = new formattingSettings.ItemDropdown({
        name: "precision",
        displayName: "小数点以下の桁数",
        items: PRECISIONS,
        value: PRECISIONS[0],
    });

    plusSign = new formattingSettings.ToggleSwitch({
        name: "plusSign",
        displayName: "増加に + を付ける",
        value: true,
    });

    /** マイナスの書き方（-・▲・△・括弧）。既定は - */
    negativeStyle = new formattingSettings.ItemDropdown({
        name: "negativeStyle",
        displayName: "マイナス",
        items: NEGATIVE_STYLE_ITEMS,
        value: NEGATIVE_STYLE_ITEMS[0],
    });

    /** 0（丸めて 0 になる値を含む）の書き方（0・±0・-）。既定は 0 */
    zeroStyle = new formattingSettings.ItemDropdown({
        name: "zeroStyle",
        displayName: "0",
        items: ZERO_STYLE_ITEMS,
        value: ZERO_STYLE_ITEMS[0],
    });

    /** 丸めて 0 になるマイナスに符号を残す（▲0）。切ると 0 の書き方にそろえる */
    negativeZero = new formattingSettings.ToggleSwitch({
        name: "negativeZero",
        displayName: "丸めて 0 のマイナスに符号",
        value: true,
    });

    rateType = new formattingSettings.ItemDropdown({
        name: "rateType",
        displayName: "率",
        description: "増減の棒のラベルとツールヒントに率を足す。伸び率・寄与率は比べる形、構成比は合流の形",
        items: RATE_TYPE_ITEMS,
        value: RATE_TYPE_ITEMS[0],
    });

    backgroundShow = new formattingSettings.ToggleSwitch({
        name: "backgroundShow",
        displayName: "背景",
        value: false,
    });

    /** 既定は標準と同じ黒・透過性 90%（棒の色の上にうっすら暗い四角が乗る）。1.7.1.0 までは白・0% */
    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "カラー",
        value: { value: DEFAULT_LABEL_BACKGROUND },
    });

    backgroundTransparency = new formattingSettings.NumUpDown({
        name: "backgroundTransparency",
        displayName: "透過性 (%)",
        value: DEFAULT_LABEL_BACKGROUND_TRANSPARENCY,
    });

    optionsGroup = new Group({
        name: "labelOptions",
        displayName: "オプション",
        slices: [this.orientation, this.position],
    });

    valuesGroup = new Group({
        name: "labelValues",
        displayName: "値",
        slices: [this.font, this.color, this.unitType, this.precision, this.plusSign, this.negativeStyle, this.zeroStyle, this.negativeZero, this.rateType],
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

    groups = [this.optionsGroup, this.valuesGroup, this.totalsGroup, this.backgroundGroup];
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
        value: LINE_STYLE_ITEMS[1],
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
 * format-pane Skill には「analyticsPane = true で書式モデル全体が出なくなった」記録がある（objectCategory の有無は不明）。
 * Desktop で書式ペイン・分析ペインの両方が出るかを確かめる。軸は広げない（標準と同じ。範囲の外なら描かない）
 */
export class ConstantLineCardSettings extends SimpleCard {
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

    labelText = new formattingSettings.TextInput({
        name: "labelText",
        displayName: "名前",
        value: "",
        placeholder: "定数線",
    });

    color = new formattingSettings.ColorPicker({
        name: "color",
        displayName: "カラー",
        value: { value: "#605E5C" },
    });

    width = new formattingSettings.NumUpDown({
        name: "width",
        displayName: "幅 (px)",
        value: 1,
    });

    lineStyle = new formattingSettings.ItemDropdown({
        name: "lineStyle",
        displayName: "線のスタイル",
        items: LINE_STYLE_ITEMS,
        value: LINE_STYLE_ITEMS[1],
    });

    labelShow = new formattingSettings.ToggleSwitch({
        name: "labelShow",
        displayName: "名前と値",
        value: true,
    });

    slices = [this.value, this.labelText, this.color, this.width, this.lineStyle, this.labelShow];
}

/** X 軸（横向きでは Y 軸と呼ぶ） */
export class CategoryAxisCardSettings extends CompositeCard {
    name = "categoryAxis";
    displayName = "X 軸";

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "値",
        value: true,
    });

    font = fontControl("", "フォント", 9);

    labelColor = new formattingSettings.ColorPicker({
        name: "labelColor",
        displayName: "カラー",
        value: { value: "#605E5C" },
    });

    /** 項目名に使う大きさの上限（ビュー全体に対する %）。縦向きは高さ、横向きは幅。素の numeric なので描画側でクランプ */
    maxHeight = new formattingSettings.NumUpDown({
        name: "maxHeight",
        displayName: "高さの最大値 (%)",
        value: 25,
    });

    /** カテゴリの軸のタイトルは初期オフ（2026-09-23 決定）。縦横どちらの向きでもこの card に当てる。数値の軸は初期オンのまま */
    titleShow = new formattingSettings.ToggleSwitch({
        name: "titleShow",
        displayName: "タイトル",
        value: false,
    });

    titleText = new formattingSettings.TextInput({
        name: "titleText",
        displayName: "タイトル テキスト",
        value: "",
        placeholder: "自動",
    });

    titleStyle = new formattingSettings.ItemDropdown({
        name: "titleStyle",
        displayName: "スタイル",
        items: TITLE_STYLE_ITEMS,
        value: TITLE_STYLE_ITEMS[0],
    });

    titleFont = fontControl("title", "フォント", AXIS_TITLE_FONT_SIZE, false, AXIS_TITLE_FONT_FAMILY);

    titleColor = new formattingSettings.ColorPicker({
        name: "titleColor",
        displayName: "カラー",
        value: { value: "#252423" },
    });

    /** 帯がこれより狭くなるときは、スクロールする（標準と同じ）。0 でスクロールしない */
    minCategoryWidth = new formattingSettings.NumUpDown({
        name: "minCategoryWidth",
        displayName: "カテゴリの最小幅 (px)",
        value: 20,
    });

    valuesGroup = new Group({
        name: "categoryValues",
        displayName: "値",
        topLevelSlice: this.show,
        slices: [this.font, this.labelColor, this.maxHeight],
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

    titleGroup = new Group({
        name: "categoryTitle",
        displayName: "タイトル",
        topLevelSlice: this.titleShow,
        slices: [this.titleText, this.titleStyle, this.titleFont, this.titleColor],
    });

    layoutGroup = new Group({
        name: "categoryLayout",
        displayName: "レイアウト",
        slices: [this.minCategoryWidth],
    });

    groups = [this.valuesGroup, this.totalsGroup, this.titleGroup, this.layoutGroup];
}

/** Y 軸（横向きでは X 軸と呼ぶ）（保存先の名前：start・end・invertRange・roundRange・switchPosition など） */
export class ValueAxisCardSettings extends CompositeCard {
    name = "valueAxis";
    displayName = "Y 軸";

    start = new formattingSettings.TextInput({
        name: "start",
        displayName: "最小値",
        description: "入れると軸を切る・0 から描くより優先する。0 より大きい値なら、合計の棒に切った印を付ける",
        value: "",
        placeholder: "自動",
    });

    end = new formattingSettings.TextInput({
        name: "end",
        displayName: "最大値",
        value: "",
        placeholder: "自動",
    });

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

    invertRange = new formattingSettings.ToggleSwitch({
        name: "invertRange",
        displayName: "範囲の反転",
        value: false,
    });

    roundRange = new formattingSettings.ToggleSwitch({
        name: "roundRange",
        displayName: "範囲を丸める",
        value: true,
    });

    /** 目盛り（グリッド線）の本数の目安。空なら自動（描く範囲の長さで決める） */
    tickCount = new formattingSettings.TextInput({
        name: "tickCount",
        displayName: "目盛りの本数 (目安)",
        description: "空なら自動。数を入れると、その本数以内で切りのいい目盛りにする（数字が重なるなら間引く）",
        value: "",
        placeholder: "自動",
    });

    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "値",
        value: true,
    });

    font = fontControl("", "フォント", 9);

    labelColor = new formattingSettings.ColorPicker({
        name: "labelColor",
        displayName: "カラー",
        value: { value: "#605E5C" },
    });

    unitType = new formattingSettings.ItemDropdown({
        name: "unitType",
        displayName: "表示単位",
        description: "自動は増減の大きさで決める。データ ラベルは既定でこの単位",
        items: UNIT_TYPES,
        value: UNIT_TYPES[0],
    });

    unitNotation = new formattingSettings.ItemDropdown({
        name: "unitNotation",
        displayName: "単位の表記",
        items: UNIT_NOTATIONS,
        value: UNIT_NOTATIONS[0],
    });

    precision = new formattingSettings.ItemDropdown({
        name: "precision",
        displayName: "小数点以下の桁数",
        items: PRECISIONS,
        value: PRECISIONS[0],
    });

    switchPosition = new formattingSettings.ToggleSwitch({
        name: "switchPosition",
        displayName: "軸の位置を切り替える",
        value: false,
    });

    titleShow = new formattingSettings.ToggleSwitch({
        name: "titleShow",
        displayName: "タイトル",
        value: true,
    });

    titleText = new formattingSettings.TextInput({
        name: "titleText",
        displayName: "タイトル テキスト",
        value: "",
        placeholder: "自動",
    });

    titleStyle = new formattingSettings.ItemDropdown({
        name: "titleStyle",
        displayName: "スタイル",
        items: TITLE_STYLE_ITEMS,
        value: TITLE_STYLE_ITEMS[0],
    });

    titleFont = fontControl("title", "フォント", AXIS_TITLE_FONT_SIZE, false, AXIS_TITLE_FONT_FAMILY);

    titleColor = new formattingSettings.ColorPicker({
        name: "titleColor",
        displayName: "カラー",
        value: { value: "#252423" },
    });

    unitShow = new formattingSettings.ToggleSwitch({
        name: "unitShow",
        displayName: "単位ラベルの表示",
        value: true,
    });

    unitText = new formattingSettings.TextInput({
        name: "unitText",
        displayName: "単位の追加文字",
        value: "",
        placeholder: "例: 円, 人, 件",
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

    titleGroup = new Group({
        name: "valueTitle",
        displayName: "タイトル",
        topLevelSlice: this.titleShow,
        slices: [this.titleText, this.titleStyle, this.titleFont, this.titleColor],
    });

    unitGroup = new Group({
        name: "valueUnit",
        displayName: "単位ラベル",
        topLevelSlice: this.unitShow,
        slices: [this.unitText],
    });

    groups = [this.rangeGroup, this.valuesGroup, this.titleGroup, this.unitGroup];
}

/** グリッド線（値の軸の目盛りの線）。既定は標準と同じ点線 */
export class GridlinesCardSettings extends SimpleCard {
    name = "gridlines";
    displayName = "グリッド線";

    horizontalShow = new formattingSettings.ToggleSwitch({
        name: "horizontalShow",
        displayName: "グリッド線",
        value: true,
    });

    topLevelSlice = this.horizontalShow;

    horizontalColor = new formattingSettings.ColorPicker({
        name: "horizontalColor",
        displayName: "カラー",
        value: { value: "#E1DFDD" },
    });

    horizontalTransparency = new formattingSettings.NumUpDown({
        name: "horizontalTransparency",
        displayName: "透過性 (%)",
        value: 0,
    });

    horizontalStyle = new formattingSettings.ItemDropdown({
        name: "horizontalStyle",
        displayName: "線のスタイル",
        items: LINE_STYLE_ITEMS,
        value: itemOf(LINE_STYLE_ITEMS, LINE_STYLES.dotted),
    });

    horizontalScaleWithWidth = new formattingSettings.ToggleSwitch({
        name: "horizontalScaleWithWidth",
        displayName: "幅で拡大縮小",
        value: false,
    });

    horizontalWidth = new formattingSettings.NumUpDown({
        name: "horizontalWidth",
        displayName: "幅 (px)",
        value: 1,
    });

    slices = [
        this.horizontalColor,
        this.horizontalTransparency,
        this.horizontalStyle,
        this.horizontalScaleWithWidth,
        this.horizontalWidth,
    ];
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
}

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

    /**
     * 基本テーマ・カスタムテーマに合わせる。テーマは標準のビジュアルの名前（showAxisTitle・showTitle）で
     * 値を持つので、capabilities にその名前も置いて受け取り、作り手が自作の設定（titleShow）を保存していないときの既定にする。
     * 書式ペインにも同じ値を出す。populate の直後に呼ぶ。凡例の位置は 1.9 までの保存値（topLeft など）を標準の値に読み替える
     */
    applyThemeDefaults(objects: powerbi.DataViewObjects | undefined): void {
        const raw = (card: string, prop: string): unknown => objects?.[card]?.[prop];
        const inherit = (slice: formattingSettings.ToggleSwitch, card: string, own: string, standard: string) => {
            const themeValue = raw(card, standard);
            if (raw(card, own) == null && typeof themeValue === "boolean") slice.value = themeValue;
        };
        inherit(this.categoryAxis.titleShow, "categoryAxis", "titleShow", "showAxisTitle");
        inherit(this.valueAxis.titleShow, "valueAxis", "titleShow", "showAxisTitle");
        inherit(this.legend.titleShow, "legend", "titleShow", "showTitle");
        const position = standardLegendPosition(raw("legend", "position"));
        if (position) this.legend.position.value = LEGEND_POSITION_ITEMS.find((i) => i.value === position)!;

        // 項目名の欄の上限：標準は categoryAxis の maxMarginFactor で持つ（Fluent 2 は 50）。描画と同じ 5〜100 に丸める
        const marginFactor = raw("categoryAxis", "maxMarginFactor");
        if (raw("categoryAxis", "maxHeight") == null && typeof marginFactor === "number") {
            this.categoryAxis.maxHeight.value = Math.max(5, Math.min(100, marginFactor));
        }

        // グリッド線：標準は数値の軸のカードの中（gridlineShow・gridlineColor・gridlineStyle・gridlineThickness）。
        // 自作の「グリッド線」カードは数値の軸の線だけ（horizontal*、横向きでも数値の軸の線）
        const own = (prop: string) => raw("gridlines", `horizontal${prop}`) != null;
        const show = raw("valueAxis", "gridlineShow");
        if (!own("Show") && typeof show === "boolean") this.gridlines.horizontalShow.value = show;
        const color = (raw("valueAxis", "gridlineColor") as powerbi.Fill | undefined)?.solid?.color;
        if (!own("Color") && typeof color === "string" && color) this.gridlines.horizontalColor.value = { value: color };
        const style = LINE_STYLE_ITEMS.find((i) => i.value === raw("valueAxis", "gridlineStyle"));
        if (!own("Style") && style) this.gridlines.horizontalStyle.value = style;
        const thickness = raw("valueAxis", "gridlineThickness");
        // 描画と同じ 0.5〜10 に丸める（書式ペインの値と描画の幅をずらさない）
        if (!own("Width") && typeof thickness === "number" && thickness > 0) this.gridlines.horizontalWidth.value = Math.max(0.5, Math.min(10, thickness));

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
        this.columns.subtotalStyle.visible = data.subtotalDiff;
        this.columns.subtotalTransparency.visible = data.subtotalDiff && this.columns.subtotalStyle.value?.value === SUBTOTAL_STYLES.fill;
        this.categoryAxis.displayName = data.horizontal ? "Y 軸" : "X 軸";
        this.valueAxis.displayName = data.horizontal ? "X 軸" : "Y 軸";
        this.categoryAxis.maxHeight.displayName = data.horizontal ? "幅の最大値 (%)" : "高さの最大値 (%)";
    }
}
