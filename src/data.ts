"use strict";

/**
 * DataView を、項目 × 値の枠（イベント）× 系列 の表に読み替える。棒の組み立ては viewModel.ts。
 *
 * DataView の形（capabilities.json）：
 * - categories：階層（複数）・小計・イベント。行はこれらの組で届く（形 1 のイベントの列は行を分ける）
 * - values：系列でまとめる（`values.group.by` は 1 本しか持てないので、系列の色と選択が要る系列をこちらにした）
 *   まとまりの中は 値（1 つ以上）・起点・目標・ツールヒント
 *
 * 起点・目標・ツールヒントのメジャーも、行（項目）×系列ごとに評価されて届く。
 * 系列ごとの値がどれも同じなら系列に左右されないメジャーとみなして 1 つ、違えば足す（棒グラフの折れ線と同じ扱い）。
 */

import powerbi from "powerbi-visuals-api";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";

import { mergeOrders } from "./order";

import DataView = powerbi.DataView;
import DataViewCategoryColumn = powerbi.DataViewCategoryColumn;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;
import DataViewValueColumn = powerbi.DataViewValueColumn;
import DataViewValueColumnGroup = powerbi.DataViewValueColumnGroup;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionId = powerbi.visuals.ISelectionId;
import PrimitiveValue = powerbi.PrimitiveValue;

/** 空白の表記。Power BI の日本語の表記に合わせる */
export const BLANK_TEXT = "(空白)";
/** 上の階層の名前のつなぎ */
export const LEVEL_SEPARATOR = " / ";
/** 行を項目にまとめるときの区切り。表示名に現れない文字 */
const KEY_SEPARATOR = String.fromCharCode(31);

export type Formatter = (value: number | null) => string;

export interface ParsedSeries {
    name: string;
    group: DataViewValueColumnGroup;
    /** 系列を選ぶ ID（凡例のクリック・系列の色の保存先）。系列の列が無ければ null */
    selectionId: ISelectionId | null;
}

/** 値の枠 1 本ぶん。形 1 ではイベント 1 つ、形 2・3 では値そのもの */
export interface ParsedSource {
    name: string;
    format: Formatter;
    /** イベントの列のとき、そのイベントの行を選ぶ ID。メジャーのイベントは空 */
    selectionIds: ISelectionId[];
}

/** 項目。同じ小計の区切り・同じ階層の行をまとめたもの */
export interface ParsedItem {
    key: string;
    levels: string[];
    /** この項目の行（どのイベントの行も） */
    rows: number[];
    /** 値の枠ごとの行 */
    sourceRows: number[][];
    /** [値の枠][系列] の合計。行が 1 つも値を持たなければ null */
    values: Array<Array<number | null>>;
    /** [値の枠][系列] のハイライトの合計。該当しなければ null */
    highlights: Array<Array<number | null>>;
}

/** 小計の区切り。name が null の区切りは小計の棒を立てない（空の行が続いたところ） */
export interface ParsedSection {
    key: string;
    name: string | null;
    items: ParsedItem[];
}

/** 項目ごとに 1 つの値を持つメジャー（起点・目標） */
export interface ItemMeasure {
    name: string;
    format: Formatter;
    /** 項目の並びでの値。行が複数なら最初の行 */
    perItem: Array<number | null>;
    /** 同じくハイライト */
    perItemHighlight: Array<number | null>;
}

export interface Parsed {
    error: string | null;
    sources: ParsedSource[];
    series: ParsedSeries[];
    hasSeries: boolean;
    hasEventColumn: boolean;
    sections: ParsedSection[];
    items: ParsedItem[];
    /** 階層のフィールド名（" / " でつなぐ）。階層が無ければ空 */
    hierarchyName: string;
    /** 階層のレベルの名前（上から） */
    levelNames: string[];
    /** 値のフィールドの名前（2 つ以上なら ", " でつなぐ）。軸のタイトルの自動 */
    valueName: string;
    /** 系列のフィールドの名前。系列が無ければ空 */
    seriesName: string;
    origin: ItemMeasure | null;
    target: ItemMeasure | null;
    /** 「ツールヒント」のフィールド。項目の行から値を出す */
    tooltipItemsOf: (rows: number[]) => Array<{ displayName: string; value: string }>;
    hasHighlights: boolean;
    /** Power BI がまだ行を残している（metadata.segment）。取り切れないと両端が合わない */
    truncated: boolean;
    /** 行の選択 ID（系列を付けるなら series の添字） */
    idsOf: (rows: number[], series?: number) => ISelectionId[];
}

export function formatterOf(column: DataViewMetadataColumn | undefined): Formatter {
    const formatter = valueFormatter.create({
        format: column ? valueFormatter.getFormatStringByColumn(column) : undefined,
    });
    return (value) => (value === null ? BLANK_TEXT : formatter.format(value));
}

/** 数値に直す。空白・数値でないものは null（0 とは区別する） */
export function toNumber(raw: PrimitiveValue | undefined): number | null {
    if (raw === null || raw === undefined || raw === "") return null;
    const n = typeof raw === "number" ? raw : Number(raw);
    return Number.isFinite(n) ? n : null;
}

function categoryText(raw: PrimitiveValue | undefined, formatter: valueFormatter.IValueFormatter): string {
    if (raw === null || raw === undefined || raw === "") return BLANK_TEXT;
    return formatter.format(raw);
}

/** 系列ごとに届いた値を 1 つにする。どれも同じなら 1 つ（系列に左右されないメジャー）、違えば足す */
export function collapseAcrossSeries(values: Array<number | null>): number | null {
    const numbers = values.filter((v): v is number => v !== null);
    if (!numbers.length) return null;
    return numbers.every((v) => v === numbers[0]) ? numbers[0] : numbers.reduce((sum, v) => sum + v, 0);
}

const EMPTY_PARSED = (error: string | null): Parsed => ({
    error,
    sources: [],
    series: [],
    hasSeries: false,
    hasEventColumn: false,
    sections: [],
    items: [],
    hierarchyName: "",
    levelNames: [],
    valueName: "",
    seriesName: "",
    origin: null,
    target: null,
    tooltipItemsOf: () => [],
    hasHighlights: false,
    truncated: false,
    idsOf: () => [],
});

export function parse(dataView: DataView | undefined, host: IVisualHost, landing: string): Parsed {
    const categorical = dataView?.categorical;
    const valueColumns = categorical?.values;
    if (!valueColumns || !valueColumns.length) return EMPTY_PARSED(landing);

    const categories: DataViewCategoryColumn[] = categorical?.categories ?? [];
    const hierarchyColumns = categories.filter((c) => c.source?.roles?.hierarchy);
    const subtotalColumn = categories.find((c) => c.source?.roles?.subtotal);
    const eventColumn = categories.find((c) => c.source?.roles?.event);
    // 項目を選ぶ列（一番下の階層）。行の identity がこの列の値を指す
    const itemColumn = hierarchyColumns[hierarchyColumns.length - 1] ?? subtotalColumn ?? categories[0];

    const groups: DataViewValueColumnGroup[] = valueColumns.grouped?.() ?? [
        { values: [...valueColumns] } as DataViewValueColumnGroup,
    ];
    const seriesColumn = valueColumns.source?.roles?.series ? valueColumns.source : undefined;
    const roleOf = (group: DataViewValueColumnGroup, role: string) => group.values.filter((v) => v.source?.roles?.[role]);
    const valueColumnsOf = groups.map((g) => roleOf(g, "value"));
    const valueCount = valueColumnsOf[0]?.length ?? 0;
    const originColumns = groups.map((g) => roleOf(g, "origin")[0]).filter((c): c is DataViewValueColumn => !!c);
    const targetColumns = groups.map((g) => roleOf(g, "target")[0]).filter((c): c is DataViewValueColumn => !!c);

    if (valueCount === 0) {
        return EMPTY_PARSED(originColumns.length ? "値を入れてください。起点に積む項目の値です。" : landing);
    }
    if (eventColumn && valueCount > 1) {
        return EMPTY_PARSED(
            `「比較」に列を入れたときは、値のメジャーを 1 つにしてください（いまは ${valueCount} つ）。2 つ以上のメジャーで比べるときは、「比較」を外してください。`
        );
    }

    const seriesFormatter = seriesColumn
        ? valueFormatter.create({ format: valueFormatter.getFormatStringByColumn(seriesColumn) })
        : null;
    const series: ParsedSeries[] = groups.map((group) => ({
        name: seriesFormatter ? categoryText(group.name, seriesFormatter) : "",
        group,
        selectionId: seriesColumn ? host.createSelectionIdBuilder().withSeries(valueColumns, group).createSelectionId() : null,
    }));

    const rowCount = categories[0]?.values.length ?? valueColumnsOf[0][0].values.length;
    const levelFormatters = hierarchyColumns.map((c) =>
        valueFormatter.create({ format: valueFormatter.getFormatStringByColumn(c.source) })
    );
    const subtotalFormatter = subtotalColumn
        ? valueFormatter.create({ format: valueFormatter.getFormatStringByColumn(subtotalColumn.source) })
        : null;
    const eventFormatter = eventColumn
        ? valueFormatter.create({ format: valueFormatter.getFormatStringByColumn(eventColumn.source) })
        : null;

    // --- 行を項目と小計の区切りにまとめる ---------------------------------------------
    // 名前のある区切りは、離れていても同じ名前ならまとめる（小計の列の値が同じ行をまとめる）。
    // 空の行は、続いているところごとに区切る（名前のある区切りの前と後ろの空の行を寄せない）
    const sections: ParsedSection[] = [];
    const sectionOfKey = new Map<string, ParsedSection>();
    const itemOfKey = new Map<string, ParsedItem>();
    const rowItem: ParsedItem[] = [];
    const rowEventName: string[] = [];
    let blankRun = 0;
    let previousBlank = false;
    for (let i = 0; i < rowCount; i++) {
        const levels = hierarchyColumns.map((c, depth) => categoryText(c.values[i], levelFormatters[depth]));
        const rawSection = subtotalColumn?.values[i];
        const sectionName =
            subtotalColumn && rawSection !== null && rawSection !== undefined && String(rawSection).trim() !== ""
                ? categoryText(rawSection, subtotalFormatter!)
                : null;
        if (sectionName === null && !previousBlank) blankRun++;
        previousBlank = sectionName === null;
        const sectionKey = sectionName === null ? `b${blankRun}` : `s${sectionName}`;
        let section = sectionOfKey.get(sectionKey);
        if (!section) {
            section = { key: sectionKey, name: sectionName, items: [] };
            sectionOfKey.set(sectionKey, section);
            sections.push(section);
        }
        const key = [sectionKey, ...levels].join(KEY_SEPARATOR);
        let item = itemOfKey.get(key);
        if (!item) {
            item = { key, levels, rows: [], sourceRows: [], values: [], highlights: [] };
            itemOfKey.set(key, item);
            section.items.push(item);
        }
        item.rows.push(i);
        rowItem.push(item);
        if (eventColumn) rowEventName.push(categoryText(eventColumn.values[i], eventFormatter!));
    }
    const items = sections.flatMap((s) => s.items);

    // --- 値の枠（イベント） ------------------------------------------------------------
    const firstValueColumn = valueColumnsOf[0][0];
    let sources: ParsedSource[];
    /** 行 i が値の枠 k に入れる値の列（系列 s ごと）。入れなければ null */
    let columnFor: (row: number, source: number, s: number) => DataViewValueColumn | null;
    if (eventColumn) {
        // 項目ごとの並びを合わせてイベントの並びを決める（最初の項目に無いイベントが後ろに回らないように）
        const sequences = items.map((item) => item.rows.map((row) => rowEventName[row]));
        const names = mergeOrders(sequences);
        const indexOf = new Map(names.map((name, k) => [name, k]));
        const format = formatterOf(firstValueColumn.source);
        const rowsOfEvent = names.map((): number[] => []);
        rowEventName.forEach((name, row) => rowsOfEvent[indexOf.get(name)!].push(row));
        sources = names.map((name, k) => ({
            name,
            format,
            selectionIds: dedupe(
                rowsOfEvent[k].map((row) => host.createSelectionIdBuilder().withCategory(eventColumn, row).createSelectionId())
            ),
        }));
        columnFor = (row, source, s) => (indexOf.get(rowEventName[row]) === source ? valueColumnsOf[s][0] : null);
    } else {
        sources = valueColumnsOf[0].map((column): ParsedSource => ({
            name: column.source.displayName,
            format: formatterOf(column.source),
            selectionIds: [],
        }));
        columnFor = (_row, source, s) => valueColumnsOf[s][source] ?? null;
    }

    let hasHighlights = false;
    for (const item of items) {
        item.sourceRows = sources.map((): number[] => []);
        item.values = sources.map(() => series.map((): number | null => null));
        item.highlights = sources.map(() => series.map((): number | null => null));
        for (const row of item.rows) {
            sources.forEach((_, k) => {
                let used = false;
                series.forEach((_, s) => {
                    const column = columnFor(row, k, s);
                    if (!column) return;
                    used = true;
                    const v = toNumber(column.values[row]);
                    if (v !== null) item.values[k][s] = (item.values[k][s] ?? 0) + v;
                    if (column.highlights) {
                        hasHighlights = true;
                        const h = toNumber(column.highlights[row]);
                        if (h !== null) item.highlights[k][s] = (item.highlights[k][s] ?? 0) + h;
                    }
                });
                if (used) item.sourceRows[k].push(row);
            });
        }
    }

    // --- 起点・目標（項目ごとに 1 つ） ---------------------------------------------------
    const itemMeasure = (columns: DataViewValueColumn[]): ItemMeasure | null => {
        if (!columns.length) return null;
        const at = (row: number, pick: (c: DataViewValueColumn) => PrimitiveValue[] | undefined) =>
            collapseAcrossSeries(columns.map((c) => toNumber(pick(c)?.[row])));
        const firstOf = (item: ParsedItem, pick: (c: DataViewValueColumn) => PrimitiveValue[] | undefined) => {
            for (const row of item.rows) {
                const v = at(row, pick);
                if (v !== null) return v;
            }
            return null;
        };
        return {
            name: columns[0].source.displayName,
            format: formatterOf(columns[0].source),
            perItem: items.map((item) => firstOf(item, (c) => c.values)),
            perItemHighlight: items.map((item) => firstOf(item, (c) => c.highlights)),
        };
    };

    // --- ツールヒントに足すフィールド ----------------------------------------------------
    const tooltipColumns = groups.map((g) => roleOf(g, "tooltips"));
    const tooltipCount = tooltipColumns[0]?.length ?? 0;
    const tooltipFormatters = (tooltipColumns[0] ?? []).map((c) =>
        valueFormatter.create({ format: valueFormatter.getFormatStringByColumn(c.source) })
    );
    const tooltipItemsOf = (rows: number[]) =>
        Array.from({ length: tooltipCount }, (_, t) => {
            const perRow = rows.map((row) => {
                const raws = tooltipColumns.map((cols) => cols[t]?.values[row]).filter((v) => v !== null && v !== undefined && v !== "");
                if (!raws.length) return null;
                if (raws.every((v) => typeof v === "number")) return collapseAcrossSeries(raws as number[]);
                return raws[0];
            });
            const present = perRow.filter((v) => v !== null);
            let value: string;
            if (!present.length) value = BLANK_TEXT;
            else if (present.length > 1 && present.every((v) => typeof v === "number"))
                value = tooltipFormatters[t].format((present as number[]).reduce((sum, v) => sum + v, 0));
            else value = tooltipFormatters[t].format(present[0]);
            return { displayName: tooltipColumns[0][t].source.displayName, value };
        });

    const idsOf = (rows: number[], s?: number): ISelectionId[] => {
        if (!itemColumn) return s !== undefined && series[s]?.selectionId ? [series[s].selectionId!] : [];
        return dedupe(
            rows.map((row) => {
                const builder = host.createSelectionIdBuilder().withCategory(itemColumn, row);
                if (s !== undefined && seriesColumn) builder.withSeries(valueColumns, series[s].group);
                return builder.createSelectionId();
            })
        );
    };

    return {
        error: null,
        sources,
        series,
        hasSeries: !!seriesColumn,
        hasEventColumn: !!eventColumn,
        sections,
        items,
        hierarchyName: hierarchyColumns.map((c) => c.source.displayName).join(LEVEL_SEPARATOR),
        levelNames: hierarchyColumns.map((c) => c.source.displayName),
        valueName: valueColumnsOf[0].map((c) => c.source.displayName).join(", "),
        seriesName: seriesColumn?.displayName ?? "",
        origin: itemMeasure(originColumns),
        target: itemMeasure(targetColumns),
        tooltipItemsOf,
        hasHighlights,
        truncated: !!dataView?.metadata?.segment,
        idsOf,
    };
}

/** 同じ行を指す ID を 1 つにする（イベントの行が分かれていても、項目は 1 回だけ選ぶ） */
function dedupe(ids: ISelectionId[]): ISelectionId[] {
    const seen = new Set<string>();
    return ids.filter((id) => {
        const key = id.getKey?.() ?? String(seen.size);
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });
}
