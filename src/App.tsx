"use strict";

/**
 * ウォーターフォールの描画。座標は layout.ts の layoutOf が出し、ここは SVG に写すだけにする
 * （テストは layoutOf と描いた SVG の両方を見る）。
 *
 * - 2 枚重ねる：下は棒・項目名・グリッド線など（項目が多いとスクロールする）、上は値の軸の目盛り・タイトル・凡例など
 *   （スクロールしない。凡例のほかはマウスを通す）
 * - 選択：棒・凡例の系列をクリック（Ctrl で追加）、背景で解除。右クリックでコンテキスト メニュー
 * - キーボード：Tab で棒と凡例の系列をたどり、Enter・Space で選ぶ、Esc で解除
 * - ハイライト：該当しない棒を薄く描き、該当分を通常の濃さで重ねる
 * - ハイコントラスト：前景色と背景色だけで描き、増減は線の種類で見分ける
 */

import * as React from "react";
import powerbi from "powerbi-visuals-api";

import { BAR_KINDS, Bar, BarKind, Segment, TooltipTarget, ViewModel } from "./viewModel";
import { ChartLayout, Line, Rect, SegmentLayout, TextLayout, barPath, layoutOf, subtotalPaint } from "./layout";

import IViewport = powerbi.IViewport;
import ISelectionId = powerbi.visuals.ISelectionId;

export { layoutOf } from "./layout";

/** 選んでいない棒・ハイライトに該当しない分の濃さ */
export const DIM_OPACITY = 0.35;
/** 透かして描く棒（比べる形の小計）を薄めたときの下限。透かした濃さにさらに DIM_OPACITY を掛けると消えそうになるため */
export const MIN_DIM_OPACITY = 0.2;

/** 濃さ alpha の棒を、選んでいない・ハイライトに該当しないときに薄めた濃さ（alpha = 1 なら DIM_OPACITY） */
export const dimmedOpacity = (alpha: number): number => Math.max(alpha * DIM_OPACITY, Math.min(alpha, MIN_DIM_OPACITY));

export interface AppProps {
    viewModel: ViewModel;
    viewport: IViewport;
    selectedIds: ISelectionId[];
    onSelect: (ids: ISelectionId[], multiSelect: boolean) => void;
    onClearSelection: () => void;
    /** データの無い場所では id が null */
    onContextMenu: (id: ISelectionId | null, x: number, y: number) => void;
    /** 棒か、目標の線（線のツールヒントは名前と値） */
    onTooltipShow?: (target: TooltipTarget, x: number, y: number) => void;
    onTooltipMove?: (target: TooltipTarget, x: number, y: number) => void;
    onTooltipHide?: () => void;
}

/** ハイコントラストで増減を見分ける線の種類 */
const HC_DASH: Record<BarKind, string | undefined> = {
    [BAR_KINDS.total]: undefined,
    [BAR_KINDS.increase]: undefined,
    [BAR_KINDS.decrease]: "4 2",
    [BAR_KINDS.others]: "1 2",
    [BAR_KINDS.subtotal]: "6 3",
};

/** 文字。(x, y) を中心に回し、背景があれば文字の後ろに敷く */
/** データ ラベルの背景の四角の角の丸み（px）。標準の見た目に近い */
export const LABEL_BACKGROUND_RADIUS = 3;

/** 目標の線にマウスを当てる当たりの太さ（px） */
const REFERENCE_HIT_WIDTH = 7;

/** 棒の形。角丸があれば値の向きの端だけを丸めた path、無ければ今までどおり rect */
function barShape(rect: Rect, corner: SegmentLayout["corner"], props: React.SVGAttributes<SVGElement>): React.ReactElement {
    // 角丸が無いときは 1.11 までと同じ属性の並び（class、位置と大きさ、塗り…）で描く
    const { className, ...rest } = props;
    return corner ? <path className={className} d={barPath(rect, corner)} {...rest} /> : <rect className={className} {...rect} {...rest} />;
}

/** hover：マウスを受けて、title をツールヒントに出す（ドリルの位置。ほかの文字はマウスを通す） */
function Text({ layout, className, opacity, title, hover }: { layout: TextLayout; className: string; opacity?: number; title?: string; hover?: boolean }) {
    const { font } = layout;
    const bg = layout.background;
    return (
        <g
            className={className}
            transform={`translate(${layout.x} ${layout.y})${layout.rotate ? ` rotate(${layout.rotate})` : ""}`}
            opacity={opacity}
            pointerEvents={hover ? "visiblePainted" : "none"}
        >
            {title && <title>{title}</title>}
            {bg && (
                <rect x={bg.x} y={bg.y} width={bg.width} height={bg.height} rx={LABEL_BACKGROUND_RADIUS} fill={bg.color} fillOpacity={bg.opacity} />
            )}
            <text
                x={0}
                y={0}
                textAnchor={layout.anchor}
                style={{
                    fontFamily: font.family,
                    fontSize: font.size,
                    fontWeight: font.weight,
                    fontStyle: font.style,
                    textDecoration: font.decoration,
                    fill: font.color,
                }}
            >
                {layout.lines.map((line, i) => (
                    <tspan key={i} x={0} dy={i === 0 ? 0 : layout.lineHeight}>
                        {line}
                    </tspan>
                ))}
            </text>
        </g>
    );
}

const lineProps = (l: Line) => ({ x1: l.x1, y1: l.y1, x2: l.x2, y2: l.y2 });

export const App: React.FC<AppProps> = ({
    viewModel,
    viewport,
    selectedIds,
    onSelect,
    onClearSelection,
    onContextMenu,
    onTooltipShow,
    onTooltipMove,
    onTooltipHide,
}) => {
    // 斜線の模様の id を、同じページのほかのウォーターフォールとぶつけない（SVG の id はページ全体で引かれる）
    const uid = React.useId().replace(/[^A-Za-z0-9_-]/g, "");
    const backgroundHandlers = {
        onClick: () => onClearSelection(),
        onContextMenu: (e: React.MouseEvent) => {
            e.preventDefault();
            onContextMenu(null, e.clientX, e.clientY);
        },
    };

    if (viewModel.isEmpty) {
        return (
            <div className="wf-landing" {...backgroundHandlers}>
                <p>{viewModel.message}</p>
            </div>
        );
    }

    const layout: ChartLayout | null = layoutOf(viewModel, viewport);
    if (!layout) {
        return <div className="wf-container" style={{ width: viewport.width, height: viewport.height }} {...backgroundHandlers} />;
    }

    const { style } = viewModel;
    const hc = style.highContrast;
    const { scroll } = layout;
    const region = scroll.region;

    const isSelected = (ids: ISelectionId[]) => ids.some((id) => selectedIds.some((s) => s.equals(id)));
    const legendIdOf = (segment: Segment) => {
        if (segment.series === null) return null;
        return viewModel.legend.find((item, i) => item.kind === "series" && i === segment.series)?.selectionId ?? null;
    };
    /** 自分で選んだとき（ハイライトが無いとき）、選んでいない区画を薄くする。凡例で系列を選べば、その系列の区画が残る */
    const segmentPicked = (bar: Bar, segment: Segment) => {
        const legendId = legendIdOf(segment);
        return isSelected(bar.selectionIds) || isSelected(segment.selectionIds) || (legendId !== null && isSelected([legendId]));
    };
    const selectionActive = !viewModel.hasHighlights && selectedIds.length > 0;

    const select = (ids: ISelectionId[], multi: boolean) => {
        if (ids.length) onSelect(ids, multi);
        else onClearSelection();
    };
    const keyHandler = (ids: ISelectionId[]) => (e: React.KeyboardEvent) => {
        if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            e.stopPropagation();
            select(ids, e.ctrlKey || e.metaKey);
        } else if (e.key === "Escape") {
            onClearSelection();
        }
    };

    // 中身の大きさ（スクロールの方向は中身の長さ、もう一方は窓と同じ）
    const contentWidth = scroll.axis === "x" ? scroll.content : region.width;
    const contentHeight = scroll.axis === "y" ? scroll.content : region.height;
    const referenceLines = [layout.target, layout.constantLine].filter((r): r is NonNullable<typeof r> => r !== null);

    return (
        <div className="wf-container" style={{ width: viewport.width, height: viewport.height }} {...backgroundHandlers}>
            <div
                className={`wf-scroll${scroll.axis ? ` wf-scroll-${scroll.axis}` : ""}`}
                style={{ left: region.x, top: region.y, width: region.width, height: region.height }}
            >
                <svg className="wf-svg" width={contentWidth} height={contentHeight} role="img" aria-label="ウォーターフォール">
                    <g transform={`translate(${-region.x} ${-region.y})`}>
                        {style.gridlines.show &&
                            layout.gridlines.map((g, i) => (
                                <line
                                    key={`grid-${i}`}
                                    className="wf-gridline"
                                    {...lineProps(g)}
                                    stroke={style.gridlines.color}
                                    strokeWidth={style.gridlines.width}
                                    strokeOpacity={style.gridlines.opacity}
                                    strokeDasharray={style.gridlines.dash ?? undefined}
                                    strokeLinecap={style.gridlines.dash === "1 3" ? "round" : undefined}
                                />
                            ))}

                        {layout.bars.map((b, index) => {
                            const { bar } = b;
                            const ids = bar.selectionIds;
                            // 比べる形の小計は透かして描く。薄めるときも、透かした濃さから薄める（消えないように下限あり）
                            const alpha = bar.kind === BAR_KINDS.subtotal ? style.subtotalOpacity : 1;
                            const baseOpacity = bar.dimmed ? dimmedOpacity(alpha) : alpha;
                            return (
                                <g
                                    key={bar.key}
                                    className={`wf-bar wf-bar-${bar.kind}`}
                                    role="option"
                                    aria-selected={isSelected(ids)}
                                    aria-label={`${bar.fullLabel}: ${bar.labelText}${bar.rateText ? ` (${bar.rateText})` : ""}`}
                                    tabIndex={0}
                                    data-index={index}
                                    onClick={(e) => {
                                        e.stopPropagation();
                                        select(ids, e.ctrlKey || e.metaKey);
                                    }}
                                    onKeyDown={keyHandler(ids)}
                                    onContextMenu={(e) => {
                                        e.preventDefault();
                                        e.stopPropagation();
                                        onContextMenu(ids[0] ?? null, e.clientX, e.clientY);
                                    }}
                                    onMouseEnter={(e) => onTooltipShow?.(bar, e.clientX, e.clientY)}
                                    onMouseMove={(e) => onTooltipMove?.(bar, e.clientX, e.clientY)}
                                    onMouseLeave={() => onTooltipHide?.()}
                                >
                                    {b.segments.map((s, i) => {
                                        const picked = !selectionActive || segmentPicked(bar, s.segment);
                                        const opacity = picked ? baseOpacity : dimmedOpacity(alpha);
                                        // 比べる形の小計は、ふつうの増減と見分けられるように描き分ける（ハイコントラストでは点線の枠）
                                        const paint = !hc && bar.kind === BAR_KINDS.subtotal ? subtotalPaint(style, s.segment.color) : null;
                                        const hatchId = paint?.hatch ? `wf-hatch-${uid}-${index}-${i}` : null;
                                        const fill = hc
                                            ? bar.kind === BAR_KINDS.total
                                                ? hc.foreground
                                                : hc.background
                                            : paint
                                              ? hatchId
                                                  ? `url(#${hatchId})`
                                                  : paint.fill
                                              : s.segment.color;
                                        return (
                                            <React.Fragment key={i}>
                                                {hatchId && (
                                                    <defs>
                                                        <pattern
                                                            id={hatchId}
                                                            width={5}
                                                            height={5}
                                                            patternUnits="userSpaceOnUse"
                                                            patternTransform="rotate(45)"
                                                        >
                                                            <rect width={5} height={5} fill={paint!.fill} />
                                                            <line x1={0} y1={0} x2={0} y2={5} stroke={s.segment.color} strokeWidth={2.4} />
                                                        </pattern>
                                                    </defs>
                                                )}
                                                {barShape(s.rect, s.corner, {
                                                    className: "wf-bar-rect",
                                                    fill,
                                                    fillOpacity: opacity,
                                                    stroke: hc ? hc.foreground : (paint?.stroke ?? undefined),
                                                    // 選んでいないときは枠も塗りと同じだけ薄める
                                                    strokeOpacity: paint?.stroke ? opacity : undefined,
                                                    strokeWidth: hc ? 1.5 : paint?.stroke ? 1.5 : undefined,
                                                    strokeDasharray: hc ? HC_DASH[bar.kind] : (paint?.dash ?? undefined),
                                                })}
                                                {s.highlight &&
                                                    barShape(s.highlight, s.highlightCorner, {
                                                        className: "wf-bar-highlight",
                                                        // 小計のハイライトも、小計と同じ見せ方で描く（枠だけ・斜線・点線の枠）
                                                        fill: hc ? hc.foreground : paint ? fill : s.segment.color,
                                                        fillOpacity: alpha < 1 ? alpha : undefined,
                                                        stroke: paint?.stroke ?? undefined,
                                                        strokeWidth: paint?.stroke ? 1.5 : undefined,
                                                        strokeDasharray: paint?.dash ?? undefined,
                                                        pointerEvents: "none",
                                                    })}
                                            </React.Fragment>
                                        );
                                    })}
                                    <rect
                                        className="wf-focus"
                                        x={b.rect.x - 2}
                                        y={b.rect.y - 2}
                                        width={b.rect.width + 4}
                                        height={b.rect.height + 4}
                                        fill="none"
                                        stroke={hc ? hc.selected : style.valueAxis.font.color}
                                        strokeWidth={2}
                                        pointerEvents="none"
                                    />
                                </g>
                            );
                        })}

                        {/* 接続線は棒の上に描く（標準と同じ）。線は棒の端に乗るので、棒の後ろに描くと半分が隠れて細く見えた（1.11.8.0 まで） */}
                        {style.connectors.show &&
                            layout.connectors.map((c, i) => (
                                <line
                                    key={`connector-${i}`}
                                    className="wf-connector"
                                    {...lineProps(c)}
                                    stroke={style.connectors.color}
                                    strokeWidth={style.connectors.width}
                                    strokeDasharray={style.connectors.dash ?? undefined}
                                    pointerEvents="none"
                                />
                            ))}

                        {layout.breaks.map((br, i) => (
                            // 切った印。背景の色の隙間と、その両側の細い線（目立たせない。濃い縁取りは使わない）
                            <g key={`break-${i}`} className="wf-break" pointerEvents="none">
                                <path d={br.gap} fill={style.background} />
                                {br.edges.map((d, j) => (
                                    <path key={j} d={d} fill="none" stroke={style.gridlines.color} strokeWidth={0.75} />
                                ))}
                            </g>
                        ))}

                        {referenceLines.map((ref, i) => (
                            <line
                                key={`reference-${i}`}
                                className={ref === layout.target ? "wf-target" : "wf-constant-line"}
                                {...lineProps(ref.line)}
                                stroke={ref.color}
                                strokeWidth={ref.width}
                                strokeDasharray={ref.dash ?? undefined}
                                pointerEvents="none"
                            />
                        ))}
                        {/* 目標の線はマウスを当てるとツールヒントで名前と値を出す（凡例を消しても名前が分かるように）。当たりは線より少し太く */}
                        {referenceLines.map(
                            (ref, i) =>
                                ref.tooltip.length > 0 && (
                                    <line
                                        key={`reference-hit-${i}`}
                                        className="wf-reference-hit"
                                        {...lineProps(ref.line)}
                                        stroke="transparent"
                                        strokeWidth={Math.max(REFERENCE_HIT_WIDTH, ref.width + 4)}
                                        pointerEvents="stroke"
                                        onMouseEnter={(e) => onTooltipShow?.({ tooltip: ref.tooltip, selectionIds: [] }, e.clientX, e.clientY)}
                                        onMouseMove={(e) => onTooltipMove?.({ tooltip: ref.tooltip, selectionIds: [] }, e.clientX, e.clientY)}
                                        onMouseLeave={() => onTooltipHide?.()}
                                    />
                                )
                        )}

                        {style.dataLabels.show &&
                            layout.bars.map(
                                (b) =>
                                    b.label && (
                                        <Text
                                            key={`label-${b.bar.key}`}
                                            layout={b.label}
                                            className="wf-label"
                                            opacity={
                                                selectionActive && !b.segments.some((s) => segmentPicked(b.bar, s.segment)) ? DIM_OPACITY : 1
                                            }
                                        />
                                    )
                            )}

                        {layout.bars.map(
                            (b) =>
                                b.category && (
                                    <Text
                                        key={`category-${b.bar.key}`}
                                        layout={b.category}
                                        className={layout.rotatedCategories ? "wf-category wf-category-rotated" : "wf-category"}
                                        title={b.category.title}
                                    />
                                )
                        )}
                    </g>
                </svg>
            </div>

            {/* スクロールしないもの。凡例のほかはマウスを通す */}
            <svg className="wf-overlay" width={viewport.width} height={viewport.height} pointerEvents="none">
                {layout.badge && <Text layout={layout.badge} className="wf-unit" />}
                {layout.drillPath && <Text layout={layout.drillPath} className="wf-drill-path" title={layout.drillPath.title} hover />}
                {layout.ticks.map((t, i) => (
                    <Text key={`tick-${i}`} layout={t} className="wf-tick" />
                ))}
                {layout.valueTitle && <Text layout={layout.valueTitle} className="wf-value-title" />}
                {layout.categoryTitle && <Text layout={layout.categoryTitle} className="wf-category-title" />}
                {referenceLines.map((ref, i) => ref.label && <Text key={`reference-label-${i}`} layout={ref.label} className="wf-reference-label" />)}

                {layout.legend && (
                    <g className="wf-legend" role="listbox" aria-label="凡例">
                        {layout.legend.title && <Text layout={layout.legend.title} className="wf-legend-title" />}
                        {layout.legend.items.map(({ item, index, x, swatchWidth, y, textY, text }) => {
                            const ids = item.selectionId ? [item.selectionId] : [];
                            const selectable = ids.length > 0;
                            const dim = selectionActive && selectable && !isSelected(ids);
                            const font = layout.legend!.font;
                            return (
                                <g
                                    key={`legend-${index}`}
                                    className={selectable ? "wf-legend-item wf-legend-selectable" : "wf-legend-item"}
                                    role="option"
                                    aria-selected={isSelected(ids)}
                                    aria-label={item.name}
                                    tabIndex={selectable ? 0 : undefined}
                                    opacity={dim ? DIM_OPACITY : 1}
                                    pointerEvents={selectable ? "auto" : "none"}
                                    onClick={(e) => {
                                        e.stopPropagation();
                                        select(ids, e.ctrlKey || e.metaKey);
                                    }}
                                    onKeyDown={keyHandler(ids)}
                                >
                                    {item.line ? (
                                        // 目標の線：グラフの線と同じ色・種類の線（色はハイコントラストでは前景色になっている）
                                        <line
                                            className="wf-legend-line"
                                            x1={x}
                                            y1={y + 5}
                                            x2={x + swatchWidth}
                                            y2={y + 5}
                                            stroke={item.color}
                                            strokeWidth={item.line.width}
                                            strokeDasharray={item.line.dash ?? undefined}
                                        />
                                    ) : (
                                        // 棒に角丸があれば、印も値の向きの端（縦向きは上、横向きは右）を丸める
                                        barShape(
                                            { x, y, width: swatchWidth, height: 10 },
                                            style.columns.cornerRadius > 0
                                                ? {
                                                      // 合計は値の端だけ（範囲の反転では逆の端）、浮いた増減は両端（棒と同じ）
                                                      sides:
                                                          style.orientation === "horizontal"
                                                              ? item.kind === BAR_KINDS.total
                                                                  ? [style.valueAxis.invert ? "left" : "right"]
                                                                  : ["left", "right"]
                                                              : item.kind === BAR_KINDS.total
                                                                ? [style.valueAxis.invert ? "bottom" : "top"]
                                                                : ["top", "bottom"],
                                                      radius: Math.min(style.columns.cornerRadius, 10 / 3),
                                                  }
                                                : null,
                                            {
                                                fill: hc ? (item.kind === BAR_KINDS.total ? hc.foreground : hc.background) : item.color,
                                                stroke: hc ? hc.foreground : undefined,
                                                strokeDasharray: hc && item.kind !== "series" && item.kind !== "target" ? HC_DASH[item.kind as BarKind] : undefined,
                                            }
                                        )
                                    )}
                                    <text
                                        x={x + swatchWidth + 4}
                                        y={textY}
                                        style={{
                                            fontFamily: font.family,
                                            fontSize: font.size,
                                            fontWeight: font.weight,
                                            fontStyle: font.style,
                                            textDecoration: font.decoration,
                                            fill: font.color,
                                        }}
                                    >
                                        {text}
                                    </text>
                                </g>
                            );
                        })}
                    </g>
                )}

                {layout.notice && (
                    <text
                        className="wf-notice"
                        x={layout.notice.x}
                        y={layout.notice.y}
                        style={{ fontFamily: style.valueAxis.font.family, fontSize: 11, fill: style.valueAxis.font.color }}
                    >
                        {viewModel.notice}
                    </text>
                )}
            </svg>
        </div>
    );
};
