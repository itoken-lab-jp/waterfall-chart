"use strict";

/**
 * 表示単位（万・億など）・文字の幅・色の計算。ほかのビジュアルと共通の部品は shared/ に置き、ここから読む
 */

export type { UnitDefinition } from "./shared/units";
export { UNIT_DEFINITIONS, resolveAutoUnitKey, toStandardUnitKey, resolveUnit, formatValue } from "./shared/units";
export type { FontSpec } from "./shared/text";
export { measureTextWidth, truncateStartToWidth, truncateToWidth } from "./shared/text";
export { DARK_TEXT, luminanceOf, contrastRatio, contrastingText, readableText } from "./shared/color";

/** Y 軸の上に出す単位の表記。「(百万円)」。語も追加文字も無ければ空 */
export function unitBadgeOf(unitWord: string, unitText: string): string {
    const composed = `${unitWord || ""}${(unitText || "").trim()}`;
    return composed ? `(${composed})` : "";
}
