"use strict";

/**
 * ラベルの置き方（docs/waterfall.md の「見せ方」の「ラベルの置き方」）。グラフの種類によらない部品。
 *
 * 1. ぶつかってはいけない物（障害物）を、四角か太さのある線分の一覧に集める
 * 2. ラベルごとに、置いてよい場所の候補を好ましい順に持つ
 * 3. 大事なラベルから順に（呼ぶ側が並べた順に）、候補を上から試し、何ともぶつからない最初の場所に置く。
 *    置いたラベルは、あとのラベルの障害物になる
 * 4. どこにも置けなければ出さない
 *
 * 自分の棒の中に置く・自分の線に沿って置く候補のために、障害物とラベルに持ち主（owner）を付けられる。
 * 同じ持ち主の障害物とはぶつからないものとみなす。棒の中に置く候補は within（この四角の中に収まること）で表す。
 * 障害物が多くても遅くならないよう、升目（CELL）ごとに障害物を分けて持ち、候補の四角が掛かる升目だけ調べる。
 */

export interface Box {
    x: number;
    y: number;
    width: number;
    height: number;
}

/** 障害物。四角か、太さのある線分（つなぎの線・目標の線など。斜めでもよい） */
export type Obstacle =
    | { kind: "box"; box: Box; owner?: string }
    | { kind: "segment"; x1: number; y1: number; x2: number; y2: number; width: number; owner?: string };

export interface Candidate<T> {
    /** 置いたときにラベルが占める四角（余白込み） */
    box: Box;
    /** この四角の中に収まらなければ置けない（棒の中の候補）。無ければ問わない */
    within?: Box;
    item: T;
}

export interface LabelRequest<T> {
    /** 同じ持ち主の障害物とはぶつからないものとみなす（自分の棒・自分の線） */
    owner?: string;
    /** 好ましい順の候補 */
    candidates: Candidate<T>[];
}

export interface PlaceOptions {
    /** ラベルが出てよい範囲。はみ出す候補は置けない。無ければ問わない */
    bounds?: Box;
}

/** 障害物を分けて持つ升目の大きさ（px） */
const CELL = 64;

/** 四角どうしが重なるか（辺が接するだけなら重ならない） */
export const overlaps = (a: Box, b: Box): boolean =>
    a.x < b.x + b.width && a.x + a.width > b.x && a.y < b.y + b.height && a.y + a.height > b.y;

/** a が b の中に収まるか（0.5px までのはみ出しは許す） */
export const contains = (outer: Box, inner: Box): boolean =>
    inner.x >= outer.x - 0.5 &&
    inner.y >= outer.y - 0.5 &&
    inner.x + inner.width <= outer.x + outer.width + 0.5 &&
    inner.y + inner.height <= outer.y + outer.height + 0.5;

/** 線分の外接の四角（太さの半分を足す） */
function extentOf(o: Obstacle): Box {
    if (o.kind === "box") return o.box;
    const half = o.width / 2;
    const x = Math.min(o.x1, o.x2) - half;
    const y = Math.min(o.y1, o.y2) - half;
    return { x, y, width: Math.abs(o.x2 - o.x1) + o.width, height: Math.abs(o.y2 - o.y1) + o.width };
}

/**
 * 太さ w の線分が四角に掛かるか。四角を w/2 だけ広げ、線分がその中を通るかを Liang–Barsky で調べる
 * （線分の端の丸みは見ない。縦横の線では外接の四角と同じになる）
 */
function segmentHits(o: Extract<Obstacle, { kind: "segment" }>, box: Box): boolean {
    const half = o.width / 2;
    const xmin = box.x - half;
    const xmax = box.x + box.width + half;
    const ymin = box.y - half;
    const ymax = box.y + box.height + half;
    const dx = o.x2 - o.x1;
    const dy = o.y2 - o.y1;
    let t0 = 0;
    let t1 = 1;
    const clip = (p: number, q: number): boolean => {
        if (p === 0) return q > 0; // 平行：内側（厳密に）なら通る
        const t = q / p;
        if (p < 0) {
            if (t > t1) return false;
            if (t > t0) t0 = t;
        } else {
            if (t < t0) return false;
            if (t < t1) t1 = t;
        }
        return true;
    };
    return (
        clip(-dx, o.x1 - xmin) && clip(dx, xmax - o.x1) && clip(-dy, o.y1 - ymin) && clip(dy, ymax - o.y1) && t0 < t1
    );
}

function hitsObstacle(o: Obstacle, box: Box): boolean {
    return o.kind === "box" ? overlaps(o.box, box) : segmentHits(o, box);
}

/** これより多くの升目に掛かる障害物（帯・凡例の場所など）は、升目に分けずに毎回まとめて調べる */
const WIDE_CELLS = 64;

/** 升目ごとに障害物を分けて持つ */
class Grid {
    private cells = new Map<string, Obstacle[]>();
    private wide: Obstacle[] = [];

    add(o: Obstacle): void {
        const range = this.rangeOf(extentOf(o));
        if ((range.c1 - range.c0 + 1) * (range.r1 - range.r0 + 1) > WIDE_CELLS) {
            this.wide.push(o);
            return;
        }
        for (const key of this.keysOf(extentOf(o))) {
            const list = this.cells.get(key);
            if (list) list.push(o);
            else this.cells.set(key, [o]);
        }
    }

    /** box に掛かる升目の障害物のどれかにぶつかるか（owner が同じものは見ない） */
    hits(box: Box, owner: string | undefined): boolean {
        for (const o of this.wide) {
            if (owner !== undefined && o.owner === owner) continue;
            if (hitsObstacle(o, box)) return true;
        }
        const seen = new Set<Obstacle>();
        for (const key of this.keysOf(box)) {
            for (const o of this.cells.get(key) ?? []) {
                if (seen.has(o)) continue;
                seen.add(o);
                if (owner !== undefined && o.owner === owner) continue;
                if (hitsObstacle(o, box)) return true;
            }
        }
        return false;
    }

    private rangeOf(box: Box): { c0: number; c1: number; r0: number; r1: number } {
        const cell = (v: number) => Math.floor(Math.max(-1e9, Math.min(1e9, v)) / CELL);
        return { c0: cell(box.x), c1: cell(box.x + box.width), r0: cell(box.y), r1: cell(box.y + box.height) };
    }

    private keysOf(box: Box): string[] {
        const { c0, c1, r0, r1 } = this.rangeOf(box);
        const keys: string[] = [];
        for (let c = c0; c <= c1; c++) for (let r = r0; r <= r1; r++) keys.push(`${c},${r}`);
        return keys;
    }
}

/**
 * ラベルを置く。requests は大事な順に並べて渡す。返すのは requests と同じ並びの、置いた候補（置けなければ null）
 */
export function placeLabels<T>(requests: LabelRequest<T>[], obstacles: Obstacle[], options: PlaceOptions = {}): Array<Candidate<T> | null> {
    const grid = new Grid();
    for (const o of obstacles) grid.add(o);
    const { bounds } = options;
    return requests.map((request) => {
        const chosen =
            request.candidates.find(
                (c) =>
                    (!bounds || contains(bounds, c.box)) &&
                    (!c.within || contains(c.within, c.box)) &&
                    !grid.hits(c.box, request.owner)
            ) ?? null;
        // 置いたラベルは、あとのラベルの障害物になる（持ち主は付けない：同じ持ち主のラベルどうしも避ける）
        if (chosen) grid.add({ kind: "box", box: chosen.box });
        return chosen;
    });
}
