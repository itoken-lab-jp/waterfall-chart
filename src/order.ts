"use strict";

/**
 * 部分的な並びをまとめて 1 本の並びにする（トポロジカルソート）。
 *
 * イベントは categories の 1 列として、項目（階層）ごとに「列で並べ替え」の順で届く。
 * 最初の項目にすべてのイベントがそろっているとは限らないので、初出順だけで決めると、
 * 最初の項目に無いイベントが後ろに回る（dataview-design Skill の「部分順序をマージして解く」）。
 * 項目ごとの並びを部分列として合わせる。決まらない組と循環したぶんは初出順で置く。落とさない。
 */
export function mergeOrders(sequences: string[][]): string[] {
    const firstSeen = new Map<string, number>();
    const next = new Map<string, Set<string>>();
    const indegree = new Map<string, number>();
    const touch = (name: string) => {
        if (!firstSeen.has(name)) {
            firstSeen.set(name, firstSeen.size);
            next.set(name, new Set());
            indegree.set(name, 0);
        }
    };
    for (const sequence of sequences) {
        sequence.forEach((name, i) => {
            touch(name);
            const before = sequence[i - 1];
            if (i > 0 && before !== name && !next.get(before)!.has(name)) {
                next.get(before)!.add(name);
                indegree.set(name, indegree.get(name)! + 1);
            }
        });
    }

    const remaining = new Set(firstSeen.keys());
    const byFirstSeen = (a: string, b: string) => firstSeen.get(a)! - firstSeen.get(b)!;
    const result: string[] = [];
    while (remaining.size) {
        const ready = [...remaining].filter((name) => indegree.get(name) === 0).sort(byFirstSeen);
        // 循環していれば、残りのうち初出の早いものから置く
        const pick = ready[0] ?? [...remaining].sort(byFirstSeen)[0];
        remaining.delete(pick);
        result.push(pick);
        for (const after of next.get(pick)!) {
            if (remaining.has(after)) indegree.set(after, indegree.get(after)! - 1);
        }
    }
    return result;
}
