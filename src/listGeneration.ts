"use strict";

/**
 * 項目の数だけ並ぶ要素（棒・データ ラベル・項目名・接続線・軸を切った印）を、項目を入れ替えたときにも項目の数に比例する時間で描き直す。
 *
 * React は key で並べた子を突き合わせ、新しい子と前より手前に移った子に「置く」印を付けて、1 つずつ DOM に差し込む。
 * 差し込むたびに、後ろに続く「置く」印の子を飛ばして、動かない次の要素を探す（react-dom の getHostSibling）。
 * 続けて置く子が r 個あると r² / 2 回たどるので、ほかのデータから 3 万項目に入れ替えると描き直しに 30 秒を超えた。
 *
 * そこで、続けて置く子が多いときは、並びの入れ物（<g>）の key を変えて入れ物ごと作り直す。新しく作る入れ物の中は、
 * 外れた所で子を順に足してから 1 回だけ差し込むので、項目の数に比例する。続けて置く子が少ない変わり方（数個の増減、
 * 選択・大きさ・書式だけの描き直し）は今までどおり要素を使い回す（フォーカスしている要素も残る）。
 * 作り直すときは、フォーカスしていた子と同じ key の子が新しい並びにあれば、作り直したあとでそこへフォーカスを戻す。
 * マウスが子の上にあれば、子のマウスで出したツールヒントを閉じる（外した子からは mouseleave が来ないため）。
 */

import * as React from "react";

/** 続けて置く子がこれより多ければ、入れ物ごと作り直す。たどる回数は「置く子の数 × これ」までに収まる */
export const PLACEMENT_RUN_LIMIT = 64;

/**
 * 前の key の並び prev から次の並び next に変えるとき、React が「置く」印を付ける子が続く長さの最長。
 * React の突き合わせ（reconcileChildrenArray の placeChild）と同じ決め方：前にあった子は、前の位置が
 * それまでに残した子の前の位置より後ろなら残し、手前なら置き直す。前に無かった子は置く
 */
export function longestPlacementRun(prev: readonly string[], next: readonly string[]): number {
    const oldIndex = new Map<string, number>();
    prev.forEach((key, i) => oldIndex.set(key, i));
    let lastPlaced = 0;
    let run = 0;
    let longest = 0;
    for (const key of next) {
        const old = oldIndex.get(key);
        if (old !== undefined && old >= lastPlaced) {
            lastPlaced = old;
            run = 0;
        } else {
            run++;
            if (run > longest) longest = run;
        }
    }
    return longest;
}

const sameKeys = (a: readonly string[], b: readonly string[]) => a.length === b.length && a.every((key, i) => key === b[i]);

/** 描き終えた並び（DOM にある並び）と、その入れ物の番号 */
export interface CommittedList {
    keys: readonly string[];
    generation: number;
}

/**
 * 次の並び keys を描く入れ物の番号。描き終えた並び committed から、続けて置く子が PLACEMENT_RUN_LIMIT を超える変わり方を
 * するなら 1 増やす（入れ物の key が変わり、入れ物ごと作り直す）。committed が無ければ（初めて描く）0
 */
export function nextGeneration(committed: CommittedList | null, keys: readonly string[]): number {
    if (committed === null) return 0;
    if (sameKeys(committed.keys, keys)) return committed.generation;
    return longestPlacementRun(committed.keys, keys) > PLACEMENT_RUN_LIMIT ? committed.generation + 1 : committed.generation;
}

export interface ListGroupProps {
    /** 子の key を、子と同じ並びで（子 1 つに要素 1 つ。描かない子は入れない） */
    keys: readonly string[];
    className: string;
    id?: string;
    /**
     * 入れ物ごと作り直すとき、マウスが入れ物の中（どれかの子の上）にあったら呼ぶ。外した子からは mouseleave が来ず、
     * 作り直した子にはマウスを動かすまで mouseenter も来ないので、子のマウスで出したツールヒントを閉じる
     */
    onRebuildUnderPointer?: () => void;
    children?: React.ReactNode;
}

/** 作り直す前にフォーカスしていた子の key と、マウスが入れ物の中にあったか */
interface RebuildSnapshot {
    focusedKey: string | null;
    hovered: boolean;
}

/**
 * 項目の数だけ並ぶ子の入れ物（<g>）。続けて置く子が多い入れ替えでは、入れ物ごと作り直す。
 *
 * 比べる元（描き終えた並び）は、描き終えたとき（componentDidMount・componentDidUpdate）に覚え、描画の中では読むだけにする。
 * 描いて捨てた回（React が途中でやめた描画）や StrictMode の二重の描画があっても、DOM にある並びと比べる。
 * フォーカスとマウスは、DOM を変える前（getSnapshotBeforeUpdate）にどの子にあるか・入れ物の中にあるかを見て、作り直したあと
 * （componentDidUpdate）で、フォーカスは同じ key の子に戻し（スクロールの位置は動かさない）、マウスがあれば onRebuildUnderPointer を呼ぶ
 */
export class ListGroup extends React.Component<ListGroupProps> {
    private committed: CommittedList | null = null;
    private readonly element = React.createRef<SVGGElement>();

    /** 描画・DOM を変える前・描き終えたときの 3 回、同じ比べ方で引く（どれも描き終えた並びを読むだけ） */
    private generation(): number {
        return nextGeneration(this.committed, this.props.keys);
    }

    override getSnapshotBeforeUpdate(): RebuildSnapshot | null {
        const box = this.element.current;
        const committed = this.committed;
        if (!box || !committed || this.generation() === committed.generation) return null;
        // 子そのものがフォーカスを受ける（棒・項目名の当たり）。子の中の要素は受けない
        const active = box.ownerDocument.activeElement;
        const focusedKey = active && active.parentNode === box ? (committed.keys[Array.prototype.indexOf.call(box.childNodes, active)] ?? null) : null;
        return { focusedKey, hovered: box.matches(":hover") };
    }

    override componentDidMount(): void {
        this.committed = { keys: this.props.keys, generation: this.generation() };
    }

    override componentDidUpdate(_prev: ListGroupProps, _state: unknown, snapshot: RebuildSnapshot | null): void {
        this.committed = { keys: this.props.keys, generation: this.generation() };
        if (snapshot?.focusedKey != null) this.restoreFocus(snapshot.focusedKey);
        if (snapshot?.hovered) this.props.onRebuildUnderPointer?.();
    }

    /** 作り直した入れ物の、同じ key の子にフォーカスを戻す。無ければ戻さない（前の版も、同じ key の子が無ければ外れた） */
    private restoreFocus(key: string): void {
        const box = this.element.current;
        if (!box) return;
        const doc = box.ownerDocument;
        // 作り直すあいだにほかの所へフォーカスが移ったなら、そちらを残す
        const active = doc.activeElement;
        if (active && active !== doc.body && active.isConnected) return;
        const target = box.childNodes[this.props.keys.indexOf(key)] as (Element & HTMLOrSVGElement) | undefined;
        if (target?.getAttribute("tabindex") != null) target.focus({ preventScroll: true });
    }

    override render(): React.ReactNode {
        const { className, id, children } = this.props;
        // JSX を使わない（.ts のまま）。子は入れ物の 1 つの子として渡す
        return React.createElement("g", { key: `${className}-${this.generation()}`, ref: this.element, className, id }, children);
    }
}
