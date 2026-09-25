/*
 *  Power BI Visual
 *  Licensed under the MIT License.
 */
"use strict";

import powerbi from "powerbi-visuals-api";
import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

import "./../style/visual.less";
import { App } from "./App";
import { VisualFormattingSettingsModel } from "./settings";
import { LOADING_NOTICE, TRUNCATED_NOTICE, TooltipTarget, transform, ViewModel } from "./viewModel";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ITooltipService = powerbi.extensibility.ITooltipService;
import ISelectionId = powerbi.visuals.ISelectionId;

export class Visual implements IVisual {
    private root: Root;
    private element: HTMLElement;
    private host: IVisualHost;
    private events: IVisualEventService;
    private selectionManager: ISelectionManager;
    private tooltipService: ITooltipService;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    /** このビジュアルで選ばれている棒。選択を変えたら描き直す */
    private selectedIds: ISelectionId[] = [];
    /** 最後の update の内容で描き直す（選択が外から変わったときに使う） */
    private renderLatest: (() => void) | null = null;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.element = options.element;
        this.events = options.host.eventService;
        this.selectionManager = options.host.createSelectionManager();
        this.tooltipService = options.host.tooltipService;
        this.formattingSettingsService = new FormattingSettingsService();
        this.root = createRoot(options.element);

        // ブックマークの適用などで、選択が外から変わったとき
        this.selectionManager.registerOnSelectCallback((ids: ISelectionId[]) => {
            this.selectedIds = ids;
            this.renderLatest?.();
        });
    }

    public update(options: VisualUpdateOptions): void {
        // レンダリングイベントは認定要件。必ず started / finished(failed) を対で呼ぶ。
        this.events.renderingStarted(options);

        try {
            const dataView = options.dataViews?.[0];
            this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
                VisualFormattingSettingsModel,
                dataView
            );
            // 軸のタイトル・凡例のタイトルと位置は、保存が無ければテーマ（基本テーマの Fluent 2 など）の値に従う
            this.formattingSettings.applyThemeDefaults(dataView?.metadata?.objects);
            const viewModel: ViewModel = transform(dataView, this.host, this.formattingSettings);
            // イベントの選択肢・系列の色などはデータ次第なので、populate のあとに流し込む
            this.formattingSettings.applyData(viewModel.format);
            // 行が 30,000 を超えると、Power BI は続きを残して届ける（metadata.segment）。
            // 続きを足し合わせて取りにいき、取り切れない（メモリの上限）ときは知らせたまま描く
            if (viewModel.truncated) {
                if (this.host.fetchMoreData(true)) {
                    viewModel.notice = LOADING_NOTICE;
                } else {
                    this.host.displayWarningIcon("すべての行を読み込めていません", TRUNCATED_NOTICE);
                }
            }
            this.selectedIds = this.selectionManager.getSelectionIds() as ISelectionId[];
            // 操作できない場所（ダッシュボードのタイルなど）では、選択を受け付けない
            const allowInteractions = this.host.hostCapabilities?.allowInteractions !== false;

            const render = () =>
                this.root.render(
                    React.createElement(App, {
                        viewModel,
                        viewport: options.viewport,
                        selectedIds: this.selectedIds,
                        onSelect: (ids, multiSelect) => {
                            if (!allowInteractions) return;
                            this.selectionManager.select(ids, multiSelect).then((selected) => {
                                this.selectedIds = selected as ISelectionId[];
                                render();
                            });
                        },
                        onClearSelection: () => {
                            if (!allowInteractions || this.selectedIds.length === 0) return;
                            this.selectionManager.clear().then(() => {
                                this.selectedIds = [];
                                render();
                            });
                        },
                        onContextMenu: (id, x, y) => {
                            if (!allowInteractions) return;
                            this.selectionManager.showContextMenu(id ?? ({} as ISelectionId), { x, y });
                        },
                        onTooltipShow: (bar, x, y) => this.showTooltip(bar, x, y, false),
                        onTooltipMove: (bar, x, y) => this.showTooltip(bar, x, y, true),
                        onTooltipHide: () => this.tooltipService.hide({ isTouchEvent: false, immediately: false }),
                    })
                );
            this.renderLatest = render;
            render();

            this.events.renderingFinished(options);
        } catch (error) {
            console.error("update failed", error);
            this.events.renderingFailed(options, String(error));
        }
    }

    /** ツールヒント。中身は viewModel が組んだもの。棒は selectionId を渡す（ドリルスルーが対象の行を知るため）。目標の線は名前と値だけ */
    private showTooltip(target: TooltipTarget, clientX: number, clientY: number, move: boolean): void {
        if (!this.tooltipService.enabled()) return;
        const rect = this.element.getBoundingClientRect();
        const options = {
            coordinates: [clientX - rect.left - this.element.clientLeft, clientY - rect.top - this.element.clientTop],
            isTouchEvent: false,
            dataItems: target.tooltip,
            identities: target.selectionIds,
        };
        if (move) this.tooltipService.move(options);
        else this.tooltipService.show(options);
    }

    /** 書式設定ペインを開くたび / 値変更のたびに呼ばれる */
    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    public destroy(): void {
        this.root?.unmount();
    }
}
