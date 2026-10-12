/*
 *  Power BI Visual
 *  Licensed under the MIT License.
 */
"use strict";

import powerbi from "powerbi-visuals-api";
import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { displayWarnings } from "./shared/warnings";
import { sentimentTextColors, showThemeToneColors } from "./shared/numberFormat";

import "./../style/visual.less";
import { App } from "./App";
import { hasBrowserMenu } from "./shared/copyImage";
import { VisualFormattingSettingsModel } from "./settings";
import { showTooltipAt } from "./shared/tooltip";
import { SERIES_TRUNCATED_NOTICE, SERIES_TRUNCATED_TITLE, TRUNCATED_NOTICE, TRUNCATED_TITLE, TooltipTarget, transform, ViewModel } from "./viewModel";

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
            // 符号の色は、書式で決めていなければテーマから作った色を書式ペインにも出す（保存はしない）
            showThemeToneColors(this.formattingSettings.dataLabels, dataView?.metadata?.objects?.dataLabels, sentimentTextColors(this.host.colorPalette as powerbi.extensibility.ISandboxExtendedColorPalette));
            // 行も系列も top で受ける（行を window にして続きを読むと、系列が 60 で切られた）。上限に達したら警告を出す
            displayWarnings(this.host, [
                viewModel.truncated && { title: TRUNCATED_TITLE, detail: TRUNCATED_NOTICE },
                viewModel.seriesTruncated && { title: SERIES_TRUNCATED_TITLE, detail: SERIES_TRUNCATED_NOTICE },
            ]);
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
                        interactive: allowInteractions,
                        // 閲覧中（viewMode が View＝0）は、空のときの案内を出さない
                        editing: options.viewMode === undefined || options.viewMode !== 0,
                        browserMenu: hasBrowserMenu(this.host.hostEnv),
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
        showTooltipAt(this.tooltipService, this.element, clientX, clientY, move, target.tooltip, target.selectionIds);
    }

    /** 書式設定ペインを開くたび / 値変更のたびに呼ばれる */
    public getFormattingModel(): powerbi.visuals.FormattingModel {
        this.formattingSettings.constantLine.syncVisibility();
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    public destroy(): void {
        this.root?.unmount();
    }
}
