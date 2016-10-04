/*
    *
    * Wijmo Library 5.20162.211
    * http://wijmo.com/
    *
    * Copyright(c) GrapeCity, Inc.  All rights reserved.
    *
    * Licensed under the Wijmo Commercial License.
    * sales@wijmo.com
    * http://wijmo.com/products/wijmo-5/license/
    *
    */
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
/**
* Defines the @see:FlexGridPdfConverter class used to export the @see:FlexGrid to PDF.
*/
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid_1) {
        var pdf;
        (function (pdf) {
            'use strict';
            /**
            * Specifies how the grid content should be scaled to fit the page.
            */
            (function (ScaleMode) {
                /**
                * Render the grid in actual size, breaking into pages as needed.
                */
                ScaleMode[ScaleMode["ActualSize"] = 0] = "ActualSize";
                /**
                * Scale the grid, so that it fits the page width.
                */
                ScaleMode[ScaleMode["PageWidth"] = 1] = "PageWidth";
                /**
                * Scale the grid, so that it fits on a single page.
                */
                ScaleMode[ScaleMode["SinglePage"] = 2] = "SinglePage";
            })(pdf.ScaleMode || (pdf.ScaleMode = {}));
            var ScaleMode = pdf.ScaleMode;
            /**
            * Specifies whether the whole grid or just a section should be rendered.
            */
            (function (ExportMode) {
                /**
                * Exports all the data from grid.
                */
                ExportMode[ExportMode["All"] = 0] = "All";
                /**
                * Exports the current selection only.
                */
                ExportMode[ExportMode["Selection"] = 1] = "Selection";
            })(pdf.ExportMode || (pdf.ExportMode = {}));
            var ExportMode = pdf.ExportMode;
            /*
            * Merges the content of the source object with the destination object.
            *
            * @param dst The destination object.
            * @param src The source object.
            * @param overwrite Indicates whether the existing properties should be overwritten.
            * @return The modified destination object.
            */
            function _merge(dst, src, overwrite) {
                if (overwrite === void 0) { overwrite = false; }
                if (src && dst) {
                    for (var key in src) {
                        var srcProp = src[key], dstProp = dst[key];
                        if (!wijmo.isObject(srcProp)) {
                            if (dstProp === undefined || (overwrite && srcProp !== undefined)) {
                                dst[key] = srcProp;
                            }
                        }
                        else {
                            if (dstProp === undefined || !wijmo.isObject(dstProp) && overwrite) {
                                dst[key] = dstProp = {};
                            }
                            if (wijmo.isObject(dstProp)) {
                                _merge(dst[key], srcProp, overwrite);
                            }
                        }
                    }
                }
                return dst;
            }
            /**
            * Provides a functionality to export the @see:FlexGrid to PDF.
            */
            var FlexGridPdfConverter = (function () {
                function FlexGridPdfConverter() {
                }
                /**
                * Draws the @see:FlexGrid to an existing @see:PdfDocument at the
                * (0, @wijmo.pdf.PdfDocument.y) coordinates.
                *
                * If width is not specified, then grid will be rendered in actual size,
                * breaking into pages as needed. If height is not specified, then grid will be
                * scaled to fit the width, breaking into pages vertically as needed.
                * If both, width and height are determined, then grid will be scaled to fit
                * the specified rectangle without any page breaks.
                *
                * <pre>
                * var doc = new wijmo.pdf.PdfDocument({
                *    ended: function (sender, args) {
                *       wijmo.pdf.saveBlob(args.blob, 'FlexGrid.pdf');
                *    }
                * });
                *
                * wijmo.grid.pdf.FlexGridPdfConverter.draw(grid, doc, null, null, {
                *    maxPages: 10,
                *    styles: {
                *       cellStyle: {
                *          backgroundColor: '#ffffff',
                *          borderColor: '#c6c6c6'
                *       },
                *       headerCellStyle: {
                *          backgroundColor: '#eaeaea'
                *       }
                *    }
                * });
                * </pre>
                *
                * @param flex The @see:FlexGrid instance to export.
                * @param doc The @see:PdfDocument instance to draw in.
                * @param width The width of the drawing area in points.
                * @param height The height of the drawing area in points.
                * @param settings The draw settings.
                */
                FlexGridPdfConverter.draw = function (flex, doc, width, height, settings) {
                    wijmo.assert(!!flex, 'The flex argument cannot be null.');
                    wijmo.assert(!!doc, 'The doc argument cannot be null.');
                    var options = _merge({}, settings); // clone
                    _merge(options, this.DefaultDrawSettings);
                    if (width == null) {
                        options.scaleMode = ScaleMode.ActualSize;
                    }
                    else {
                        options.scaleMode = height == null ? ScaleMode.PageWidth : ScaleMode.SinglePage;
                    }
                    try {
                        if (options.recalculateStarWidths) {
                            flex.columns._updateStarSizes(wijmo.pdf.ptToPx(doc.width));
                        }
                        this._draw(flex, doc, null, width, height, options);
                    }
                    finally {
                        if (options.recalculateStarWidths) {
                            flex.invalidate(true); // Rollback changes.
                        }
                    }
                };
                /**
                * Draws the @see:FlexGrid to an existing @see:PdfDocument instance at the
                * specified coordinates.
                *
                * If width is not specified, then grid will be rendered in actual size
                * without any page breaks.
                * If height is not specified, then grid will be scaled to fit the width
                * without any page breaks.
                * If both, width and height are determined, then grid will be scaled to fit
                * the specified rectangle without any page breaks.
                *
                * <pre>
                * var doc = new wijmo.pdf.PdfDocument({
                *    ended: function (sender, args) {
                *       wijmo.pdf.saveBlob(args.blob, 'FlexGrid.pdf');
                *    }
                * });
                *
                * wijmo.grid.pdf.FlexGridPdfConverter.drawToPosition(grid, doc, new wijmo.Point(0, 0), null, null, {
                *    maxPages: 10,
                *    styles: {
                *       cellStyle: {
                *          backgroundColor: '#ffffff',
                *          borderColor: '#c6c6c6'
                *       },
                *       headerCellStyle: {
                *          backgroundColor: '#eaeaea'
                *       }
                *    }
                * });
                * </pre>
                *
                * @param flex The @see:FlexGrid instance to export.
                * @param doc The @see:PdfDocument instance to draw in.
                * @param point The position to draw at, in points.
                * @param width The width of the drawing area in points.
                * @param height The height of the drawing area in points.
                * @param settings The draw settings.
                */
                FlexGridPdfConverter.drawToPosition = function (flex, doc, point, width, height, settings) {
                    wijmo.assert(!!flex, 'The flex argument cannot be null.');
                    wijmo.assert(!!doc, 'The doc argument cannot be null.');
                    wijmo.assert(!!point, 'The point argument cannot be null.');
                    var options = _merge({}, settings); // clone
                    _merge(options, this.DefaultDrawSettings);
                    if (width == null) {
                        options.scaleMode = ScaleMode.ActualSize;
                    }
                    else {
                        options.scaleMode = height == null ? ScaleMode.PageWidth : ScaleMode.SinglePage;
                    }
                    try {
                        if (options.recalculateStarWidths) {
                            flex.columns._updateStarSizes(wijmo.pdf.ptToPx(doc.width));
                        }
                        this._draw(flex, doc, point, width, height, options);
                    }
                    finally {
                        if (options.recalculateStarWidths) {
                            flex.invalidate(true); // Rollback changes.
                        }
                    }
                };
                /**
                * Exports the @see:FlexGrid to PDF.
                *
                * <pre>
                * wijmo.grid.pdf.FlexGridPdfConverter.export(grid, 'FlexGrid.pdf', {
                *    scaleMode: wijmo.grid.pdf.ScaleMode.PageWidth,
                *    maxPages: 10,
                *    styles: {
                *       cellStyle: {
                *          backgroundColor: '#ffffff',
                *          borderColor: '#c6c6c6'
                *       },
                *       headerCellStyle: {
                *          backgroundColor: '#eaeaea'
                *       }
                *    },
                *    documentOptions: {
                *       info: {
                *          title: 'Sample'
                *       }
                *    }
                * });
                * </pre>
                *
                * @param flex The @see:FlexGrid instance to export.
                * @param fileName Name of the file to export.
                * @param settings The export settings.
                */
                FlexGridPdfConverter.export = function (flex, fileName, settings) {
                    wijmo.assert(!!flex, 'The flex argument cannot be null.');
                    wijmo.assert(!!fileName, 'The fileName argument cannot be empty.');
                    settings = _merge({}, settings); // clone
                    _merge(settings, this.DefaultExportSettings);
                    var originalEnded = settings.documentOptions.ended;
                    settings.documentOptions.ended = function (sender, args) {
                        wijmo.pdf.saveBlob(args.blob, fileName);
                        if (originalEnded) {
                            originalEnded.apply(doc, [sender, args]);
                        }
                    };
                    var doc = new wijmo.pdf.PdfDocument(settings.documentOptions);
                    try {
                        if (settings.recalculateStarWidths) {
                            flex.columns._updateStarSizes(wijmo.pdf.ptToPx(doc.width));
                        }
                        this._draw(flex, doc, null, null, null, settings);
                        doc.end();
                    }
                    finally {
                        if (settings.recalculateStarWidths) {
                            flex.invalidate(true); // Rollback changes.
                        }
                    }
                };
                FlexGridPdfConverter._draw = function (flex, doc, point, width, height, settings) {
                    var isPositionedMode = point != null, clSize = new wijmo.Size(doc.width, doc.height);
                    if (!point) {
                        point = new wijmo.Point(0, doc.y);
                    }
                    if (wijmo.isArray(settings.embeddedFonts)) {
                        settings.embeddedFonts.forEach(function (font) {
                            doc.registerFont(font);
                        });
                    }
                    // ** initialize
                    var range = RowRange.getSelection(flex, settings.exportMode), grid = new GridRenderer(flex, range, false, this.BorderWidth, null), rect = new wijmo.Rect(point.x || 0, point.y || 0, width || clSize.width, height || clSize.height), scaleFactor = this._getScaleFactor(grid, settings.scaleMode, rect), pages = this._getPages(flex, range, rect, settings, isPositionedMode, scaleFactor);
                    // ** render
                    for (var i = 0; i < pages.length; i++) {
                        if (i > 0) {
                            doc.addPage();
                        }
                        var page = pages[i], x = page.pageCol === 0 ? rect.left : 0, y = page.pageRow === 0 ? rect.top : 0;
                        doc.saveState();
                        doc.paths.rect(0, 0, clSize.width, clSize.height).clip();
                        doc.scale(scaleFactor, scaleFactor, new wijmo.Point(x, y));
                        doc.translate(x, y);
                        var gridPage = new GridRenderer(flex, page.range, settings.repeatMergedValuesAcrossPages, this.BorderWidth, settings.styles);
                        gridPage.render(doc);
                        doc.restoreState();
                        // move document cursor to the grid's left bottom corner.
                        doc.x = x;
                        doc.y = y + grid.renderSize.height * scaleFactor;
                    }
                };
                FlexGridPdfConverter._getScaleFactor = function (grid, scaleMode, rect) {
                    var factor = 1;
                    if (scaleMode === ScaleMode.ActualSize) {
                        return factor;
                    }
                    var size = grid.renderSize;
                    if (scaleMode === ScaleMode.SinglePage) {
                        var f = Math.min(rect.width / size.width, rect.height / size.height);
                        if (f < 1) {
                            factor = f;
                        }
                    }
                    else {
                        var f = rect.width / size.width;
                        if (f < 1) {
                            factor = f;
                        }
                    }
                    return factor;
                };
                FlexGridPdfConverter._getPages = function (flex, ranges, rect, settings, isPositionedMode, scaleFactor) {
                    var _this = this;
                    var rowBreaks = [], colBreaks = [], p2u = wijmo.pdf.pxToPt, showColumnHeader = flex.headersVisibility & grid_1.HeadersVisibility.Column, showRowHeader = flex.headersVisibility & grid_1.HeadersVisibility.Row, colHeaderHeight = showColumnHeader ? p2u(flex.columnHeaders.height) : 0, rowHeaderWidth = showRowHeader ? p2u(flex.rowHeaders.width) : 0, breakRows = settings.scaleMode === ScaleMode.ActualSize || settings.scaleMode === ScaleMode.PageWidth, breakColumns = settings.scaleMode === ScaleMode.ActualSize, zeroColWidth = (rect.width - rect.left) * (1 / scaleFactor), // the width of the leftmost grids
                    zeroRowHeight = (rect.height - rect.top) * (1 / scaleFactor), // the height of the topmost grids
                    rectWidth = rect.width * (1 / scaleFactor), rectHeight = rect.height * (1 / scaleFactor), totalHeight = colHeaderHeight, totalWidth = rowHeaderWidth, 
                    // Normally in ActualSize mode we are inserting page breaks before partially visible columns\ rows to display them completely.
                    // But there is no page breaks in positioned mode, so we need to omit this to fit the maximum amount of content in a drawing area.
                    dontBreakBeforePartiallyVisibleElements = isPositionedMode && (settings.scaleMode == ScaleMode.ActualSize);
                    if (breakRows) {
                        var visibleRowsCnt = 0;
                        ranges.forEach(flex.cells, function (row, rng, rIdx, sIdx) {
                            var renderAreaHeight = rowBreaks.length ? rectHeight : zeroRowHeight;
                            if (PanelSection.isRenderableRow(row)) {
                                var rowHeight = p2u(row.renderHeight);
                                visibleRowsCnt++;
                                totalHeight += rowHeight;
                                if (showColumnHeader || visibleRowsCnt > 1) {
                                    totalHeight -= _this.BorderWidth; // border collapse
                                }
                                if (totalHeight > renderAreaHeight) {
                                    if (colHeaderHeight + rowHeight > renderAreaHeight || dontBreakBeforePartiallyVisibleElements) {
                                        rowBreaks.push(sIdx);
                                        totalHeight = colHeaderHeight;
                                    }
                                    else {
                                        rowBreaks.push(sIdx - 1);
                                        totalHeight = colHeaderHeight + rowHeight;
                                    }
                                    if (showColumnHeader) {
                                        totalHeight -= _this.BorderWidth; // border collapse
                                    }
                                }
                            }
                        });
                    }
                    var len = ranges.length() - 1;
                    if (!rowBreaks.length || (rowBreaks[rowBreaks.length - 1] !== len)) {
                        rowBreaks.push(len);
                    }
                    if (breakColumns) {
                        var visibleColumnsCnt = 0;
                        for (var i = ranges.leftCol; i <= ranges.rightCol; i++) {
                            var col = flex.columns[i];
                            if (col.isVisible) {
                                var colWidth = p2u(col.renderWidth), renderAreaWidth = colBreaks.length ? rectWidth : zeroColWidth;
                                visibleColumnsCnt++;
                                totalWidth += colWidth;
                                if (showRowHeader > 0 || visibleColumnsCnt > 1) {
                                    totalWidth -= this.BorderWidth; // border collapse
                                }
                                if (totalWidth > renderAreaWidth) {
                                    if (rowHeaderWidth + colWidth > renderAreaWidth || dontBreakBeforePartiallyVisibleElements) {
                                        colBreaks.push(i);
                                        totalWidth = rowHeaderWidth;
                                    }
                                    else {
                                        colBreaks.push(i - 1);
                                        totalWidth = rowHeaderWidth + colWidth;
                                    }
                                    if (showRowHeader) {
                                        totalWidth -= this.BorderWidth; // border collapse
                                    }
                                }
                            }
                        }
                    }
                    if (!colBreaks.length || (colBreaks[colBreaks.length - 1] !== ranges.rightCol)) {
                        colBreaks.push(ranges.rightCol);
                    }
                    var pages = [], flag = false, pageCount = 1, maxPages = isPositionedMode && (settings.maxPages > 0) ? 1 : settings.maxPages;
                    for (var i = 0; i < rowBreaks.length && !flag; i++) {
                        for (var j = 0; j < colBreaks.length && !flag; j++, pageCount++) {
                            if (!(flag = pageCount > maxPages)) {
                                var r = i == 0 ? 0 : rowBreaks[i - 1] + 1, c = j == 0 ? ranges.leftCol : colBreaks[j - 1] + 1;
                                pages.push(new PdfPageRowRange(ranges.subrange(r, rowBreaks[i] - r + 1, c, colBreaks[j]), j, i));
                            }
                        }
                    }
                    return pages;
                };
                FlexGridPdfConverter.BorderWidth = 1; // pt, hardcoded because of border collapsing.
                FlexGridPdfConverter.DefaultDrawSettings = {
                    maxPages: Number.MAX_VALUE,
                    exportMode: ExportMode.All,
                    repeatMergedValuesAcrossPages: true,
                    recalculateStarWidths: true,
                    styles: {
                        cellStyle: {
                            font: new wijmo.pdf.PdfFont(),
                            padding: 1.5,
                            verticalAlign: 'middle'
                        },
                        headerCellStyle: {
                            font: { weight: 'bold' } // Don't use PdfFont. It's necessary to specify exclusive properties only, no default values (in order to merge cell styles properly).
                        }
                    }
                };
                FlexGridPdfConverter.DefaultExportSettings = _merge({
                    scaleMode: ScaleMode.PageWidth,
                    documentOptions: {
                        compress: false,
                        pageSettings: {
                            margins: {
                                left: 36,
                                right: 36,
                                top: 18,
                                bottom: 18
                            }
                        }
                    }
                }, FlexGridPdfConverter.DefaultDrawSettings);
                return FlexGridPdfConverter;
            }());
            pdf.FlexGridPdfConverter = FlexGridPdfConverter;
            var PanelSection = (function () {
                function PanelSection(panel, range) {
                    this._panel = panel;
                    this._range = range.clone();
                }
                PanelSection.isRenderableRow = function (row) {
                    return row.isVisible && !(row instanceof grid_1._NewRowTemplate);
                };
                Object.defineProperty(PanelSection.prototype, "visibleRows", {
                    get: function () {
                        var _this = this;
                        if (this._visibleRows == null) {
                            this._visibleRows = 0;
                            this._range.forEach(this._panel, function (row) {
                                if (_this.isRenderableRow(row)) {
                                    _this._visibleRows++;
                                }
                            });
                        }
                        return this._visibleRows;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(PanelSection.prototype, "visibleColumns", {
                    get: function () {
                        if (this._visibleColumns == null) {
                            this._visibleColumns = 0;
                            if (this._range.isValid) {
                                for (var i = this._range.leftCol; i <= this._range.rightCol; i++) {
                                    if (this._panel.columns[i].isVisible) {
                                        this._visibleColumns++;
                                    }
                                }
                            }
                        }
                        return this._visibleColumns;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(PanelSection.prototype, "size", {
                    // pt units
                    get: function () {
                        if (this._size == null) {
                            var sz = this._range.getRenderSize(this._panel);
                            this._size = new wijmo.Size(wijmo.pdf.pxToPt(sz.width), wijmo.pdf.pxToPt(sz.height));
                        }
                        return this._size;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(PanelSection.prototype, "range", {
                    get: function () {
                        return this._range;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(PanelSection.prototype, "panel", {
                    get: function () {
                        return this._panel;
                    },
                    enumerable: true,
                    configurable: true
                });
                PanelSection.prototype.isRenderableRow = function (row) {
                    return PanelSection.isRenderableRow(row);
                };
                return PanelSection;
            }());
            var PanelSectionRenderer = (function (_super) {
                __extends(PanelSectionRenderer, _super);
                function PanelSectionRenderer(flex, panel, range, repeatMergedValues, borderWidth, styles) {
                    _super.call(this, panel, range);
                    this._flex = flex;
                    this._repeatMergedValues = repeatMergedValues;
                    this._borderWidth = borderWidth;
                    this._styles = styles;
                }
                Object.defineProperty(PanelSectionRenderer.prototype, "renderSize", {
                    // pt units
                    get: function () {
                        if (this._renderSize == null) {
                            this._renderSize = this.size.clone();
                            if (this.visibleColumns > 1) {
                                this._renderSize.width -= this._borderWidth * (this.visibleColumns - 1);
                            }
                            if (this.visibleRows > 1) {
                                this._renderSize.height -= this._borderWidth * (this.visibleRows - 1);
                            }
                        }
                        return this._renderSize;
                    },
                    enumerable: true,
                    configurable: true
                });
                PanelSectionRenderer.prototype.getRangeWidth = function (leftCol, rightCol) {
                    var width = 0, visibleColumns = 0, pnl = this.panel;
                    for (var i = leftCol; i <= rightCol; i++) {
                        var col = pnl.columns[i];
                        if (col.isVisible) {
                            visibleColumns++;
                            width += col.renderWidth;
                        }
                    }
                    width = wijmo.pdf.pxToPt(width);
                    if (visibleColumns > 1) {
                        width -= this._borderWidth * (visibleColumns - 1);
                    }
                    return width;
                };
                PanelSectionRenderer.prototype.getRangeHeight = function (topRow, bottomRow) {
                    var height = 0, visibleRows = 0, pnl = this.panel;
                    for (var i = topRow; i <= bottomRow; i++) {
                        var row = pnl.rows[i];
                        if (this.isRenderableRow(row)) {
                            visibleRows++;
                            height += row.renderHeight;
                        }
                    }
                    height = wijmo.pdf.pxToPt(height);
                    if (visibleRows > 1) {
                        height = height - this._borderWidth * (visibleRows - 1);
                    }
                    return height;
                };
                PanelSectionRenderer.prototype.render = function (doc, x, y) {
                    var _this = this;
                    var ranges = this.range, pnl = this.panel, mngr = new GetMergedRangeProxy(this._flex);
                    if (!ranges.isValid) {
                        return;
                    }
                    var pY = {}; // tracks the current Y position for each column
                    for (var c = ranges.leftCol; c <= ranges.rightCol; c++) {
                        pY[c] = y;
                    }
                    ranges.forEach(pnl, function (row, rng, r) {
                        if (!_this.isRenderableRow(row)) {
                            return;
                        }
                        var pX = x;
                        for (var c = ranges.leftCol; c <= ranges.rightCol; c++) {
                            var col = pnl.columns[c], cell = undefined, height = undefined, width = undefined, skipC = undefined;
                            if (!col.isVisible) {
                                continue;
                            }
                            var value = _this._getCellValue(c, r), mergedRng;
                            if ((row.allowMerging || col.allowMerging || row instanceof grid_1.GroupRow || ((typeof (wijmo.grid.multirow) !== "undefined") && (_this._flex instanceof grid_1.multirow.MultiRow))) && (mergedRng = mngr.getMergedRange(pnl, r, c))) {
                                if (mergedRng.topRow !== mergedRng.bottomRow) {
                                    if (mergedRng.topRow === r || r === rng.topRow) {
                                        cell = {
                                            value: _this._repeatMergedValues
                                                ? value
                                                : (mergedRng.topRow === r ? value : ''),
                                            height: height = _this.getRangeHeight(r, Math.min(mergedRng.bottomRow, rng.bottomRow)),
                                            width: width = _this.getRangeWidth(c, c)
                                        };
                                    }
                                    else {
                                        width = _this.getRangeWidth(c, c); // an absorbed cell
                                    }
                                }
                                else {
                                    // c === mrg.leftCol means the very first cell of the range, otherwise it is the remains of the range spreaded between multiple pages
                                    cell = {
                                        value: _this._repeatMergedValues
                                            ? value
                                            : (c === mergedRng.leftCol ? value : ''),
                                        height: height = _this.getRangeHeight(r, r),
                                        width: width = _this.getRangeWidth(Math.max(ranges.leftCol, mergedRng.leftCol), Math.min(ranges.rightCol, mergedRng.rightCol))
                                    };
                                    // skip absorbed cells until the end of the merged range or page (which comes first)
                                    skipC = Math.min(ranges.rightCol, mergedRng.rightCol); // to update loop variable later
                                    for (var t = c + 1; t <= skipC; t++) {
                                        pY[t] += height - _this._borderWidth; // collapse borders
                                    }
                                }
                            }
                            else {
                                cell = {
                                    value: value,
                                    height: height = _this.getRangeHeight(r, r),
                                    width: width = _this.getRangeWidth(c, c)
                                };
                            }
                            if (cell) {
                                _this._renderCell(doc, cell, row, col, r, c, pX, pY[c]);
                            }
                            if (height) {
                                pY[c] += height - _this._borderWidth; // collapse borders
                            }
                            if (width) {
                                pX += width - _this._borderWidth; // collapse borders
                            }
                            if (skipC) {
                                c = skipC;
                            }
                        }
                    });
                };
                PanelSectionRenderer.prototype._getCellValue = function (col, row) {
                    var pnl = this.panel, value = pnl.getCellData(row, col, true);
                    if (!value && value !== 0 && pnl.cellType === grid_1.CellType.Cell) {
                        var flexRow = pnl.rows[row];
                        // seems that FlexGrid doesn't provide an API for getting group header text, so build it manually
                        if (flexRow instanceof grid_1.GroupRow && flexRow.dataItem && flexRow.dataItem.groupDescription && (col === pnl.columns.firstVisibleIndex)) {
                            var propName = flexRow.dataItem.groupDescription.propertyName, column = pnl.columns.getColumn(propName);
                            if (column && column.header) {
                                propName = column.header;
                            }
                            value = propName + ': ' + flexRow.dataItem.name + ' (' + flexRow.dataItem.items.length + ' items)';
                        }
                    }
                    return value;
                };
                PanelSectionRenderer.prototype._isGroupRow = function (row) {
                    return row instanceof grid_1.GroupRow && row.hasChildren; // Group row with no children should be treated as a data row (hierarchical grid)
                };
                PanelSectionRenderer.prototype._renderCell = function (doc, cell, row, column, rowIndex, columnIndex, x, y) {
                    // merge cell styles
                    var cellStyle = _merge({}, this._styles.cellStyle), panel = this.panel;
                    switch (panel.cellType) {
                        case grid_1.CellType.Cell:
                            if (this._isGroupRow(row)) {
                                _merge(cellStyle, this._styles.groupCellStyle, true);
                            }
                            else {
                                // alternating row?
                                if ((typeof (wijmo.grid.multirow) !== 'undefined') && (row instanceof grid_1.multirow._MultiRow)) {
                                    if (row.dataIndex % 2 != 0) {
                                        _merge(cellStyle, this._styles.altCellStyle, true);
                                    }
                                }
                                else {
                                    if (rowIndex % 2 != 0) {
                                        _merge(cellStyle, this._styles.altCellStyle, true);
                                    }
                                }
                            }
                            break;
                        case grid_1.CellType.ColumnHeader:
                        case grid_1.CellType.RowHeader:
                        case grid_1.CellType.TopLeft:
                            _merge(cellStyle, this._styles.headerCellStyle, true);
                            break;
                    }
                    // convert ICellStyle to CssStyleDeclaration
                    var cssStyle = cellStyle; // sharing the some object!
                    if (cellStyle.font) {
                        var tmp;
                        if (tmp = cellStyle.font.family) {
                            cssStyle.fontFamily = tmp;
                        }
                        if (tmp = cellStyle.font.size) {
                            cssStyle.fontSize = tmp;
                        }
                        if (tmp = cellStyle.font.style) {
                            cssStyle.fontStyle = tmp;
                        }
                        if (tmp = cellStyle.font.weight) {
                            cssStyle.fontWeight = tmp;
                        }
                        cellStyle.font = undefined;
                    }
                    // text horizontal alignment
                    // TODO: Does the column.align and column.dataType properties correspond to each other? Need to check...
                    if (!(row instanceof grid_1.GroupRow && !column.aggregate)) {
                        switch (column.dataType) {
                            case wijmo.DataType.Number:
                                cssStyle.textAlign = 'right';
                                break;
                            case wijmo.DataType.Boolean:
                                cssStyle.textAlign = 'center';
                                break;
                        }
                    }
                    cssStyle.left = x;
                    cssStyle.top = y;
                    cssStyle.width = cell.width;
                    cssStyle.height = cell.height;
                    cssStyle.boxSizing = 'no-box';
                    // required border styles
                    cssStyle.borderWidth = this._borderWidth;
                    cssStyle.borderStyle = 'solid';
                    // add indent
                    var grid = panel.grid;
                    if (panel.cellType === grid_1.CellType.Cell && grid.rows.maxGroupLevel >= 0 && columnIndex === grid.columns.firstVisibleIndex) {
                        var level = (row instanceof grid_1.GroupRow)
                            ? Math.max(row.level, 0) // group row cell
                            : grid.rows.maxGroupLevel + 1; // data cell
                        var basePadding = wijmo.pdf._asPt(cssStyle.paddingLeft || cssStyle.padding), levelPadding = wijmo.pdf.pxToPt(level * grid.treeIndent);
                        cssStyle.paddingLeft = (basePadding + levelPadding);
                    }
                    if (column.dataType === wijmo.DataType.Boolean && panel.cellType === grid_1.CellType.Cell && !this._isGroupRow(row)) {
                        doc._renderBooleanCell(cell.value, cssStyle);
                    }
                    else {
                        doc._renderTextCell(cell.value, cssStyle);
                    }
                };
                return PanelSectionRenderer;
            }(PanelSection));
            var GridRenderer = (function () {
                function GridRenderer(flex, range, repeatMergedValues, borderWidth, styles) {
                    this._flex = flex;
                    this._borderWidth = borderWidth;
                    this._topLeft = new PanelSectionRenderer(flex, flex.topLeftCells, this._showRowHeader && this._showColumnHeader
                        ? new RowRange(flex, [new grid_1.CellRange(0, 0, flex.topLeftCells.rows.length - 1, flex.topLeftCells.columns.length - 1)])
                        : new RowRange(flex, []), repeatMergedValues, borderWidth, styles);
                    this._rowHeader = new PanelSectionRenderer(flex, flex.rowHeaders, this._showRowHeader
                        ? range.clone(0, flex.rowHeaders.columns.length - 1)
                        : new RowRange(flex, []), repeatMergedValues, borderWidth, styles);
                    this._columnHeader = new PanelSectionRenderer(flex, flex.columnHeaders, this._showColumnHeader
                        ? new RowRange(flex, [new grid_1.CellRange(0, range.leftCol, flex.columnHeaders.rows.length - 1, range.rightCol)])
                        : new RowRange(flex, []), repeatMergedValues, borderWidth, styles);
                    this._cells = new PanelSectionRenderer(flex, flex.cells, range, repeatMergedValues, borderWidth, styles);
                }
                GridRenderer.prototype.render = function (doc) {
                    var offsetX = Math.max(0, this._rowHeader.renderSize.width - this._borderWidth), offsetY = Math.max(0, this._columnHeader.renderSize.height - this._borderWidth);
                    this._topLeft.render(doc, 0, 0);
                    this._rowHeader.render(doc, 0, offsetY);
                    this._columnHeader.render(doc, offsetX, 0);
                    this._cells.render(doc, offsetX, offsetY);
                };
                Object.defineProperty(GridRenderer.prototype, "renderSize", {
                    get: function () {
                        var height = this._columnHeader.renderSize.height + this._cells.renderSize.height, width = this._rowHeader.renderSize.width + this._cells.renderSize.width;
                        if (this._columnHeader.visibleRows > 0) {
                            height -= this._borderWidth;
                        }
                        if (this._rowHeader.visibleColumns > 0) {
                            width -= this._borderWidth;
                        }
                        return new wijmo.Size(width, height);
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(GridRenderer.prototype, "_showColumnHeader", {
                    get: function () {
                        return !!(this._flex.headersVisibility & grid_1.HeadersVisibility.Column);
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(GridRenderer.prototype, "_showRowHeader", {
                    get: function () {
                        return !!(this._flex.headersVisibility & grid_1.HeadersVisibility.Row);
                    },
                    enumerable: true,
                    configurable: true
                });
                return GridRenderer;
            }());
            // A caching proxy for the flex.getMergedRange method, caches last vertical range for each column.
            var GetMergedRangeProxy = (function () {
                function GetMergedRangeProxy(flex) {
                    this._columns = {};
                    this._flex = flex;
                }
                GetMergedRangeProxy.prototype.getMergedRange = function (panel, r, c) {
                    var rng = this._columns[c];
                    if (rng && r >= rng.topRow && r <= rng.bottomRow) {
                        return rng;
                    }
                    else {
                        return this._columns[c] = this._flex.getMergedRange(panel, r, c, false);
                    }
                };
                return GetMergedRangeProxy;
            }());
            var RowRange = (function () {
                function RowRange(flex, ranges) {
                    this._flex = flex;
                    this._ranges = ranges || [];
                }
                RowRange.getSelection = function (flex, exportMode) {
                    var ranges = [];
                    if (exportMode === ExportMode.All) {
                        ranges.push(new grid_1.CellRange(0, 0, flex.rows.length - 1, flex.columns.length - 1));
                    }
                    else {
                        var selection = flex.selection;
                        switch (flex.selectionMode) {
                            case grid_1.SelectionMode.None:
                                break;
                            case grid_1.SelectionMode.Cell:
                            case grid_1.SelectionMode.CellRange:
                                ranges.push(selection);
                                break;
                            case grid_1.SelectionMode.Row:
                                ranges.push(new grid_1.CellRange(selection.topRow, 0, selection.topRow, flex.cells.columns.length - 1));
                                break;
                            case grid_1.SelectionMode.RowRange:
                                ranges.push(new grid_1.CellRange(selection.topRow, 0, selection.bottomRow, flex.cells.columns.length - 1));
                                break;
                            case grid_1.SelectionMode.ListBox:
                                var top = -1;
                                for (var r = 0; r < flex.rows.length; r++) {
                                    var row = flex.rows[r];
                                    if (row.isSelected) {
                                        if (top < 0) {
                                            top = r;
                                        }
                                        if (r === flex.rows.length - 1) {
                                            ranges.push(new grid_1.CellRange(top, 0, r, flex.cells.columns.length - 1));
                                        }
                                    }
                                    else {
                                        if (top >= 0) {
                                            ranges.push(new grid_1.CellRange(top, 0, r - 1, flex.cells.columns.length - 1));
                                        }
                                        top = -1;
                                    }
                                }
                                break;
                        }
                    }
                    return new RowRange(flex, ranges);
                };
                RowRange.prototype.length = function () {
                    var res = 0;
                    for (var i = 0; i < this._ranges.length; i++) {
                        var r = this._ranges[i];
                        if (r.isValid) {
                            res += r.bottomRow - r.topRow + 1;
                        }
                    }
                    return res;
                };
                Object.defineProperty(RowRange.prototype, "isValid", {
                    get: function () {
                        return this._ranges.length && this._ranges[0].isValid;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(RowRange.prototype, "leftCol", {
                    get: function () {
                        if (this._ranges.length) {
                            return this._ranges[0].leftCol;
                        }
                        return -1;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(RowRange.prototype, "rightCol", {
                    get: function () {
                        if (this._ranges.length) {
                            return this._ranges[0].rightCol;
                        }
                        return -1;
                    },
                    enumerable: true,
                    configurable: true
                });
                RowRange.prototype.clone = function (leftCol, rightCol) {
                    var ranges = [];
                    for (var i = 0; i < this._ranges.length; i++) {
                        var range = this._ranges[i].clone();
                        if (arguments.length > 0) {
                            range.col = leftCol;
                        }
                        if (arguments.length > 1) {
                            range.col2 = rightCol;
                        }
                        ranges.push(range);
                    }
                    return new RowRange(this._flex, ranges);
                };
                RowRange.prototype.getRenderSize = function (panel) {
                    var res = new wijmo.Size(0, 0);
                    for (var i = 0; i < this._ranges.length; i++) {
                        var size = this._ranges[i].getRenderSize(panel);
                        res.width = Math.max(res.width, size.width);
                        res.height += size.height;
                    }
                    return res;
                };
                RowRange.prototype.forEach = function (panel, fn) {
                    var idx = 0;
                    for (var i = 0; i < this._ranges.length; i++) {
                        var range = this._ranges[i];
                        if (range.isValid) {
                            for (var j = range.topRow; j <= range.bottomRow; j++) {
                                fn(panel.rows[j], range, j, idx++);
                            }
                        }
                    }
                };
                RowRange.prototype.subrange = function (from, count, leftCol, rightCol) {
                    var ranges = [];
                    if (from >= 0 && count > 0) {
                        var start = 0, end = 0;
                        for (var i = 0; i < this._ranges.length && count > 0; i++, start = end + 1) {
                            var r = this._ranges[i];
                            end = start + (r.bottomRow - r.topRow);
                            if (from > end) {
                                continue;
                            }
                            var r1 = (from > start) ? r.topRow + (from - start) : r.topRow, r2 = Math.min(r.bottomRow, r1 + count - 1), lCol = arguments.length > 2 ? leftCol : r.leftCol, rCol = arguments.length > 2 ? rightCol : r.rightCol;
                            ranges.push(new grid_1.CellRange(r1, lCol, r2, rCol));
                            count -= r2 - r1 + 1;
                        }
                    }
                    return new RowRange(this._flex, ranges);
                };
                return RowRange;
            }());
            var PdfPageRowRange = (function () {
                function PdfPageRowRange(range, col, row) {
                    this._col = col;
                    this._row = row;
                    this._range = range;
                }
                Object.defineProperty(PdfPageRowRange.prototype, "range", {
                    get: function () {
                        return this._range;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(PdfPageRowRange.prototype, "pageCol", {
                    get: function () {
                        return this._col;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(PdfPageRowRange.prototype, "pageRow", {
                    get: function () {
                        return this._row;
                    },
                    enumerable: true,
                    configurable: true
                });
                return PdfPageRowRange;
            }());
        })(pdf = grid_1.pdf || (grid_1.pdf = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=FlexGridPdfConverter.js.map
