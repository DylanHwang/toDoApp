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
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var multirow;
        (function (multirow) {
            'use strict';
            /**
             * Extends the @see:Row to provide additional information for multi-row records.
             */
            var _MultiRow = (function (_super) {
                __extends(_MultiRow, _super);
                /**
                 * Initializes a new instance of the @see:Row class.
                 *
                 * @param dataItem The data item this row is bound to.
                 * @param dataIndex The index of the record within the items source.
                 * @param recordIndex The index of this row within the record (data item).
                 */
                function _MultiRow(dataItem, dataIndex, recordIndex) {
                    _super.call(this, dataItem);
                    this._idxData = dataIndex;
                    this._idxRecord = recordIndex;
                }
                Object.defineProperty(_MultiRow.prototype, "recordIndex", {
                    /**
                     * Gets the index of this row within the record (data item) it represents.
                     */
                    get: function () {
                        return this._idxRecord;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(_MultiRow.prototype, "dataIndex", {
                    /**
                     * Gets the index of this row within the data source collection.
                     */
                    get: function () {
                        return this._idxData;
                    },
                    enumerable: true,
                    configurable: true
                });
                return _MultiRow;
            }(grid.Row));
            multirow._MultiRow = _MultiRow;
        })(multirow = grid.multirow || (grid.multirow = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_MultiRow.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var multirow;
        (function (multirow) {
            'use strict';
            /**
             * Extends the @see:Column class with <b>colspan</b> property to
             * describe a cell in a @see:_CellGroup.
             */
            var _Cell = (function (_super) {
                __extends(_Cell, _super);
                /**
                 * Initializes a new instance of the @see:_Cell class.
                 *
                 * @param options JavaScript object containing initialization data for the @see:_Cell.
                 */
                function _Cell(options) {
                    _super.call(this);
                    this._row = this._col = 0;
                    this._rowspan = this._colspan = 1;
                    if (options) {
                        wijmo.copy(this, options);
                    }
                }
                Object.defineProperty(_Cell.prototype, "colspan", {
                    /**
                     * Gets or sets the number of physical columns spanned by the @see:_Cell.
                     */
                    get: function () {
                        return this._colspan;
                    },
                    set: function (value) {
                        this._colspan = wijmo.asInt(value, false, true);
                    },
                    enumerable: true,
                    configurable: true
                });
                return _Cell;
            }(grid.Column));
            multirow._Cell = _Cell;
        })(multirow = grid.multirow || (grid.multirow = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_Cell.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid_1) {
        var multirow;
        (function (multirow) {
            'use strict';
            /**
             * Describes a group of cells that may span multiple rows and columns.
             */
            var _CellGroup = (function (_super) {
                __extends(_CellGroup, _super);
                /**
                 * Initializes a new instance of the @see:_CellGroup class.
                 *
                 * @param grid @see:MultiRow that owns the @see:_CellGroup.
                 * @param options JavaScript object containing initialization data for the new @see:_CellGroup.
                 */
                function _CellGroup(grid, options) {
                    _super.call(this);
                    this._colstart = 0; // index of the column where this group starts
                    // save reference to owner grid
                    this._g = grid;
                    // parse options
                    if (options) {
                        wijmo.copy(this, options);
                    }
                    if (!this._cells) {
                        throw 'Cell group with no cells?';
                    }
                    // count rows/columns
                    var r = 0, c = 0;
                    for (var i = 0; i < this._cells.length; i++) {
                        var cell = this._cells[i];
                        // if the cell doesn't fit in this row, start a new row
                        if (c + cell.colspan > this._colspan) {
                            r++;
                            c = 0;
                        }
                        // store cell position within the group
                        cell._row = r;
                        cell._col = c;
                        // update column and continue
                        c += cell.colspan;
                    }
                    this._rowspan = r + 1;
                    // adjust colspans to fill every row
                    for (var i = 0; i < this._cells.length; i++) {
                        var cell = this._cells[i];
                        if (i == this._cells.length - 1 || this._cells[i + 1]._row > cell._row) {
                            c = cell._col;
                            cell._colspan = this._colspan - c;
                        }
                    }
                }
                // method used in JSON-style initialization
                _CellGroup.prototype._copy = function (key, value) {
                    if (key == 'cells') {
                        this._cells = [];
                        if (wijmo.isArray(value)) {
                            for (var i = 0; i < value.length; i++) {
                                var cell = new multirow._Cell(value[i]);
                                if (!value[i].header && cell.binding) {
                                    value.header = wijmo.toHeaderCase(cell.binding);
                                }
                                this._cells.push(cell);
                                this._colspan = Math.max(this._colspan, cell.colspan);
                            }
                        }
                        return true;
                    }
                    return false;
                };
                Object.defineProperty(_CellGroup.prototype, "cells", {
                    // required for JSON-style initialization
                    get: function () {
                        return this._cells;
                    },
                    enumerable: true,
                    configurable: true
                });
                // calculate merged ranges
                _CellGroup.prototype.closeGroup = function (rowsPerItem) {
                    // adjust rowspan to match longest group in the grid
                    if (rowsPerItem > this._rowspan) {
                        for (var i = 0; i < this._cells.length; i++) {
                            var cell = this._cells[i];
                            if (cell._row == this._rowspan - 1) {
                                cell._rowspan = rowsPerItem - cell._row;
                            }
                        }
                        this._rowspan = rowsPerItem;
                    }
                    // create arrays with binding columns and merge ranges for each cell
                    this._cols = new grid_1.ColumnCollection(this._g, this._g.columns.defaultSize);
                    this._rng = new Array(rowsPerItem * this._colspan);
                    for (var i = 0; i < this._cells.length; i++) {
                        var cell = this._cells[i];
                        for (var r = 0; r < cell._rowspan; r++) {
                            for (var c = 0; c < cell._colspan; c++) {
                                var index = (cell._row + r) * this._colspan + (cell._col) + c;
                                // save binding column for this cell offset
                                // (using 'setAt' to handle list ownership)
                                this._cols.setAt(index, cell);
                                //console.log('binding[' + index + '] = ' + cell.binding);
                                // save merge range for this cell offset
                                var rng = new grid_1.CellRange(0 - r, 0 - c, 0 - r + cell._rowspan - 1, 0 - c + cell._colspan - 1);
                                if (!rng.isSingleCell) {
                                    //console.log('rng[' + index + '] = ' + format('({row},{col})-({row2},{col2})', rng));
                                    this._rng[index] = rng;
                                }
                            }
                        }
                    }
                    // add extra range for collapsed group headers
                    this._rng[-1] = new grid_1.CellRange(0, this._colstart, 0, this._colstart + this._colspan - 1);
                };
                // get the preferred column width for a column in the group
                _CellGroup.prototype.getColumnWidth = function (c) {
                    for (var i = 0; i < this._cells.length; i++) {
                        var cell = this._cells[i];
                        if (cell._col == c && cell.colspan == 1) {
                            return cell.width;
                        }
                    }
                    return null;
                };
                // get merged range for a cell in this group
                _CellGroup.prototype.getMergedRange = function (p, r, c) {
                    // merged column header range
                    if (r < 0) {
                        return this._rng[-1];
                    }
                    // regular cell range
                    var row = p.rows[r], rs = row.recordIndex != null ? row.recordIndex : r % this._rowspan, cs = c - this._colstart, rng = this._rng[rs * this._colspan + cs];
                    return rng
                        ? new grid_1.CellRange(r + rng.row, c + rng.col, r + rng.row2, c + rng.col2)
                        : null;
                };
                // get the binding column for a cell in this group
                _CellGroup.prototype.getBindingColumn = function (p, r, c) {
                    // merged column header binding
                    // return 'this' to render the collapsed column header
                    if (r < 0) {
                        return this;
                    }
                    // regular cells
                    var row = p.rows[r], rs = row.recordIndex != null ? row.recordIndex : r % this._rowspan, cs = c - this._colstart;
                    return this._cols[rs * this._colspan + cs];
                };
                return _CellGroup;
            }(multirow._Cell));
            multirow._CellGroup = _CellGroup;
        })(multirow = grid_1.multirow || (grid_1.multirow = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_CellGroup.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid_1) {
        var multirow;
        (function (multirow) {
            'use strict';
            /**
             * Provides custom merging for @see:MultiRow controls.
             */
            var _MergeManager = (function (_super) {
                __extends(_MergeManager, _super);
                function _MergeManager() {
                    _super.apply(this, arguments);
                }
                /**
                 * Gets a @see:CellRange that specifies the merged extent of a cell
                 * in a @see:GridPanel.
                 *
                 * @param p The @see:GridPanel that contains the range.
                 * @param r The index of the row that contains the cell.
                 * @param c The index of the column that contains the cell.
                 * @param clip Specifies whether to clip the merged range to the grid's current view range.
                 * @return A @see:CellRange that specifies the merged range, or null if the cell is not merged.
                 */
                _MergeManager.prototype.getMergedRange = function (p, r, c, clip) {
                    if (clip === void 0) { clip = true; }
                    var grid = p.grid;
                    // handle group rows
                    switch (p.cellType) {
                        case grid_1.CellType.Cell:
                        case grid_1.CellType.RowHeader:
                            if (p.rows[r] instanceof grid_1.GroupRow) {
                                return _super.prototype.getMergedRange.call(this, p, r, c, clip);
                            }
                    }
                    // other cells
                    switch (p.cellType) {
                        // merge cells in cells and column headers panels
                        case grid_1.CellType.Cell:
                        case grid_1.CellType.ColumnHeader:
                            // get the group range
                            var group = grid._cellGroupsByColumn[c];
                            wijmo.assert(group instanceof multirow._CellGroup, 'Failed to get the group!');
                            if (p.cellType == grid_1.CellType.ColumnHeader && grid.collapsedHeaders) {
                                r = -1; // handle collapsed headers
                            }
                            var rng = group.getMergedRange(p, r, c);
                            // prevent merging across frozen column boundary (TFS 192385)
                            if (rng && p.columns.frozen) {
                                var frz = p.columns.frozen;
                                if (rng.col < frz && rng.col2 >= frz) {
                                    if (c < frz) {
                                        rng.col2 = frz - 1;
                                    }
                                    else {
                                        rng.col = frz;
                                    }
                                }
                            }
                            // prevent merging across frozen row boundary (TFS 192385)
                            if (rng && p.rows.frozen) {
                                var frz = p.rows.frozen;
                                if (rng.row < frz && rng.row2 >= frz) {
                                    if (r < frz) {
                                        rng.row2 = frz - 1;
                                    }
                                    else {
                                        rng.row = frz;
                                    }
                                }
                            }
                            // return the range
                            return rng; //group.getMergedRange(p, r, c);
                        // merge cells in row headers panel
                        case grid_1.CellType.RowHeader:
                            var rpi = grid._rowsPerItem, row = p.rows[r], top = r - row.recordIndex;
                            return new grid_1.CellRange(top, 0, top + rpi - 1, p.columns.length - 1);
                        // merge cells in top/left cell
                        case grid_1.CellType.TopLeft:
                            return new grid_1.CellRange(0, 0, p.rows.length - 1, p.columns.length - 1);
                    }
                    // no merging
                    return null;
                };
                return _MergeManager;
            }(grid_1.MergeManager));
            multirow._MergeManager = _MergeManager;
        })(multirow = grid_1.multirow || (grid_1.multirow = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_MergeManager.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid_1) {
        var multirow;
        (function (multirow) {
            'use strict';
            /**
             * Manages the new row template used to add rows to the grid.
             */
            var _AddNewHandler = (function (_super) {
                __extends(_AddNewHandler, _super);
                /**
                 * Initializes a new instance of the @see:_AddNewHandler class.
                 *
                 * @param grid @see:FlexGrid that owns this @see:_AddNewHandler.
                 */
                function _AddNewHandler(grid) {
                    // detach old handler
                    var old = grid._addHdl;
                    old._detach();
                    // attach this handler instead
                    _super.call(this, grid);
                }
                /**
                 * Updates the new row template to ensure that it is visible only when the grid is
                 * bound to a data source that supports adding new items, and that it is
                 * in the right position.
                 */
                _AddNewHandler.prototype.updateNewRowTemplate = function () {
                    // get variables
                    var ecv = wijmo.tryCast(this._g.collectionView, 'IEditableCollectionView'), g = this._g, rows = g.rows;
                    // see if we need a new row template
                    var needTemplate = ecv && ecv.canAddNew && g.allowAddNew && !g.isReadOnly;
                    // see if we have new row template
                    var hasTemplate = true;
                    for (var i = rows.length - g.rowsPerItem; i < rows.length; i++) {
                        if (!(rows[i] instanceof _NewRowTemplate)) {
                            hasTemplate = false;
                            break;
                        }
                    }
                    // add template
                    if (needTemplate && !hasTemplate) {
                        for (var i = 0; i < g.rowsPerItem; i++) {
                            var nrt = new _NewRowTemplate(i);
                            rows.push(nrt);
                        }
                    }
                    // remove template
                    if (!needTemplate && hasTemplate) {
                        for (var i = 0; i < rows.length; i++) {
                            if (rows[i] instanceof _NewRowTemplate) {
                                rows.removeAt(i);
                                i--;
                            }
                        }
                    }
                };
                return _AddNewHandler;
            }(wijmo.grid._AddNewHandler));
            multirow._AddNewHandler = _AddNewHandler;
            /**
             * Represents a row template used to add items to the source collection.
             */
            var _NewRowTemplate = (function (_super) {
                __extends(_NewRowTemplate, _super);
                function _NewRowTemplate(indexInRecord) {
                    _super.call(this);
                    this._idxRecord = indexInRecord;
                }
                Object.defineProperty(_NewRowTemplate.prototype, "recordIndex", {
                    get: function () {
                        return this._idxRecord;
                    },
                    enumerable: true,
                    configurable: true
                });
                return _NewRowTemplate;
            }(wijmo.grid._NewRowTemplate));
        })(multirow = grid_1.multirow || (grid_1.multirow = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_AddNewHandler.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
/**
 * Defines the @see:MultiRow control and its associated classes.
 */
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var multirow;
        (function (multirow) {
            'use strict';
            /**
             * Extends the @see:FlexGrid control to provide multiple rows per item.
             *
             * Use the <b>layoutDefinition</b> property to define the layout of the rows
             * used to display each data item.
             *
             * A few @see:FlexGrid properties are disabled in the @see:MultiRow control
             * because they would interfere with the custom multi-row layouts.
             * The list of disabled properties includes @see:FlexGrid.allowMerging and
             * @see:FlexGrid.childItemsPath.
             */
            var MultiRow = (function (_super) {
                __extends(MultiRow, _super);
                /**
                 * Initializes a new instance of the @see:MultiRow class.
                 *
                 * In most cases, the <b>options</b> parameter will include the value for the
                 * @see:layoutDefinition property.
                 *
                 * @param element The DOM element that will host the control, or a selector for the host element (e.g. '#theCtrl').
                 * @param options JavaScript object containing initialization data for the control.
                 */
                function MultiRow(element, options) {
                    var _this = this;
                    _super.call(this, element);
                    this._rowsPerItem = 1;
                    this._cellBindingGroups = [];
                    this._centerVert = true;
                    this._collapsedHeaders = false;
                    // add class name to enable styling
                    wijmo.addClass(this.hostElement, 'wj-multirow');
                    // add header collapse/expand button
                    var hdr = this.columnHeaders.hostElement.parentElement, btn = wijmo.createElement('<div class="wj-hdr-collapse"><span></span></div>');
                    btn.style.display = 'none';
                    hdr.appendChild(btn);
                    this._btnCollapse = btn;
                    this._updateButtonGlyph();
                    // handle mousedown on collapse/expand button (not click: TFS 190572)
                    this.addEventListener(btn, 'mousedown', function (e) {
                        _this.collapsedHeaders = !_this.collapsedHeaders;
                        e.preventDefault();
                    }, true);
                    // change some defaults
                    this.autoGenerateColumns = false;
                    this.allowDragging = grid.AllowDragging.None;
                    this.mergeManager = new multirow._MergeManager(this);
                    // custom AddNewHandler
                    this._addHdl = new multirow._AddNewHandler(this);
                    // customize cell rendering
                    this.formatItem.addHandler(this._formatItem, this);
                    // select multi-row items when clicking the row headers
                    this.addEventListener(this.rowHeaders.hostElement, 'click', function (e) {
                        var ht = _this.hitTest(e);
                        if (ht.panel == _this.rowHeaders && ht.row > -1) {
                            var row = _this.rows[ht.row];
                            if (row.recordIndex != null) {
                                var top = row.index - row.recordIndex;
                                _this.select(new grid.CellRange(top, 0, top + _this.rowsPerItem - 1, _this.columns.length - 1));
                            }
                        }
                    });
                    // apply options after everything else is ready
                    this.initialize(options);
                }
                Object.defineProperty(MultiRow.prototype, "layoutDefinition", {
                    /**
                     * Gets or sets an array that defines the layout of the rows used to display each data item.
                     *
                     * The array contains a list of cell group objects which have the following properties:
                     *
                     * <ul>
                     * <li><b>header</b>: Group header (shown when the headers are collapsed)</li>
                     * <li><b>colspan</b>: Number of grid columns spanned by the group</li>
                     * <li><b>cells</b>: Array of cell objects, which extend @see:Column with a <b>colspan</b> property.</li>
                     * </ul>
                     *
                     * When the @see:layoutDefinition property is set, the grid scans the cells in each
                     * group as follows:
                     *
                     * <ol>
                     * <li>The grid calculates the colspan of the group either as group's own colspan
                     * or as span of the widest cell in the group, whichever is wider.</li>
                     * <li>If the cell fits the current row within the group, it is added to the current row.</li>
                     * <li>If it doesn't fit, it is added to a new row.</li>
                     * </ol>
                     *
                     * When all groups are ready, the grid calculates the number of rows per record to the maximum
                     * rowspan of all groups, and adds rows to each group to pad their height as needed.
                     *
                     * This scheme is simple and flexible. For example:
                     * <pre>{ header: 'Group 1', cells: [{ binding: 'c1' }, { bnding: 'c2'}, { binding: 'c3' }]}</pre>
                     *
                     * The group has colspan 1, so there will be one cell per column. The result is:
                     * <pre>
                     * | C1 |
                     * | C2 |
                     * | C3 |
                     * </pre>
                     *
                     * To create a group with two columns, set <b>colspan</b> property of the group:
                     *
                     * <pre>{ header: 'Group 1', colspan: 2, cells:[{ binding: 'c1' }, { binding: 'c2'}, { binding: 'c3' }]}</pre>
                     *
                     * The cells will wrap as follows:
                     * <pre>
                     * | C1  | C2 |
                     * | C3       |
                     * </pre>
                     *
                     * Note that the last cell spans two columns (to fill the group).
                     *
                     * You can also specify the colspan on individual cells rather than on the group:
                     *
                     * <pre>{ header: 'Group 1', cells: [{binding: 'c1', colspan: 2 }, { bnding: 'c2'}, { binding: 'c3' }]}</pre>
                     *
                     * Now the first cell has colspan 2, so the result is:
                     * <pre>
                     * | C1       |
                     * | C2 |  C3 |
                     * </pre>
                     *
                     * Because cells extend the @see:Column class, you can add all the usual @see:Column
                     * properties to any cells:
                     * <pre>
                     * { header: 'Group 1', cells: [
                     *    { binding: 'c1', colspan: 2 },
                     *    { bnding: 'c2'},
                     *    { binding: 'c3', format: 'n0', required: false, etc... }
                     * ]}</pre>
                     */
                    get: function () {
                        return this._layoutDef;
                    },
                    set: function (value) {
                        // store original value so user can get it back
                        this._layoutDef = wijmo.asArray(value);
                        // parse cell bindings
                        this._rowsPerItem = 1;
                        this._cellBindingGroups = this._parseCellGroups(this._layoutDef);
                        for (var i = 0; i < this._cellBindingGroups.length; i++) {
                            var group = this._cellBindingGroups[i];
                            this._rowsPerItem = Math.max(this._rowsPerItem, group._rowspan);
                        }
                        // go bind/rebind the grid
                        this._bindGrid(true);
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(MultiRow.prototype, "rowsPerItem", {
                    /**
                     * Gets the number of rows used to display each item.
                     *
                     * This value is calculated automatically based on the value
                     * of the <b>layoutDefinition</b> property.
                     */
                    get: function () {
                        return this._rowsPerItem;
                    },
                    enumerable: true,
                    configurable: true
                });
                /**
                 * Gets the @see:Column object used to bind a data item to a grid cell.
                 *
                 * @param p @see:GridPanel that contains the cell.
                 * @param r Index of the row that contains the cell.
                 * @param c Index of the column that contains the cell.
                 */
                MultiRow.prototype.getBindingColumn = function (p, r, c) {
                    return this._getBindingColumn(p, r, p.columns[c]);
                };
                Object.defineProperty(MultiRow.prototype, "centerHeadersVertically", {
                    /**
                     * Gets or sets a value that determines whether the content of cells
                     * that span multiple rows should be vertically centered.
                     */
                    get: function () {
                        return this._centerVert;
                    },
                    set: function (value) {
                        if (value != this._centerVert) {
                            this._centerVert = wijmo.asBoolean(value);
                            this.invalidate();
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(MultiRow.prototype, "collapsedHeaders", {
                    /**
                     * Gets or sets a value that determines whether column headers
                     * should be collapsed and displayed as a single row displaying
                     * the group headers.
                     *
                     * If you set the <b>collapsedHeaders</b> property to true,
                     * remember to set the <b>header</b> property of every group in order
                     * to avoid any empty headers.
                     */
                    get: function () {
                        return this._collapsedHeaders;
                    },
                    set: function (value) {
                        if (value != this._collapsedHeaders) {
                            this._collapsedHeaders = wijmo.asBoolean(value);
                            this._updateButtonGlyph();
                            this._bindGrid(true);
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(MultiRow.prototype, "showHeaderCollapseButton", {
                    /**
                     * Gets or sets a value that determines whether the grid should display
                     * a button in the column header panel to allow users to collapse and
                     * expand the column headers.
                     *
                     * If the button is visible, clicking on it will cause the grid to
                     * toggle the value of the <b>collapsedHeaders</b> property.
                     */
                    get: function () {
                        return this._btnCollapse.style.display == '';
                    },
                    set: function (value) {
                        if (value != this.showHeaderCollapseButton) {
                            this._btnCollapse.style.display = wijmo.asBoolean(value) ? '' : 'none';
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                // ** overrides
                // bind rows
                /*protected*/ MultiRow.prototype._addBoundRow = function (items, index) {
                    var item = items[index];
                    for (var i = 0; i < this._rowsPerItem; i++) {
                        this.rows.push(new multirow._MultiRow(item, index, i));
                    }
                };
                /*protected*/ MultiRow.prototype._addNode = function (items, index, level) {
                    this._addBoundRow(items, index); // childItemsPath not supported
                };
                // bind columns
                /*protected*/ MultiRow.prototype._bindColumns = function () {
                    // update column header row count
                    var rows = this.columnHeaders.rows, cnt = this._collapsedHeaders ? 1 : this._rowsPerItem;
                    while (rows.length > cnt) {
                        rows.removeAt(rows.length - 1);
                    }
                    while (rows.length < cnt) {
                        rows.push(new grid.Row());
                    }
                    // remove old columns
                    this.columns.clear();
                    this._cellGroupsByColumn = {};
                    // get first item to infer data types
                    var item = null, cv = this.collectionView;
                    if (cv && cv.sourceCollection && cv.sourceCollection.length) {
                        item = cv.sourceCollection[0];
                    }
                    // generate columns
                    if (this._cellBindingGroups) {
                        for (var i = 0; i < this._cellBindingGroups.length; i++) {
                            var group = this._cellBindingGroups[i];
                            for (var c = 0; c < group._colspan; c++) {
                                this._cellGroupsByColumn[this.columns.length] = group;
                                var col = new grid.Column();
                                col.width = group.getColumnWidth(c);
                                this.columns.push(col);
                            }
                        }
                    }
                };
                // update missing column types to match data
                /*protected*/ MultiRow.prototype._updateColumnTypes = function () {
                    // allow base class
                    _super.prototype._updateColumnTypes.call(this);
                    // update missing column types in all binding groups
                    var cv = this.collectionView;
                    if (wijmo.hasItems(cv)) {
                        var item = cv.items[0];
                        for (var i = 0; i < this._cellBindingGroups.length; i++) {
                            var group = this._cellBindingGroups[i];
                            for (var c = 0; c < group._cols.length; c++) {
                                var col = group._cols[c];
                                if (col.dataType == null && col._binding) {
                                    col.dataType = wijmo.getType(col._binding.getValue(item));
                                }
                            }
                        }
                    }
                };
                // get the binding column 
                // (in the MultiRow grid, each physical column may contain several binding columns)
                /*protected*/ MultiRow.prototype._getBindingColumn = function (p, r, c) {
                    // convert column to binding column (cell)
                    if (p == this.cells || p == this.columnHeaders) {
                        var group = this._cellGroupsByColumn[c.index];
                        if (p == this.columnHeaders && this.collapsedHeaders) {
                            r = -1; // handle collapsed headers
                        }
                        c = group.getBindingColumn(p, r, c.index);
                    }
                    // done
                    return c;
                };
                // update grid rows to sync with data source
                /*protected*/ MultiRow.prototype._cvCollectionChanged = function (sender, e) {
                    if (this.autoGenerateColumns && this.columns.length == 0) {
                        this._bindGrid(true);
                    }
                    else {
                        switch (e.action) {
                            // item changes don't require re-binding
                            case wijmo.collections.NotifyCollectionChangedAction.Change:
                                this.invalidate();
                                break;
                            // always add at the bottom (TFS 193086)
                            case wijmo.collections.NotifyCollectionChangedAction.Add:
                                if (e.index == this.collectionView.items.length - 1) {
                                    var index = this.rows.length;
                                    while (index > 0 && this.rows[index - 1] instanceof grid._NewRowTemplate) {
                                        index--;
                                    }
                                    for (var i = 0; i < this._rowsPerItem; i++) {
                                        this.rows.insert(index + i, new multirow._MultiRow(e.item, e.index, i));
                                    }
                                    return;
                                }
                                wijmo.assert(false, 'added item should be the last one.');
                                break;
                            // remove/refresh require re-binding
                            default:
                                this._bindGrid(false);
                                break;
                        }
                    }
                };
                // ** implementation
                // parse an array of JavaScript objects into an array of _BindingGroup objects
                MultiRow.prototype._parseCellGroups = function (groups) {
                    var arr = [], rowsPerItem = 1;
                    if (groups) {
                        // parse binding groups
                        for (var i = 0, colstart = 0; i < groups.length; i++) {
                            var group = new multirow._CellGroup(this, groups[i]);
                            group._colstart = colstart;
                            colstart += group._colspan;
                            rowsPerItem = Math.max(rowsPerItem, group._rowspan);
                            arr.push(group);
                        }
                        // close binding groups (calculate group's rowspan, ranges, and bindings)
                        for (var i = 0; i < arr.length; i++) {
                            arr[i].closeGroup(rowsPerItem);
                        }
                    }
                    return arr;
                };
                // customize cells
                MultiRow.prototype._formatItem = function (s, e) {
                    var rpi = this._rowsPerItem, row = e.panel.rows[e.range.row], row2 = e.panel.rows[e.range.row2];
                    // add group start/end class markers
                    if (e.panel.cellType == grid.CellType.Cell || e.panel.cellType == grid.CellType.ColumnHeader) {
                        var group = this._cellGroupsByColumn[e.col];
                        wijmo.assert(group instanceof multirow._CellGroup, 'Failed to get the group!');
                        wijmo.toggleClass(e.cell, 'wj-group-start', group._colstart == e.range.col);
                        wijmo.toggleClass(e.cell, 'wj-group-end', group._colstart + group._colspan - 1 == e.range.col2);
                    }
                    // add item start/end class markers
                    if (rpi > 1) {
                        if (e.panel.cellType == grid.CellType.Cell || e.panel.cellType == grid.CellType.RowHeader) {
                            wijmo.toggleClass(e.cell, 'wj-record-start', row instanceof multirow._MultiRow ? row.recordIndex == 0 : false);
                            wijmo.toggleClass(e.cell, 'wj-record-end', row2 instanceof multirow._MultiRow ? row2.recordIndex == rpi - 1 : false);
                        }
                    }
                    // handle alternating rows
                    if (this.showAlternatingRows) {
                        wijmo.toggleClass(e.cell, 'wj-alt', row instanceof multirow._MultiRow ? row.dataIndex % 2 != 0 : false);
                    }
                    // center-align cells vertically if they span multiple rows
                    if (this._centerVert) {
                        if (e.cell.hasChildNodes && e.range.rowSpan > 1) {
                            // surround cell content in a vertically centered table-cell div
                            var div = wijmo.createElement('<div style="display:table-cell;vertical-align:middle"></div>'), rng = document.createRange();
                            rng.selectNodeContents(e.cell);
                            rng.surroundContents(div);
                            // make the cell display as a table
                            wijmo.setCss(e.cell, {
                                display: 'table',
                                tableLayout: 'fixed',
                                paddingTop: 0,
                                paddingBottom: 0
                            });
                        }
                        else {
                            wijmo.setCss(e.cell, {
                                display: '',
                                tableLayout: '',
                                paddingTop: '',
                                paddingBottom: ''
                            });
                        }
                    }
                };
                // update glyph in collapse/expand headers button
                MultiRow.prototype._updateButtonGlyph = function () {
                    var span = this._btnCollapse.querySelector('span');
                    if (span instanceof HTMLElement) {
                        span.className = this.collapsedHeaders ? 'wj-glyph-left' : 'wj-glyph-down-left';
                    }
                };
                return MultiRow;
            }(grid.FlexGrid));
            multirow.MultiRow = MultiRow;
        })(multirow = grid.multirow || (grid.multirow = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=MultiRow.js.map
