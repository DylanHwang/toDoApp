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
// enable use of EcmaScript6 maps
//declare var Map: any;

// initialize groupHeaderFormat
wijmo.culture.FlexGrid = {
    groupHeaderFormat: '{name}: <b>{value} </b>({count:n0} items)'
};

/**
 * Defines the @see:FlexGrid control and associated classes.
 *
 * The example below creates a @see:FlexGrid control and binds it to a
 * 'data' array. The grid has four columns, specified by explicitly
 * populating the grid's @see:FlexGrid.columns array.
 *
 * @fiddle:6GB66
 */
module wijmo.grid {
    'use strict';

    /**
     * Specifies constants that define the visibility of row and column headers.
     */
    export enum HeadersVisibility {
        /** No header cells are displayed. */
        None = 0,
        /** Only column header cells are displayed. */
        Column = 1,
        /** Only row header cells are displayed. */
        Row = 2,
        /** Both column and row header cells are displayed. */
        All = 3,
    }

    /**
     * The @see:FlexGrid control provides a powerful and flexible way to
     * display and edit data in a tabular format.
     *
     * The @see:FlexGrid control is a full-featured grid, providing all the
     * features you are used to including several selection modes, sorting,
     * column reordering, grouping, filtering, editing, custom cells,
     * XAML-style star-sizing columns, row and column virtualization, etc.
     */
    export class FlexGrid extends Control {

        // constants
        static _WJS_STICKY = 'wj-state-sticky';
        static _WJS_MEASURE = 'wj-state-measuring';

        // child elements
        private _root: HTMLDivElement;
        private _eCt: HTMLDivElement;
        private _fCt: HTMLDivElement;
        /*private*/ _eTL: HTMLDivElement; // visible to HitTestInfo
        /*private*/ _eBL: HTMLDivElement; // visible to HitTestInfo
        private _eCHdr: HTMLDivElement;
        private _eCFtr: HTMLDivElement;
        private _eRHdr: HTMLDivElement;
        private _eCHdrCt: HTMLDivElement;
        private _eCFtrCt: HTMLDivElement;
        private _eRHdrCt: HTMLDivElement;
        private _eTLCt: HTMLDivElement;
        private _eBLCt: HTMLDivElement;
        private _eSz: HTMLDivElement;
        private _eMarquee: HTMLDivElement;
        private _eFocus: HTMLDivElement;

        // child panels
        private _gpTL: GridPanel;
        private _gpCHdr: GridPanel;
        private _gpRHdr: GridPanel;
        private _gpCells: GridPanel;
        private _gpBL: GridPanel;
        private _gpCFtr: GridPanel;

        // private stuff
        private _maxOffsetY: number;
        private _heightBrowser: number;
        /*private*/ _szClient = new Size(0, 0);
        /*private*/ _offsetY: number;
        /*private*/ _lastCount: number;
        /*private*/ _rcBounds: Rect; // accessible to MouseHandler
        /*private*/ _ptScrl = new Point(0, 0); // accessible to GridPanel
        /*private*/ _rtl = false; // accessible to CellFactory, GridFilter
        /*private*/ _cellPadding = 3; // accessible to CellFactory

        // selection/key/mouse handlers
        /*private*/_mouseHdl: _MouseHandler;    // accessible to keyboard handler
        /*private*/_edtHdl: _EditHandler;       // accessible to keyboard handler
        /*private*/_selHdl: _SelectionHandler;  // accessible to keyboard handler
        /*private*/_addHdl: _AddNewHandler;     // accessible to derived classes (e.g. MultiRow)
        /*private*/_keyHdl: _KeyboardHandler;   // accessible to IME handler
        private _imeHdl: _ImeHandler;
        private _mrgMgr: MergeManager;

        // property storage
        private _autoGenCols = true;
        private _autoClipboard = true;
        private _readOnly = false;
        private _indent = 14;
        private _autoSizeMode = AutoSizeMode.Both;
        private _hdrVis = HeadersVisibility.All;
        private _alSorting = true;
        private _alAddNew = false;
        private _alDelete = false;
        private _alResizing = AllowResizing.Columns;
        private _alDragging = AllowDragging.Columns;
        private _alMerging = AllowMerging.None;
        private _ssHdr = HeadersVisibility.None;
        private _shSort = true;
        private _shGroups = true;
        private _shAlt = true;
        private _gHdrFmt: string;
        private _rows: RowCollection;
        private _cols: ColumnCollection;
        private _hdrRows: RowCollection;
        private _ftrRows: RowCollection;
        private _hdrCols: ColumnCollection;
        private _cf: CellFactory;
        private _itemFormatter: Function;
        private _items: any; // any[] or ICollectionView
        private _cv: collections.ICollectionView;
        private _childItemsPath: any;
        private _sortRowIndex: number;
        private _deferResizing = false;
        private _bndSortConverter;
        private _afScrl: number;
        private _stickyHdr: boolean;
        private _toSticky: number;
        private _pSel = true;
        private _pOutline = true;

        /**
         * Gets or sets the template used to instantiate @see:FlexGrid controls.
         */
        static controlTemplate = '<div style="position:relative;width:100%;height:100%;overflow:hidden;max-width:inherit;max-height:inherit">' +
          '<div wj-part="focus" tabIndex="0" style="position:absolute;opacity:0;pointer-events:none;left:-10px;top:-10px"></div>' +
          '<div wj-part="root" style="position:absolute;width:100%;height:100%;overflow:auto;-webkit-overflow-scrolling:touch;max-width:inherit;max-height:inherit;boxSizing:content-box">' + // cell container
            '<div wj-part="cells" class="wj-cells" style="position:absolute"></div>' + // cells
            '<div wj-part="marquee" class="wj-marquee" style="display:none;pointer-events:none">' +  // marquee (no pointer events!!!!)
              '<div style="width:100%;height:100%"></div>' +
            '</div>' +
          '</div>' +
          '<div wj-part="fcells" style="position:absolute;pointer-events:none;overflow:hidden"></div>' + // frozen cells
          '<div wj-part="rh" style="position:absolute;overflow:hidden;outline:none">' + // row header container
            '<div wj-part="rhcells" class="wj-rowheaders" style="position:relative"></div>' + // row header cells
          '</div>' +
          '<div wj-part="cf" style="position:absolute;overflow:hidden;outline:none">' + // col footer container
            '<div wj-part="cfcells" class="wj-colfooters" style="position:relative"></div>' + // col footer cells
          '</div>' +
          '<div wj-part="ch" style="position:absolute;overflow:hidden;outline:none">' + // col header container
            '<div wj-part="chcells" class="wj-colheaders" style="position:relative"></div>' + // col header cells
          '</div>' +
          '<div wj-part="bl" style="position:absolute;overflow:hidden;outline:none">' + // bottom-left container
            '<div wj-part="blcells" class="wj-bottomleft" style="position:relative"></div>' + // top-left cells
          '</div>' +
          '<div wj-part="tl" style="position:absolute;overflow:hidden;outline:none">' + // top-left container
            '<div wj-part="tlcells" class="wj-topleft" style="position:relative"></div>' + // top-left cells
          '</div>' +
          '<div wj-part="sz" style="position:relative;visibility:hidden"></div>' + // auto sizing
        '</div>';

        /**
         * Initializes a new instance of the @see:FlexGrid class.
         *
         * @param element The DOM element that will host the control, or a selector for the host element (e.g. '#theCtrl').
         * @param options JavaScript object containing initialization data for the control.
         */
        constructor(element: any, options?) {
            super(element, null, true);
            var e = this.hostElement;

            // make sure we have no border radius if the browser is IE/Edge
            // (rounded borders **kill** scrolling perf!!!!)
            if (isIE()) {
                e.style.borderRadius = '0px';
            }

            // instantiate and apply template
            var tpl = this.getTemplate();
            this.applyTemplate('wj-control wj-flexgrid wj-content', tpl, {
                _root: 'root',
                _eSz: 'sz',
                _eCt: 'cells',
                _fCt: 'fcells',
                _eTL: 'tl',
                _eBL: 'bl',
                _eCHdr: 'ch',
                _eRHdr: 'rh',
                _eCFtr: 'cf',
                _eTLCt: 'tlcells',
                _eBLCt: 'blcells',
                _eCHdrCt: 'chcells',
                _eCFtrCt: 'cfcells',
                _eRHdrCt: 'rhcells',
                _eMarquee: 'marquee',
                _eFocus: 'focus'
            });

            // calculate default row height
            var defRowHei = this._getDefaultRowHeight();

            // build the control
            this.deferUpdate(() => {

                // create row and column collections
                this._rows = new RowCollection(this, defRowHei);
                this._cols = new ColumnCollection(this, defRowHei * 4);
                this._hdrRows = new RowCollection(this, defRowHei);
                this._hdrCols = new ColumnCollection(this, Math.round(defRowHei * 1.25));
                this._ftrRows = new RowCollection(this, defRowHei);

                // create grid panels
                this._gpTL = new GridPanel(this, CellType.TopLeft, this._hdrRows, this._hdrCols, this._eTLCt);
                this._gpCHdr = new GridPanel(this, CellType.ColumnHeader, this._hdrRows, this._cols, this._eCHdrCt);
                this._gpRHdr = new GridPanel(this, CellType.RowHeader, this._rows, this._hdrCols, this._eRHdrCt);
                this._gpCells = new GridPanel(this, CellType.Cell, this._rows, this._cols, this._eCt);
                this._gpBL = new GridPanel(this, CellType.BottomLeft, this._ftrRows, this._hdrCols, this._eBLCt);
                this._gpCFtr = new GridPanel(this, CellType.ColumnFooter, this._ftrRows, this._cols, this._eCFtrCt);

                // add row and column headers
                this._hdrRows.push(new Row());
                this._hdrCols.push(new Column());
                this._hdrCols[0].align = 'center';

                // initialize control
                this._cf = new CellFactory();
                this._keyHdl = new _KeyboardHandler(this);
                this._mouseHdl = new _MouseHandler(this);
                this._edtHdl = new _EditHandler(this);
                this._selHdl = new _SelectionHandler(this);
                this._addHdl = new _AddNewHandler(this);
                this._mrgMgr = new MergeManager(this);
                this._bndSortConverter = this._sortConverter.bind(this);

                // apply options after grid has been initialized
                this.initialize(options);
            });

            // update content when user scrolls the control
            this.addEventListener(this._root, 'scroll', (e) => {
                if (this._updateScrollPosition()) { // TFS 150650
                    if (this._afScrl) {
                        cancelAnimationFrame(this._afScrl);
                    }
                    this._afScrl = requestAnimationFrame(() => {
                        this.finishEditing();
                        this._updateContent(true);
                        this._afScrl = null;
                    });
                }
            });
        }

        // reset rcBounds when window is resized
        // (even if the control size didn't change, because it may have moved: TFS 112961)
        _handleResize() {
            super._handleResize();
            this._rcBounds = null;
        }

        //--------------------------------------------------------------------------
        //#region ** object model

        /** 
         * Gets or sets a value that determines whether the row and column headers
         * are visible.
         */
        get headersVisibility(): HeadersVisibility {
            return this._hdrVis;
        }
        set headersVisibility(value: HeadersVisibility) {
            if (value != this._hdrVis) {
                this._hdrVis = asEnum(value, HeadersVisibility);
                this.invalidate();
            }
        }
        /**
         * Gets or sets a value that determines whether column headers should remain
         * visible when the user scrolls the document.
         */
        get stickyHeaders(): boolean {
            return this._stickyHdr;
        }
        set stickyHeaders(value: boolean) {
            if (value != this._stickyHdr) {

                // save new setting
                this._stickyHdr = asBoolean(value);
                this._updateStickyHeaders();

                // attach/detach scroll handler to update the sticky headers (with capture)
                this.removeEventListener(window, 'scroll');
                if (this._stickyHdr) {
                    this.addEventListener(window, 'scroll', (e) => {
                        if (contains(e.target, this.hostElement)) {
                            if (this._toSticky) {
                                cancelAnimationFrame(this._toSticky);
                            }
                            this._toSticky = requestAnimationFrame(() => {
                                var e = new CancelEventArgs();
                                if (this.onUpdatingLayout(e)) {
                                    this._updateStickyHeaders();
                                    this.onUpdatedLayout(e);
                                }
                                this._toSticky = null;
                            });
                        }
                    }, true);
                }
            }
        }
        /**
         * Gets or sets a value that determines whether the grid should preserve
         * the selected state of rows when the data is refreshed.
         */
        get preserveSelectedState(): boolean {
            return this._pSel;
        }
        set preserveSelectedState(value: boolean) {
            this._pSel = asBoolean(value);
        }
        /**
         * Gets or sets a value that determines whether the grid should preserve
         * the expanded/collapsed state of nodes when the data is refreshed.
         *
         * The @see:preserveOutlineState property implementation is based on
         * JavaScript's @see:Map object, which is not available in IE 9 or 10.
         */
        get preserveOutlineState(): boolean {
            return this._pOutline;
        }
        set preserveOutlineState(value: boolean) {
            this._pOutline = asBoolean(value);
        }

        /**
         * Gets or sets a value that determines whether the grid should generate columns 
         * automatically based on the @see:itemsSource.
         *
         * The column generation depends on the @see:itemsSource property containing
         * at least one item. This data item is inspected and a column is created and
         * bound to each property that contains a primitive value (number, string,
         * Boolean, or Date).
         *
         * Properties set to null do not generate columns, because the grid would
         * have no way of guessing the appropriate type. In this type of scenario,
         * you should set the @see:autoGenerateColumns property to false and create
         * the columns explicitly. For example:
         *
         * <pre>var grid = new wijmo.grid.FlexGrid('#theGrid', {
         *   autoGenerateColumns: false, // data items may contain null values
         *   columns: [                  // so define columns explicitly
         *     { binding: 'name', header: 'Name', type: 'String' },
         *     { binding: 'amount', header: 'Amount', type: 'Number' },
         *     { binding: 'date', header: 'Date', type: 'Date' },
         *     { binding: 'active', header: 'Active', type: 'Boolean' }
         *   ],
         *   itemsSource: customers
         * });</pre>
         */
        get autoGenerateColumns(): boolean {
            return this._autoGenCols;
        }
        set autoGenerateColumns(value: boolean) {
            this._autoGenCols = asBoolean(value);
        }
        /**
         * Gets or sets a value that determines whether the grid should handle 
         * clipboard shortcuts.
         *
         * The clipboard shortcuts are as follows:
         *
         * <dl class="dl-horizontal">
         *   <dt>ctrl+C, ctrl+Ins</dt>    <dd>Copy grid selection to clipboard.</dd>
         *   <dt>ctrl+V, shift+Ins</dt>   <dd>Paste clipboard text to grid selection.</dd>
         * </dl>
         *
         * Only visible rows and columns are included in clipboard operations.
         *
         * Read-only cells are not affected by paste operations.
         */
        get autoClipboard(): boolean {
            return this._autoClipboard;
        }
        set autoClipboard(value: boolean) {
            this._autoClipboard = asBoolean(value);
        }
        /**
         * Gets or sets a JSON string that defines the current column layout.
         *
         * The column layout string represents an array with the columns and their
         * properties. It can be used to persist column layouts defined by users so 
         * they are preserved across sessions, and can also be used to implement undo/redo
         * functionality in applications that allow users to modify the column layout.
         *
         * The column layout string does not include <b>dataMap</b> properties, because
         * data maps are not serializable.
         */
        get columnLayout(): string {
            var props = FlexGrid._getSerializableProperties(Column),
                defs = new Column(),
                proxyCols = [];

            // populate array with proxy columns
            // save only primitive value and non-default settings
            // don't save 'size', we are already saving 'width'
            for (var i = 0; i < this.columns.length; i++) {
                var col = this.columns[i],
                    proxyCol = {};
                for (var j = 0; j < props.length; j++) {
                    var prop = props[j],
                        value = col[prop];
                    if (value != defs[prop] && isPrimitive(value) && prop != 'size') {
                        proxyCol[prop] = value;
                    }
                }
                proxyCols.push(proxyCol)
            }

            // return JSON string with proxy columns
            return JSON.stringify({ columns: proxyCols });
        }
        set columnLayout(value: string) {
            var colOptions = JSON.parse(asString(value));
            if (!colOptions || colOptions.columns == null) {
                throw 'Invalid columnLayout data.';
            }
            this.columns.clear();
            this.initialize(colOptions);
        }
        /**
         * Gets or sets a value that determines whether the user can modify
         * cell values using the mouse and keyboard.
         */
        get isReadOnly(): boolean {
            return this._readOnly;
        }
        set isReadOnly(value: boolean) {
            if (value != this._readOnly) {
                this._readOnly = asBoolean(value);
                this.finishEditing();
                this.invalidate(true); // TFS 79965
                this._addHdl.updateNewRowTemplate(); // TFS 97544
                toggleClass(this.hostElement, 'wj-state-readonly', this.isReadOnly);
            }
        }
        /**
         * Gets or sets a value that determines whether the grid should support
         * Input Method Editors (IME) while not in edit mode.
         * 
         * This property is relevant only for sites/applications in Japanese,
         * Chinese, Korean, and other languages that require IME support.
         */
        get imeEnabled(): boolean {
            return this._imeHdl != null;
        }
        set imeEnabled(value: boolean) {
            if (value != this.imeEnabled) {
                if (this._imeHdl) {
                    this._imeHdl.dispose();
                    this._imeHdl = null;
                }
                if (value) {
                    this._imeHdl = new _ImeHandler(this);
                }
            }
        }
        /**
         * Gets or sets a value that determines whether users may resize 
         * rows and/or columns with the mouse.
         *
         * If resizing is enabled, users can resize columns by dragging
         * the right edge of column header cells, or rows by dragging the
         * bottom edge of row header cells.
         *
         * Users may also double-click the edge of the header cells to 
         * automatically resize rows and columns to fit their content.
         * The auto-size behavior can be customized using the @see:autoSizeMode
         * property.
         */
        get allowResizing(): AllowResizing {
            return this._alResizing;
        }
        set allowResizing(value: AllowResizing) {
            this._alResizing = asEnum(value, AllowResizing);
        }
        /**
         * Gets or sets a value that determines whether row and column resizing 
         * should be deferred until the user releases the mouse button.
         *
         * By default, @see:deferResizing is set to false, causing rows and columns
         * to be resized as the user drags the mouse. Setting this property to true
         * causes the grid to show a resizing marker and to resize the row or column
         * only when the user releases the mouse button.
         */
        get deferResizing(): boolean {
            return this._deferResizing;
        }
        set deferResizing(value: boolean) {
            this._deferResizing = asBoolean(value);
        }
        /**
         * Gets or sets which cells should be taken into account when auto-sizing a
         * row or column.
         *
         * This property controls what happens when users double-click the edge of
         * a column header.
         *
         * By default, the grid will automatically set the column width based on the
         * content of the header and data cells in the column. This property allows 
         * you to change that to include only the headers or only the data.
         */
        get autoSizeMode(): AutoSizeMode {
            return this._autoSizeMode;
        }
        set autoSizeMode(value: AutoSizeMode) {
            this._autoSizeMode = asEnum(value, AutoSizeMode);
        }
        /**
         * Gets or sets a value that determines whether users are allowed to sort columns 
         * by clicking the column header cells.
         */
        get allowSorting(): boolean {
            return this._alSorting;
        }
        set allowSorting(value: boolean) {
            this._alSorting = asBoolean(value);
        }
        /**
         * Gets or sets a value that indicates whether the grid should provide a new row
         * template so users can add items to the source collection.
         *
         * The new row template will not be displayed if the @see:isReadOnly property
         * is set to true.
         */
        get allowAddNew(): boolean {
            return this._alAddNew;
        }
        set allowAddNew(value: boolean) {
            if (value != this._alAddNew) {
                this._alAddNew = asBoolean(value);
                this._addHdl.updateNewRowTemplate();
            }
        }
        /**
         * Gets or sets a value that indicates whether the new row template should be located
         * at the top of the grid or at the bottom.
         *
         * The new row template will be displayed only if the @see:allowAddNew property is set
         * to true and if the @see:itemsSource object supports adding new items.
         */
        get newRowAtTop(): boolean {
            return this._addHdl.newRowAtTop;
        }
        set newRowAtTop(value: boolean) {
            this._addHdl.newRowAtTop = asBoolean(value);
        }
        /**
         * Gets or sets a value that indicates whether the grid should delete 
         * selected rows when the user presses the Delete key.
         *
         * Selected rows will not be deleted if the @see:isReadOnly property
         * is set to true.
         */
        get allowDelete(): boolean {
            return this._alDelete;
        }
        set allowDelete(value: boolean) {
            if (value != this._alDelete) {
                this._alDelete = asBoolean(value);
            }
        }
        /**
         * Gets or sets which parts of the grid provide cell merging.
         */
        get allowMerging(): AllowMerging {
            return this._alMerging;
        }
        set allowMerging(value: AllowMerging) {
            if (value != this._alMerging) {
                this._alMerging = asEnum(value, AllowMerging);
                this.invalidate();
            }
        }
        /**
         * Gets or sets a value that indicates whether the grid should
         * add class names to indicate selected header cells.
         */
        get showSelectedHeaders(): HeadersVisibility {
            return this._ssHdr;
        }
        set showSelectedHeaders(value: HeadersVisibility) {
            if (value != this._ssHdr) {
                this._ssHdr = asEnum(value, HeadersVisibility);
                this.invalidate();
            }
        }
        /**
         * Gets or sets a value that indicates whether the grid should
         * display a marquee element around the current selection.
         */
        get showMarquee(): boolean {
            return !this._eMarquee.style.display;
        }
        set showMarquee(value: boolean) {
            if (value != this.showMarquee) {
                var s = this._eMarquee.style;
                s.visibility = 'collapse'; // show only after positioning
                s.display = asBoolean(value) ? '' : 'none';
                this.invalidate();
            }
        }
        /**
         * Gets or sets a value that determines whether the grid should display 
         * sort indicators in the column headers.
         *
         * Sorting is controlled by the @see:ICollectionView.sortDescriptions
         * property of the @see:ICollectionView object used as a the grid's
         * @see:itemsSource.
         */
        get showSort(): boolean {
            return this._shSort;
        }
        set showSort(value: boolean) {
            if (value != this._shSort) {
                this._shSort = asBoolean(value);
                this.invalidate();
            }
        }
        /**
         * Gets or sets a value that determines whether the grid should insert group 
         * rows to delimit data groups.
         *
         * Data groups are created by modifying the @see:ICollectionView.groupDescriptions
         * property of the @see:ICollectionView object used as a the grid's @see:itemsSource.
         */
        get showGroups(): boolean {
            return this._shGroups;
        }
        set showGroups(value: boolean) {
            if (value != this._shGroups) {
                this._shGroups = asBoolean(value);
                this._bindGrid(false);
            }
        }
        /**
         * Gets or sets a value that determines whether the grid should add the 'wj-alt' 
         * class to cells in alternating rows.
         * 
         * Setting this property to false disables alternate row styles without any
         * changes to the CSS.
         */
        get showAlternatingRows(): boolean {
            return this._shAlt;
        }
        set showAlternatingRows(value: boolean) {
            if (value != this._shAlt) {
                this._shAlt = asBoolean(value);
                this.invalidate();
            }
        }
        /**
         * Gets or sets the format string used to create the group header content.
         *
         * The string may contain any text, plus the following replacement strings:
         * <ul>
         *   <li><b>{name}</b>: The name of the property being grouped on.</li>
         *   <li><b>{value}</b>: The value of the property being grouped on.</li>
         *   <li><b>{level}</b>: The group level.</li>
         *   <li><b>{count}</b>: The total number of items in this group.</li>
         * </ul>
         *
         * If a column is bound to the grouping property, the column header is used 
         * to replace the <code>{name}</code> parameter, and the column's format and
         * data maps are used to calculate the <code>{value}</code> parameter.
         * If no column is available, the group information is used instead.
         *
         * You may add invisible columns bound to the group properties in order to 
         * customize the formatting of the group header cells.
         *
         * The default value for this property is<br/>
         * <code>'{name}: &lt;b&gt;{value}&lt;/b&gt;({count:n0} items)'</code>,
         * which creates group headers similar to<br/>
         * <code>'Country: <b>UK</b> (12 items)'</code> or<br/>
         * <code>'Country: <b>Japan</b> (8 items)'</code>.
         */
        get groupHeaderFormat(): string {
            return this._gHdrFmt;
        }
        set groupHeaderFormat(value: string) {
            if (value != this._gHdrFmt) {
                this._gHdrFmt = asString(value)
                this._bindGrid(false);
            }
        }
        /**
         * Gets or sets a value that determines whether users are allowed to drag 
         * rows and/or columns with the mouse.
         */
        get allowDragging(): AllowDragging {
            return this._alDragging;
        }
        set allowDragging(value: AllowDragging) {
            if (value != this._alDragging) {
                this._alDragging = asEnum(value, AllowDragging);
                this.invalidate(); // to re-create row/col headers
            }
        }
        /**
         * Gets or sets the array or @see:ICollectionView that contains items shown on the grid.
         */
        get itemsSource(): any {
            return this._items;
        }
        set itemsSource(value: any) {
            if (value != this._items) {

                // unbind current collection view
                if (this._cv) {
                    var cv = tryCast(this._cv, collections.CollectionView);
                    if (cv && cv.sortConverter == this._bndSortConverter) {
                        cv.sortConverter = null;
                    }
                    this._cv.currentChanged.removeHandler(this._cvCurrentChanged, this);
                    this._cv.collectionChanged.removeHandler(this._cvCollectionChanged, this);
                    this._cv = null;
                }

                // save new data source and collection view
                this._items = value;
                this._cv = this._getCollectionView(value);
                this._lastCount = 0;

                // bind new collection view
                if (this._cv) {
                    this._cv.currentChanged.addHandler(this._cvCurrentChanged, this);
                    this._cv.collectionChanged.addHandler(this._cvCollectionChanged, this);
                    var cv = tryCast(this._cv, collections.CollectionView);
                    if (cv && !cv.sortConverter) {
                        cv.sortConverter = this._bndSortConverter;
                    }
                }

                // bind grid
                this._bindGrid(true);

                // raise itemsSourceChanged
                this.onItemsSourceChanged();
            }
        }
        /**
         * Gets the @see:ICollectionView that contains the grid data.
         */
        get collectionView(): collections.ICollectionView {
            return this._cv;
        }
        /**
         * Gets or sets the name of the property (or properties) used to generate 
         * child rows in hierarchical grids.
         *
         * Set this property to a string to specify the name of the property that
         * contains an item's child items (e.g. <code>'items'</code>). 
         * 
         * If items at different levels child items with different names, then 
         * set this property to an array containing the names of the properties 
         * that contain child items et each level 
         * (e.g. <code>[ 'accounts', 'checks', 'earnings' ]</code>).
         * 
         * @fiddle:t0ncmjwp
         */
        get childItemsPath(): any {
            return this._childItemsPath;
        }
        set childItemsPath(value: any) {
            if (value != this._childItemsPath) {
                assert(value == null || isArray(value) || isString(value), 'childItemsPath should be an array or a string.');
                this._childItemsPath = value;
                this._bindGrid(true);
            }
        }
        /**
         * Gets the @see:GridPanel that contains the data cells.
         */
        get cells(): GridPanel {
            return this._gpCells;
        }
        /**
         * Gets the @see:GridPanel that contains the column header cells.
         */
        get columnHeaders(): GridPanel {
            return this._gpCHdr;
        }
        /**
         * Gets the @see:GridPanel that contains the column footer cells.
         *
         * The @see:columnFooters panel appears below the grid cells, to the
         * right of the @see:bottomLeftCells panel. It can be used to display
         * summary information below the grid data.
         *
         * The example below shows how you can add a row to the @see:columnFooters
         * panel to display summary data for columns that have the
         * @see:Column.aggregate property set:
         *
         * <pre>function addFooterRow(flex) {
         *   // create a GroupRow to show aggregates
         *   var row = new wijmo.grid.GroupRow();
         *
         *   // add the row to the column footer panel
         *   flex.columnFooters.rows.push(row);
         *
         *   // show a sigma on the header
         *   flex.bottomLeftCells.setCellData(0, 0, '\u03A3');
         * }</pre>
         */
        get columnFooters(): GridPanel {
            return this._gpCFtr;
        }
        /**
         * Gets the @see:GridPanel that contains the row header cells.
         */
        get rowHeaders(): GridPanel {
            return this._gpRHdr;
        }
        /**
         * Gets the @see:GridPanel that contains the top left cells
         * (to the left of the column headers).
         */
        get topLeftCells(): GridPanel {
            return this._gpTL;
        }
        /**
         * Gets the @see:GridPanel that contains the bottom left cells.
         *
         * The @see:bottomLeftCells panel appears below the row headers, to the
         * left of the @see:columnFooters panel.
         */
        get bottomLeftCells(): GridPanel {
            return this._gpBL;
        }
        /**
         * Gets the grid's row collection.
         */
        get rows(): RowCollection {
            return this._rows;
        }
        /**
         * Gets the grid's column collection.
         */
        get columns(): ColumnCollection {
            return this._cols;
        }
        /**
         * Gets a column by name or by binding.
         *
         * The method searches the column by name. If a column with the given name 
         * is not found, it searches by binding. The searches are case-sensitive.
         *
         * @param name The name or binding to find.
         * @return The column with the specified name or binding, or null if not found.
         */
        getColumn(name: string): Column {
            return this.columns.getColumn(name);
        }
        /**
         * Gets or sets the number of frozen rows.
         *
         * Frozen rows do not scroll, but the cells they contain 
         * may be selected and edited.
         */
        get frozenRows(): number {
            return this.rows.frozen;
        }
        set frozenRows(value: number) {
            this.rows.frozen = value;
        }
        /**
         * Gets or sets the number of frozen columns.
         *
         * Frozen columns do not scroll, but the cells they contain 
         * may be selected and edited.
         */
        get frozenColumns(): number {
            return this.columns.frozen;
        }
        set frozenColumns(value: number) {
            this.columns.frozen = value;
        }
        /**
         * Gets or sets the index of row in the column header panel that
         * shows and changes the current sort.
         *
         * This property is set to null by default, causing the last row
         * in the @see:columnHeaders panel to act as the sort row.
         */
        get sortRowIndex(): number {
            return this._sortRowIndex;
        }
        set sortRowIndex(value: number) {
            if (value != this._sortRowIndex) {
                this._sortRowIndex = asNumber(value, true);
                this.invalidate();
            }
        }
        /**
         * Gets or sets a @see:Point that represents the value of the grid's scrollbars.
         */
        get scrollPosition(): Point {
            return this._ptScrl.clone();
        }
        set scrollPosition(pt: Point) {
            var root = this._root,
                left = -pt.x;

            // IE/Chrome/FF handle scrollLeft differently under RTL:
            // Chrome reverses direction, FF uses negative values, IE does the right thing (nothing)
            if (this._rtl) {
                switch (FlexGrid._getRtlMode()) {
                    case 'rev':
                        left = (root.scrollWidth - root.clientWidth) + pt.x;
                        break;
                    case 'neg':
                        left = pt.x;
                        break;
                    default:
                        left = -pt.x;
                        break;
                }
            }
            root.scrollLeft = left;
            root.scrollTop = -pt.y;
        }
        /**
         * Gets the client size of the control (control size minus headers and scrollbars).
         */
        get clientSize(): Size {
            return this._szClient;
        }
        /**
         * Gets the bounding rectangle of the control in page coordinates.
         */
        get controlRect(): Rect {
            if (!this._rcBounds) {
                this._rcBounds = getElementRect(this._root);
            }
            return this._rcBounds;
        }
        /**
         * Gets the size of the grid content in pixels.
         */
        get scrollSize(): Size {
            return new Size(this._gpCells.width, this._heightBrowser);
        }
        /**
         * Gets the range of cells currently in view.
         */
        get viewRange(): CellRange {
            return this._gpCells.viewRange;
        }
        /**
         * Gets or sets the @see:CellFactory that creates and updates cells for this grid.
         */
        get cellFactory(): CellFactory {
            return this._cf;
        }
        set cellFactory(value: CellFactory) {
            if (value != this._cf) {
                this._cf = asType(value, CellFactory, false);
                this.invalidate();
            }
        }

        /**
         * Gets or sets a formatter function used to customize cells on this grid.
         *
         * The formatter function can add any content to any cell. It provides 
         * complete flexibility over the appearance and behavior of grid cells.
         *
         * If specified, the function should take four parameters: the @see:GridPanel
         * that contains the cell, the row and column indices of the cell, and the
         * HTML element that represents the cell. The function will typically change
         * the <b>innerHTML</b> property of the cell element.
         *
         * For example:
         * <pre>
         * flex.itemFormatter = function(panel, r, c, cell) {
         *   if (panel.cellType == CellType.Cell) {
         *     // draw sparklines in the cell
         *     var col = panel.columns[c];
         *     if (col.name == 'sparklines') {
         *       cell.innerHTML = getSparklike(panel, r, c);
         *     }
         *   }
         * }
         * </pre>
         *
         * Note that the FlexGrid recycles cells, so if your @see:itemFormatter 
         * modifies the cell's style attributes, you must make sure that it resets 
         * these attributes for cells that should not have them. For example:
         *
         * <pre>
         * flex.itemFormatter = function(panel, r, c, cell) {
         *   // reset attributes we are about to customize
         *   var s = cell.style;
         *   s.color = '';
         *   s.backgroundColor = '';
         *   // customize color and backgroundColor attributes for this cell
         *   ...
         * }
         * </pre>
         *
         * If you have a scenario where multiple clients may want to customize the
         * grid rendering (for example when creating directives or re-usable libraries),
         * consider using the @see:formatItem event instead. The event allows multiple
         * clients to attach their own handlers.
         */
        get itemFormatter(): Function {
            return this._itemFormatter;
        }
        set itemFormatter(value: Function) {
            if (value != this._itemFormatter) {
                this._itemFormatter = asFunction(value);
                this.invalidate();
            }
        }
        /**
         * Gets the value stored in a cell in the scrollable area of the grid.
         *
         * @param r Index of the row that contains the cell.
         * @param c Index of the column that contains the cell.
         * @param formatted Whether to format the value for display.
         */
        getCellData(r: number, c: number, formatted: boolean): any {
            return this.cells.getCellData(r, c, formatted);
        }
        /**
         * Gets a the bounds of a cell element in viewport coordinates.
         *
         * This method returns the bounds of cells in the @see:cells 
         * panel (scrollable data cells). To get the bounds of cells
         * in other panels, use the @see:getCellBoundingRect method
         * in the appropriate @see:GridPanel object.
         *
         * The returned value is a @see:Rect object which contains the
         * position and dimensions of the cell in viewport coordinates.
         * The viewport coordinates are the same used by the 
         * <a href="https://developer.mozilla.org/en-US/docs/Web/API/Element.getBoundingClientRect">getBoundingClientRect</a>
         * method.
         *
         * @param r Index of the row that contains the cell.
         * @param c Index of the column that contains the cell.
         * @param raw Whether to return the rectangle in raw panel coordinates as opposed to viewport coordinates.
         */
        getCellBoundingRect(r: number, c: number, raw?: boolean): Rect {
            return this.cells.getCellBoundingRect(r, c, raw);
        }
        /**
         * Sets the value of a cell in the scrollable area of the grid.
         *
         * @param r Index of the row that contains the cell.
         * @param c Index, name, or binding of the column that contains the cell.
         * @param value Value to store in the cell.
         * @param coerce Whether to change the value automatically to match the column's data type.
         * @param invalidate Whether to invalidate the grid to show the change.
         * @return True if the value was stored successfully, false otherwise.
         */
        setCellData(r: number, c: any, value: any, coerce = true, invalidate = true): boolean {
            return this.cells.setCellData(r, c, value, coerce, invalidate);
        }
        /**
         * Gets a @see:HitTestInfo object with information about a given point.
         *
         * For example:
         *
         * <pre>// hit test a point when the user clicks on the grid
         * flex.hostElement.addEventListener('click', function (e) {
         *   var ht = flex.hitTest(e.pageX, e.pageY);
         *   console.log('you clicked a cell of type "' +
         *     wijmo.grid.CellType[ht.cellType] + '".');
         * });</pre>
         *
         * @param pt @see:Point to investigate, in page coordinates, or a MouseEvent object, or x coordinate of the point.
         * @param y Y coordinate of the point in page coordinates (if the first parameter is a number).
         * @return A @see:HitTestInfo object with information about the point.
         */
        hitTest(pt: any, y?: number): HitTestInfo {
            if (isNumber(pt) && isNumber(y)) { // accept hitTest(x, y) as well
                pt = new Point(pt, y);
            }
            return new HitTestInfo(this, pt);
        }
        /**
         * Gets the content of a @see:CellRange as a string suitable for 
         * copying to the clipboard.
         *
         * Hidden rows and columns are not included in the clip string.
         *
         * @param rng @see:CellRange to copy. If omitted, the current selection is used.
         */
        getClipString(rng?: CellRange): string {
            var clipString = '',
                firstRow = true,
                firstCell = true;

            // scan rows
            rng = rng ? asType(rng, CellRange) : this.selection;
            for (var r = rng.topRow; r <= rng.bottomRow; r++) {

                // skip invisible, add separator
                if (!this.rows[r].isVisible) continue;
                if (!firstRow) clipString += '\n';
                firstRow = false;

                // scan cells
                for (var c = rng.leftCol, firstCell = true; c <= rng.rightCol; c++) {

                    // skip invisible, add separator
                    if (!this.columns[c].isVisible) continue;
                    if (!firstCell) clipString += '\t';
                    firstCell = false;

                    // append cell
                    var cell = this.cells.getCellData(r, c, true).toString();
                    cell = cell.replace(/\t/g, ' ');
                    clipString += cell;
                }
            }

            // done
            return clipString;
        }
        /**
         * Parses a string into rows and columns and applies the content to a given range.
         *
         * Hidden rows and columns are skipped.
         *
         * @param text Tab and newline delimited text to parse into the grid.
         * @param rng @see:CellRange to copy. If omitted, the current selection is used.
         */
        setClipString(text: string, rng?: CellRange) {

            // get target range
            var autoRange = rng == null;
            rng = rng ? asType(rng, CellRange) : this.selection;

            // normalize text
            text = asString(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
            if (text && text[text.length - 1] == '\n') {
                text = text.substring(0, text.length - 1);
            }
            if (autoRange && !rng.isSingleCell) {
                text = this._expandClipString(text, rng);
            }

            // keep track of paste range to select later
            var rngPaste = new CellRange(rng.topRow, rng.leftCol);

            // copy lines to rows
            this.beginUpdate();
            var row = rng.topRow,
                lines = text.split('\n'),
                pasted = false,
                e: CellRangeEventArgs;
            for (var i = 0; i < lines.length && row < this.rows.length; i++ , row++) {

                // skip invisible row, keep clip line
                if (!this.rows[row].isVisible) {
                    i--;
                    continue;
                }

                // copy cells to columns
                var cells = lines[i].split('\t'),
                    col = rng.leftCol;
                for (var j = 0; j < cells.length && col < this.columns.length; j++ , col++) {

                    // skip invisible column, keep clip cell
                    if (!this.columns[col].isVisible) {
                        j--;
                        continue;
                    }

                    // assign cell
                    if (!this.columns[col].isReadOnly && !this.rows[row].isReadOnly) {

                        // raise events so user can cancel the paste
                        e = new CellRangeEventArgs(this.cells, new CellRange(row, col), cells[j]);
                        if (this.onPastingCell(e)) {
                            if (this.cells.setCellData(row, col, e.data)) {
                                this.onPastedCell(e);
                                pasted = true;
                            }
                        }

                        // update paste range
                        rngPaste.row2 = Math.max(rngPaste.row2, row);
                        rngPaste.col2 = Math.max(rngPaste.col2, col);
                    }
                }
            }
            this.endUpdate();

            // handle editing
            if (pasted) {
                var ecv = tryCast(this.collectionView, 'IEditableCollectionView');
                if (ecv && ecv.currentItem == ecv.currentAddItem) {
                    ecv.editItem(ecv.currentItem); // editing a new item, get into edit mode
                } else if (this.collectionView) {
                    this.collectionView.refresh(); // update sorting/filtering 
                }
            }

            // select pasted range
            this.select(rngPaste);
        }

        // expand clip string to get Excel-like behavior
        _expandClipString(text: string, rng: CellRange): string {

            // sanity
            if (!text) return text;

            // get clip string dimensions and cells
            var lines = text.split('\n'),
                srcRows = lines.length,
                srcCols = 0,
                rows = [];
            for (var r = 0; r < srcRows; r++) {
                var cells = lines[r].split('\t');
                rows.push(cells);
                if (r > 1 && cells.length != srcCols) return text;
                srcCols = cells.length;
            }

            // expand if destination size is a multiple of source size (like Excel)
            var dstRows = rng.rowSpan,
                dstCols = rng.columnSpan;
            if (dstRows > 1 || dstCols > 1) {
                if (dstRows == 1) dstRows = srcRows;
                if (dstCols == 1) dstCols = srcCols;
                if (dstCols % srcCols == 0 && dstRows % srcRows == 0) {
                    text = '';
                    for (var r = 0; r < dstRows; r++) {
                        for (var c = 0; c < dstCols; c++) {
                            if (r > 0 && c == 0) text += '\n';
                            if (c > 0) text += '\t';
                            text += rows[r % srcRows][c % srcCols];
                        }
                    }
                }
            }

            // done
            return text;
        }
        /** 
         * Overridden to set the focus to the grid without scrolling the whole grid into view.
         */
        focus() {

            // if we have an active editor, let it keep the focus
            if (this.activeEditor) {
                this.activeEditor.focus();
                return;
            }

            // save work if possible
            if (wijmo.getActiveElement() == this._eFocus) {
                return;
            }

            // if the grid is visible, use inner element '_efocus'
            var rc = this.hostElement.getBoundingClientRect();
            if (rc.bottom > 0 && rc.right > 0 && rc.top < innerHeight && rc.left < innerWidth) {
                setCss(this._eFocus, {
                    top: Math.max(0, -rc.top),
                    left: Math.max(0, -rc.left)
                });
                this._eFocus.focus();
                return;
            }

            // if the grid is off the screen, try focusing on the selected cell
            var cell = <HTMLElement>this.cells.hostElement.querySelector('.wj-cell.wj-state-selected');
            if (cell) {
                cell.focus();
                return;
            }

            // off-screen, no current cell, focus on the grid's host element (will scroll it into view)
            super.focus();
        }
        /**
         * Checks whether this control contains the focused element.
         */
        containsFocus(): boolean {
            var lbx = this._edtHdl ? this._edtHdl._lbx : null;
            return super.containsFocus() || (lbx && lbx.containsFocus());
        }
        /**
         * Disposes of the control by removing its association with the host element.
         */
        dispose() {

            // cancel any pending edits, close drop-down list
            this.finishEditing(true);

            // remove itemsSource so it doesn't have references to our
            // change event handlers that would prevent the grid from being
            // garbage-collected.
            this.itemsSource = null;

            // remove any pending animation frame requests
            if (this._afScrl) {
                cancelAnimationFrame(this._afScrl);
            }

            // allow base class
            super.dispose();
        }
        /**
         * Refreshes the grid display.
         *
         * @param fullUpdate Whether to update the grid layout and content, or just the content.
         */
        refresh(fullUpdate = true) {

            // always call base class to handle being/endUpdate logic
            super.refresh(fullUpdate);

            // close any open drop-downs
            this.finishEditing();

            // on full updates, get missing column types based on bindings and
            // update scroll position in case the control just became visible
            // and IE wrongly reset the element's scroll position to the origin
            // http://wijmo.com/topic/flexgrid-refresh-issue-when-hidden/
            if (fullUpdate) {
                this._updateColumnTypes();
                this.scrollPosition = this._ptScrl; // update element to match grid
                //this._updateScrollPosition(); // update grid to match element
            }

            // go refresh the cells
            this.refreshCells(fullUpdate);
        }
        /**
         * Refreshes the grid display.
         *
         * @param fullUpdate Whether to update the grid layout and content, or just the content.
         * @param recycle Whether to recycle existing elements.
         * @param state Whether to keep existing elements and update their state.
         */
        refreshCells(fullUpdate: boolean, recycle?: boolean, state?: boolean) {
            if (!this.isUpdating) {
                if (fullUpdate) {
                    this._updateLayout();
                } else {
                    this._updateContent(recycle, state);
                }
            }
        }
        /**
         * Resizes a column to fit its content.
         *
         * @param c Index of the column to resize.
         * @param header Whether the column index refers to a regular or a header row.
         * @param extra Extra spacing, in pixels.
         */
        autoSizeColumn(c: number, header = false, extra = 4) {
            this.autoSizeColumns(c, c, header, extra);
        }
        /**
         * Resizes a range of columns to fit their content.
         *
         * The grid will always measure all rows in the current view range, plus up to 2,000 rows
         * not currently in view. If the grid contains a large amount of data (say 50,000 rows), 
         * then not all rows will be measured since that could potentially take a long time.
         *
         * @param firstColumn Index of the first column to resize (defaults to the first column).
         * @param lastColumn Index of the last column to resize (defaults to the last column).
         * @param header Whether the column indices refer to regular or header columns.
         * @param extra Extra spacing, in pixels.
         */
        autoSizeColumns(firstColumn?: number, lastColumn?: number, header = false, extra = 4) {
            var max = 0,
                pHdr = header ? this.topLeftCells : this.columnHeaders,
                pCells = header ? this.rowHeaders : this.cells,
                rowRange = this.viewRange,
                text: string, lastText: string;
            firstColumn = firstColumn == null ? 0 : asInt(firstColumn);
            lastColumn = lastColumn == null ? pCells.columns.length - 1 : asInt(lastColumn);
            asBoolean(header);
            asNumber(extra);

            // choose row range to measure
            // (viewrange by default, everything if we have only a few items)
            rowRange.row = Math.max(0, rowRange.row - 1000);
            rowRange.row2 = Math.min(rowRange.row2 + 1000, this.rows.length - 1);

            this.deferUpdate(() => {

                // create element to measure content
                var eMeasure = document.createElement('div');
                eMeasure.setAttribute(FlexGrid._WJS_MEASURE, 'true');
                eMeasure.style.visibility = 'hidden';
                this.hostElement.appendChild(eMeasure);

                // measure cells in the range
                for (var c = firstColumn; c <= lastColumn && c > -1 && c < pCells.columns.length; c++) {
                    max = 0;

                    // headers
                    if (this.autoSizeMode & AutoSizeMode.Headers) {
                        for (var r = 0; r < pHdr.rows.length; r++) {
                            if (pHdr.rows[r].isVisible) {
                                var w = this._getDesiredWidth(pHdr, r, c, eMeasure);
                                max = Math.max(max, w);
                            }
                        }
                    }

                    // cells
                    if (this.autoSizeMode & AutoSizeMode.Cells) {
                        lastText = null;
                        for (var r = rowRange.row; r <= rowRange.row2 && r > -1 && r < pCells.rows.length; r++) {
                            if (pCells.rows[r].isVisible) {

                                if (!header && c == pCells.columns.firstVisibleIndex && pCells.rows.maxGroupLevel > -1) {

                                    // ignore last text for outline cells
                                    var w = this._getDesiredWidth(pCells, r, c, eMeasure);
                                    max = Math.max(max, w);

                                } else {

                                    // regular cells
                                    text = pCells.getCellData(r, c, true);
                                    if (text != lastText) {
                                        lastText = text;
                                        var w = this._getDesiredWidth(pCells, r, c, eMeasure);
                                        max = Math.max(max, w);
                                    }
                                }
                            }
                        }
                    }

                    // set size
                    pCells.columns[c].width = max + extra + 2;
                }

                // done with measuring element
                this.hostElement.removeChild(eMeasure);
            });
        }
        /**
         * Resizes a row to fit its content.
         *
         * @param r Index of the row to resize.
         * @param header Whether the row index refers to a regular or a header row.
         * @param extra Extra spacing, in pixels.
         */
        autoSizeRow(r: number, header = false, extra = 0) {
            this.autoSizeRows(r, r, header, extra);
        }
        /**
         * Resizes a range of rows to fit their content.
         *
         * @param firstRow Index of the first row to resize.
         * @param lastRow Index of the last row to resize.
         * @param header Whether the row indices refer to regular or header rows.
         * @param extra Extra spacing, in pixels.
         */
        autoSizeRows(firstRow?: number, lastRow?: number, header = false, extra = 0) {
            var max = 0,
                pHdr = header ? this.topLeftCells : this.rowHeaders,
                pCells = header ? this.columnHeaders : this.cells;
            header = asBoolean(header);
            extra = asNumber(extra);
            firstRow = firstRow == null ? 0 : asInt(firstRow);
            lastRow = lastRow == null ? pCells.rows.length - 1 : asInt(lastRow);

            this.deferUpdate(() => {

                // create element to measure content
                var eMeasure = document.createElement('div');
                eMeasure.setAttribute(FlexGrid._WJS_MEASURE, 'true');
                eMeasure.style.visibility = 'hidden';
                this.hostElement.appendChild(eMeasure);

                // measure cells in the range
                for (var r = firstRow; r <= lastRow && r > -1 && r < pCells.rows.length; r++) {
                    max = 0;

                    // headers
                    if (this.autoSizeMode & AutoSizeMode.Headers) {
                        for (var c = 0; c < pHdr.columns.length; c++) {
                            if (pHdr.columns[c].renderSize > 0) {
                                var h = this._getDesiredHeight(pHdr, r, c, eMeasure);
                                max = Math.max(max, h);
                            }
                        }
                    }

                    // cells
                    if (this.autoSizeMode & AutoSizeMode.Cells) {
                        for (var c = 0; c < pCells.columns.length; c++) {
                            if (pCells.columns[c].renderSize > 0) {
                                var h = this._getDesiredHeight(pCells, r, c, eMeasure);
                                max = Math.max(max, h);
                            }
                        }
                    }

                    // update size
                    pCells.rows[r].height = max + extra;
                }

                // done with measuring element
                this.hostElement.removeChild(eMeasure);
            });
        }
        /**
         * Gets or sets the indent used to offset row groups of different levels.
         */
        get treeIndent(): number {
            return this._indent;
        }
        set treeIndent(value: number) {
            if (value != this._indent) {
                this._indent = asNumber(value, false, true);
                this.columns.onCollectionChanged();
            }
        }
        /**
         * Collapses all the group rows to a given level.
         *
         * @param level Maximum group level to show.
         */
        collapseGroupsToLevel(level: number) {

            // finish editing first (this may change the collection)
            if (this.finishEditing()) {

                // set collapsed state for all rows in the grid
                var rows = this.rows;
                rows.deferUpdate(() => {
                    for (var r = 0; r < rows.length; r++) {
                        var gr = tryCast(rows[r], GroupRow);
                        if (gr) {
                            gr.isCollapsed = gr.level >= level;
                        }
                    }
                });
            }
        }
        /**
         * Gets or sets the current selection mode.
         */
        get selectionMode(): SelectionMode {
            return this._selHdl.selectionMode;
        }
        set selectionMode(value: SelectionMode) {
            if (value != this.selectionMode) {
                this._selHdl.selectionMode = asEnum(value, SelectionMode);
                this.invalidate();
            }
        }
        /**
         * Gets or sets the current selection.
         */
        get selection(): CellRange {
            return this._selHdl.selection.clone();
        }
        set selection(value: CellRange) {
            this._selHdl.selection = value;
        }
        /**
         * Selects a cell range and optionally scrolls it into view.
         *
         * @param rng Range to select.
         * @param show Whether to scroll the new selection into view.
         */
        select(rng: any, show: any = true) {
            this._selHdl.select(rng, show);
        }
        /**
         * Gets a @see:SelectedState value that indicates the selected state of a cell.
         *
         * @param r Row index of the cell to inspect.
         * @param c Column index of the cell to inspect.
         */
        getSelectedState(r: number, c: number): SelectedState {
            return this.cells.getSelectedState(r, c, null);
        }
        /**
         * Gets or sets an array containing the rows that are currently selected.
         *
         * Note: this property can be read in all selection modes, but it can be
         * set only when @see:selectionMode is set to <b>SelectionMode.ListBox</b>.
         */
        get selectedRows(): any[] {
            var rows = [];
            if (this.selectionMode == SelectionMode.ListBox) {
                for (var i = 0; i < this.rows.length; i++) {
                    if (this.rows[i].isSelected) {
                        rows.push(this.rows[i]);
                    }
                }
            } else if (this.rows.length) {
                var sel = this.selection;
                for (var i = sel.topRow; i > -1 && i <= sel.bottomRow; i++) {
                    rows.push(this.rows[i]);
                }
            }
            return rows;
        }
        set selectedRows(value: any[]) {
            assert(this.selectionMode == SelectionMode.ListBox, 'This property can be set only in ListBox mode.');
            value = asArray(value);
            this.deferUpdate(() => {
                for (var i = 0, first = true; i < this.rows.length; i++) {
                    var row = <Row>this.rows[i],
                        sel = value && value.indexOf(row) > -1;
                    if (sel && first) {
                        first = false;
                        this.select(i, this.selection.col);
                    }
                    row.isSelected = sel;
                }
            });
        }
        /**
         * Gets or sets an array containing the data items that are currently selected.
         *
         * Note: this property can be read in all selection modes, but it can be
         * set only when @see:selectionMode is set to <b>SelectionMode.ListBox</b>.
         */
        get selectedItems(): any[] {
            var items = this.selectedRows;
            for (var i = 0; i < items.length; i++) {
                items[i] = (<Row>items[i]).dataItem;
            }
            return items;
        }
        set selectedItems(value: any[]) {
            assert(this.selectionMode == SelectionMode.ListBox, 'This property can be set only in ListBox mode.');
            value = asArray(value);
            this.deferUpdate(() => {
                for (var i = 0, first = true; i < this.rows.length; i++) {
                    var row = <Row>this.rows[i],
                        sel = value && value.indexOf(row.dataItem) > -1;
                    if (sel && first) {
                        first = false;
                        this.select(i, this.selection.col);
                    }
                    row.isSelected = sel;
                }
            });
        }
        /**
         * Scrolls the grid to bring a specific cell into view.
         *
         * @param r Index of the row to scroll into view.
         * @param c Index of the column to scroll into view.
         * @return True if the grid scrolled.
         */
        scrollIntoView(r: number, c: number): boolean {

            // make sure our dimensions are set and up-to-date
            if (this._maxOffsetY == null) {
                this._updateLayout();
            }

            // and go to work
            var sp = this.scrollPosition,
                wid = this._szClient.width,
                hei = this._szClient.height - this._gpCFtr.rows.getTotalSize(),
                ptFrz = this.cells._getFrozenPos();

            // calculate row offset
            r = asInt(r);
            if (r > -1 && r < this._rows.length && r >= this._rows.frozen) {
                var row = <Row>this._rows[r],
                    pct = this.cells.height > hei ? Math.round(row.pos / (this.cells.height - hei) * 100) / 100 : 0,
                    offsetY = Math.round(this._maxOffsetY * pct),
                    rpos = row.pos - offsetY,
                    rbot = rpos + row.renderSize - 1;
                if (rbot > hei - sp.y) {
                    sp.y = Math.max(-rpos, hei - rbot);
                }
                if (rpos - ptFrz.y < -sp.y) {
                    sp.y = -(rpos - ptFrz.y);
                }
            }

            // calculate column offset
            c = asInt(c);
            if (c > -1 && c < this._cols.length && c >= this._cols.frozen) {
                var col = <Column>this._cols[c],
                    rgt = col.pos + col.renderSize - 1;
                if (rgt > -sp.x + wid) {
                    sp.x = Math.max(-col.pos, wid - rgt);
                }
                if (col.pos - ptFrz.x < -sp.x) {
                    sp.x = -(col.pos - ptFrz.x);
                }
            }

            // update scroll position
            if (!sp.equals(this._ptScrl)) {
                this.scrollPosition = sp;
                return true;
            }

            // no change
            return false;
        }
        /**
         * Checks whether a given CellRange is valid for this grid's row and column collections.
         *
         * @param rng Range to check.
         */
        isRangeValid(rng: CellRange): boolean {
            return rng.isValid && rng.bottomRow < this.rows.length && rng.rightCol < this.columns.length;
        }
        /**
         * Starts editing a given cell.
         *
         * Editing in the @see:FlexGrid is similar to editing in Excel:
         * Pressing F2 or double-clicking a cell puts the grid in <b>full-edit</b> mode. 
         * In this mode, the cell editor remains active until the user presses Enter, Tab, 
         * or Escape, or until he moves the selection with the mouse. In full-edit mode, 
         * pressing the cursor keys does not cause the grid to exit edit mode.
         *
         * Typing text directly into a cell puts the grid in <b>quick-edit mode</b>.
         * In this mode, the cell editor remains active until the user presses Enter, 
         * Tab, or Escape, or any arrow keys.
         *
         * Full-edit mode is normally used to make changes to existing values. 
         * Quick-edit mode is normally used for entering new data quickly.
         *
         * While editing, the user can toggle between full and quick modes by 
         * pressing the F2 key.
         *
         * @param fullEdit Whether to stay in edit mode when the user presses the cursor keys. Defaults to false.
         * @param r Index of the row to be edited. Defaults to the currently selected row.
         * @param c Index of the column to be edited. Defaults to the currently selected column.
         * @param focus Whether to give the editor the focus when editing starts. Defaults to true.
         * @return True if the edit operation started successfully.
         */
        startEditing(fullEdit = true, r?: number, c?: number, focus?: boolean): boolean {
            return this._edtHdl.startEditing(fullEdit, r, c, focus);
        }
        /**
         * Commits any pending edits and exits edit mode.
         *
         * @param cancel Whether pending edits should be canceled or committed.
         * @return True if the edit operation finished successfully.
         */
        finishEditing(cancel = false): boolean {
            return this._edtHdl.finishEditing(cancel);
        }
        /**
         * Gets the <b>HTMLInputElement</b> that represents the cell editor currently active.
         */
        get activeEditor(): HTMLInputElement {
            return this._edtHdl.activeEditor;
        }
        /**
         * Gets a @see:CellRange that identifies the cell currently being edited.
         */
        get editRange(): CellRange {
            return this._edtHdl.editRange;
        }
        /**
         * Gets or sets the @see:MergeManager object responsible for determining how cells
         * should be merged.
         */
        get mergeManager(): MergeManager {
            return this._mrgMgr;
        }
        set mergeManager(value: MergeManager) {
            if (value != this._mrgMgr) {
                this._mrgMgr = asType(value, MergeManager, true);
                this.invalidate();
            }
        }
        /**
         * Gets a @see:CellRange that specifies the merged extent of a cell
         * in a @see:GridPanel.
         *
         * @param p The @see:GridPanel that contains the range.
         * @param r Index of the row that contains the cell.
         * @param c Index of the column that contains the cell.
         * @param clip Whether to clip the merged range to the grid's current view range.
         * @return A @see:CellRange that specifies the merged range, or null if the cell is not merged.
         */
        getMergedRange(p: GridPanel, r: number, c: number, clip = true): CellRange {
            return this._mrgMgr ? this._mrgMgr.getMergedRange(p, r, c, clip) : null;
        }

        //#endregion

        //--------------------------------------------------------------------------
        //#region ** events

        /**
         * Occurs after the grid has been bound to a new items source.
         */
        itemsSourceChanged = new Event();
        /**
         * Raises the @see:itemsSourceChanged event.
         */
        onItemsSourceChanged(e?: EventArgs) {
            this.itemsSourceChanged.raise(this, e);
        }
        /**
         * Occurs after the control has scrolled.
         */
        scrollPositionChanged = new Event();
        /**
         * Raises the @see:scrollPositionChanged event.
         */
        onScrollPositionChanged(e?: EventArgs) {
            this.scrollPositionChanged.raise(this, e);
        }
        /**
         * Occurs before selection changes.
         */
        selectionChanging = new Event();
        /**
         * Raises the @see:selectionChanging event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onSelectionChanging(e: CellRangeEventArgs): boolean {
            this.selectionChanging.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs after selection changes.
         */
        selectionChanged = new Event();
        /**
         * Raises the @see:selectionChanged event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onSelectionChanged(e: CellRangeEventArgs) {
            this.selectionChanged.raise(this, e);
        }
        /**
         * Occurs before the grid rows are bound to items in the data source.
         */
        loadingRows = new Event();
        /**
         * Raises the @see:loadingRows event.
         *
         * @param e @see:CancelEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onLoadingRows(e: CancelEventArgs): boolean {
            this.loadingRows.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs after the grid rows have been bound to items in the data source.
         */
        loadedRows = new Event();
        /**
         * Raises the @see:loadedRows event.
         */
        onLoadedRows(e?: EventArgs) {
            this.loadedRows.raise(this, e);
        }
        /**
         * Occurs before the grid updates its internal layout.
         */
        updatingLayout = new Event();
        /**
         * Raises the @see:updatingLayout event.
         *
         * @param e @see:CancelEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onUpdatingLayout(e: CancelEventArgs): boolean {
            this.updatingLayout.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs after the grid has updated its internal layout.
         */
        updatedLayout = new Event();
        /**
         * Raises the @see:updatedLayout event.
         */
        onUpdatedLayout(e?: EventArgs) {
            this.updatedLayout.raise(this, e);
        }
        /**
         * Occurs as columns are resized.
         */
        resizingColumn = new Event();
        /**
         * Raises the @see:resizingColumn event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onResizingColumn(e: CellRangeEventArgs): boolean {
            this.resizingColumn.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs when the user finishes resizing a column.
         */
        resizedColumn = new Event();
        /**
         * Raises the @see:resizedColumn event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onResizedColumn(e: CellRangeEventArgs) {
            this.resizedColumn.raise(this, e);
        }
        /**
         * Occurs before the user auto-sizes a column by double-clicking the 
         * right edge of a column header cell.
         */
        autoSizingColumn = new Event();
        /**
         * Raises the @see:autoSizingColumn event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onAutoSizingColumn(e: CellRangeEventArgs): boolean {
            this.autoSizingColumn.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs after the user auto-sizes a column by double-clicking the 
         * right edge of a column header cell.
         */
        autoSizedColumn = new Event();
        /**
         * Raises the @see:autoSizedColumn event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onAutoSizedColumn(e: CellRangeEventArgs) {
            this.autoSizedColumn.raise(this, e);
        }
        /**
         * Occurs when the user starts dragging a column.
         */
        draggingColumn = new Event();
        /**
         * Raises the @see:draggingColumn event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onDraggingColumn(e: CellRangeEventArgs): boolean {
            this.draggingColumn.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs when the user finishes dragging a column.
         */
        draggedColumn = new Event();
        /**
         * Raises the @see:draggedColumn event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onDraggedColumn(e: CellRangeEventArgs) {
            this.draggedColumn.raise(this, e);
        }
        /**
         * Occurs as rows are resized.
         */
        resizingRow = new Event();
        /**
         * Raises the @see:resizingRow event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onResizingRow(e: CellRangeEventArgs): boolean {
            this.resizingRow.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs when the user finishes resizing rows.
         */
        resizedRow = new Event();
        /**
         * Raises the @see:resizedRow event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onResizedRow(e: CellRangeEventArgs) {
            this.resizedRow.raise(this, e);
        }
        /**
         * Occurs before the user auto-sizes a row by double-clicking the 
         * bottom edge of a row header cell.
         */
        autoSizingRow = new Event();
        /**
         * Raises the @see:autoSizingRow event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onAutoSizingRow(e: CellRangeEventArgs): boolean {
            this.autoSizingRow.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs after the user auto-sizes a row by double-clicking the 
         * bottom edge of a row header cell.
         */
        autoSizedRow = new Event();
        /**
         * Raises the @see:autoSizedRow event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onAutoSizedRow(e: CellRangeEventArgs) {
            this.autoSizedRow.raise(this, e);
        }
        /**
         * Occurs when the user starts dragging a row.
         */
        draggingRow = new Event();
        /**
         * Raises the @see:draggingRow event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onDraggingRow(e: CellRangeEventArgs): boolean {
            this.draggingRow.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs when the user finishes dragging a row.
         */
        draggedRow = new Event();
        /**
         * Raises the @see:draggedRow event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onDraggedRow(e: CellRangeEventArgs) {
            this.draggedRow.raise(this, e);
        }
        /**
         * Occurs when a group is about to be expanded or collapsed.
         */
        groupCollapsedChanging = new Event();
        /**
         * Raises the @see:groupCollapsedChanging event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onGroupCollapsedChanging(e: CellRangeEventArgs): boolean {
            this.groupCollapsedChanging.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs after a group has been expanded or collapsed.
         */
        groupCollapsedChanged = new Event();
        /**
         * Raises the @see:groupCollapsedChanged event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onGroupCollapsedChanged(e: CellRangeEventArgs) {
            this.groupCollapsedChanged.raise(this, e);
        }
        /**
         * Occurs before the user applies a sort by clicking on a column header.
         */
        sortingColumn = new Event();
        /**
         * Raises the @see:sortingColumn event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onSortingColumn(e: CellRangeEventArgs): boolean {
            this.sortingColumn.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs after the user applies a sort by clicking on a column header.
         */
        sortedColumn = new Event();
        /**
         * Raises the @see:sortedColumn event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onSortedColumn(e: CellRangeEventArgs) {
            this.sortedColumn.raise(this, e);
        }
        /**
         * Occurs before a cell enters edit mode.
         */
        beginningEdit = new Event();
        /**
         * Raises the @see:beginningEdit event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onBeginningEdit(e: CellRangeEventArgs): boolean {
            this.beginningEdit.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs when an editor cell is created and before it becomes active.
         */
        prepareCellForEdit = new Event();
        /**
         * Raises the @see:prepareCellForEdit event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onPrepareCellForEdit(e: CellRangeEventArgs) {
            this.prepareCellForEdit.raise(this, e);
        }
        /**
         * Occurs when a cell edit is ending.
         */
        cellEditEnding = new Event();
        /**
         * Raises the @see:cellEditEnding event.
         *
         * You can use this event to perform validation and prevent invalid edits.
         * For example, the code below prevents users from entering values that
         * do not contain the letter 'a'. The code demonstrates how you can obtain
         * the old and new values before the edits are applied.
         *
         * <pre>function cellEditEnding (sender, e) {
         *   // get old and new values
         *   var flex = sender,
         *       oldVal = flex.getCellData(e.row, e.col),
         *       newVal = flex.activeEditor.value;
         *   // cancel edits if newVal doesn't contain 'a'
         *   e.cancel = newVal.indexOf('a') &lt; 0;
         * }</pre>
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onCellEditEnding(e: CellRangeEventArgs): boolean {
            this.cellEditEnding.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs when a cell edit has been committed or canceled.
         */
        cellEditEnded = new Event();
        /**
         * Raises the @see:cellEditEnded event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onCellEditEnded(e: CellRangeEventArgs) {
            this.cellEditEnded.raise(this, e);
        }
        /**
         * Occurs when a row edit is ending, before the changes are committed or canceled.
         */
        rowEditEnding = new Event();
        /**
         * Raises the @see:rowEditEnding event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onRowEditEnding(e: CellRangeEventArgs) {
            this.rowEditEnding.raise(this, e);
        }
        /**
         * Occurs when a row edit has been committed or canceled.
         */
        rowEditEnded = new Event();
        /**
         * Raises the @see:rowEditEnded event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onRowEditEnded(e: CellRangeEventArgs) {
            this.rowEditEnded.raise(this, e);
        }
        /**
         * Occurs when the user creates a new item by editing the new row template
         * (see the @see:allowAddNew property).
         *
         * The event handler may customize the content of the new item or cancel
         * the new item creation.
         */
        rowAdded = new Event();
        /**
         * Raises the @see:rowAdded event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onRowAdded(e: CellRangeEventArgs) {
            this.rowAdded.raise(this, e);
        }
        /**
         * Occurs when the user is deleting a selected row by pressing the Delete 
         * key (see the @see:allowDelete property).
         *
         * The event handler may cancel the row deletion.
         */
        deletingRow = new Event();
        /**
         * Raises the @see:deletingRow event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onDeletingRow(e: CellRangeEventArgs): boolean {
            this.deletingRow.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs after the user has deleted a row by pressing the Delete 
         * key (see the @see:allowDelete property).
         */
        deletedRow = new Event();
        /**
         * Raises the @see:deletedRow event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onDeletedRow(e: CellRangeEventArgs) {
            this.deletedRow.raise(this, e);
        }
        /**
         * Occurs when the user is copying the selection content to the 
         * clipboard by pressing one of the clipboard shortcut keys
         * (see the @see:autoClipboard property).
         *
         * The event handler may cancel the copy operation.
         */
        copying = new Event();
        /**
         * Raises the @see:copying event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onCopying(e: CellRangeEventArgs): boolean {
            this.copying.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs after the user has copied the selection content to the 
         * clipboard by pressing one of the clipboard shortcut keys
         * (see the @see:autoClipboard property).
         */
        copied = new Event();
        /**
         * Raises the @see:copied event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onCopied(e: CellRangeEventArgs) {
            this.copied.raise(this, e);
        }
        /**
         * Occurs when the user is pasting content from the clipboard 
         * by pressing one of the clipboard shortcut keys
         * (see the @see:autoClipboard property).
         *
         * The event handler may cancel the copy operation.
         */
        pasting = new Event();
        /**
         * Raises the @see:pasting event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onPasting(e: CellRangeEventArgs): boolean {
            this.pasting.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs after the user has pasted content from the 
         * clipboard by pressing one of the clipboard shortcut keys
         * (see the @see:autoClipboard property).
         */
        pasted = new Event();
        /**
         * Raises the @see:pasted event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onPasted(e: CellRangeEventArgs) {
            this.pasted.raise(this, e);
        }
        /**
         * Occurs when the user is pasting content from the clipboard 
         * into a cell (see the @see:autoClipboard property).
         *
         * The event handler may cancel the copy operation.
         */
        pastingCell = new Event();
        /**
         * Raises the @see:pastingCell event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onPastingCell(e: CellRangeEventArgs): boolean {
            this.pastingCell.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs after the user has pasted content from the 
         * clipboard into a cell (see the @see:autoClipboard property).
         */
        pastedCell = new Event();
        /**
         * Raises the @see:pastedCell event.
         *
         * @param e @see:CellRangeEventArgs that contains the event data.
         */
        onPastedCell(e: CellRangeEventArgs) {
            this.pastedCell.raise(this, e);
        }
        /**
         * Occurs when an element representing a cell has been created.
         *
         * This event can be used to format cells for display. It is similar 
         * in purpose to the @see:itemFormatter property, but has the advantage 
         * of allowing multiple independent handlers.
         *
         * For example, this code removes the 'wj-wrap' class from cells in 
         * group rows:
         *
         * <pre>flex.formatItem.addHandler(function (s, e) {
         *   if (flex.rows[e.row] instanceof wijmo.grid.GroupRow) {
         *     wijmo.removeClass(e.cell, 'wj-wrap');
         *   }
         * });</pre>
         */
        formatItem = new Event();
        /**
         * Raises the @see:formatItem event.
         *
         * @param e @see:FormatItemEventArgs that contains the event data.
         */
        onFormatItem(e: FormatItemEventArgs) {
            this.formatItem.raise(this, e);
        }
        /**
         * Occurs when the grid starts creating/updating the elements that
         * make up the current view.
         */
        updatingView = new Event();
        /**
         * Raises the @see:updatingView event.
         *
         * @param e @see:CancelEventArgs that contains the event data.
         * @return True if the event was not canceled.
         */
        onUpdatingView(e: CancelEventArgs): boolean {
            this.updatingView.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs when the grid finishes creating/updating the elements that 
         * make up the current view.
         *
         * The grid updates the view in response to several actions, including:
         *
         * <ul>
         * <li>refreshing the grid or its data source,</li>
         * <li>adding, removing, or changing rows or columns,</li>
         * <li>resizing or scrolling the grid,</li>
         * <li>changing the selection.</li>
         * </ul>
         */
        updatedView = new Event();
        /**
         * Raises the @see:updatedView event.
         */
        onUpdatedView(e?: EventArgs) {
            this.updatedView.raise(this, e);
        }

        //#endregion

        //--------------------------------------------------------------------------
        //#region ** implementation

        // measure the control's default row height based on current styles
        _getDefaultRowHeight() {

            // add control's root element to document body if it's detached 
            // (so CSS rules apply to our measurements)
            var host = this.hostElement,
                body = document.body,
                root = null;
            if (body && !contains(body, host)) {
                for (var p = host; p; p = p.parentElement) {
                    root = p;
                }
                if (root) {
                    body.appendChild(root);
                }
            }

            // create cell, measure, and remove it
            // (add extra 2px to keep default compatible with old algorithm)
            var cell = createElement('<div class="wj-cell">123</div>', host),
                defRowHei = cell.scrollHeight + 2;
            host.removeChild(cell);

            // remove root if we added it earlier
            if (root) {
                body.removeChild(root);
            }

            // make 100% sure we have a reasonable default row height!
            if (defRowHei <= 6 || isNaN(defRowHei) || !body) {
                defRowHei = 28;
            }

            // done, return what we have
            return defRowHei;
        }

        // gets the collection view associated with an itemsSource object
        _getCollectionView(value: any): collections.ICollectionView {
            return asCollectionView(value);
        }

        // measures the desired width of a cell
        _getDesiredWidth(p: GridPanel, r: number, c: number, e: HTMLElement) {
            var rng = this.getMergedRange(p, r, c);
            this.cellFactory.updateCell(p, r, c, e, rng);
            e.style.width = '';
            var w = e.offsetWidth;
            return rng && rng.columnSpan > 1
                ? w / rng.columnSpan
                : w;
        }

        // measures the desired height of a cell
        _getDesiredHeight(p: GridPanel, r: number, c: number, e: HTMLElement) {
            var rng = this.getMergedRange(p, r, c);
            this.cellFactory.updateCell(p, r, c, e, rng);
            e.style.height = '';
            var h = e.offsetHeight;
            return rng && rng.rowSpan > 1
                ? h / rng.rowSpan
                : h;
        }

        // gets the index of the sort row, with special handling for nulls
        _getSortRowIndex(): number {
            return this._sortRowIndex != null
                ? this._sortRowIndex
                : this.columnHeaders.rows.length - 1;
        }

        // sort converter used to sort mapped columns by display value
        _mappedColumns = null;
        private _sortConverter(sd: collections.SortDescription, item: any, value: any, init: boolean) {
            var col: Column;

            // initialize mapped column dictionary
            if (init) {
                this._mappedColumns = null;
                if (this.collectionView) {
                    var sds = this.collectionView.sortDescriptions;
                    for (var i = 0; i < sds.length; i++) {
                        col = this.columns.getColumn(sds[i].property);
                        if (col && col.dataMap) {
                            if (!this._mappedColumns) {
                                this._mappedColumns = {};
                            }
                            this._mappedColumns[col.binding] = col.dataMap;
                        }
                    }
                }

                // prioritize the column that was clicked
                // (in case multiple columns map the same property)
                if (this._mouseHdl._htDown && this._mouseHdl._htDown.col > -1) {
                    col = this.columns[this._mouseHdl._htDown.col];
                    if (this._mappedColumns && col.dataMap) { // TFS 118453
                        this._mappedColumns[col.binding] = col.dataMap;
                    }
                }
            }

            // convert value if we have a map
            if (this._mappedColumns) {
                var map = <DataMap>this._mappedColumns[sd.property];
                if (map && map.sortByDisplayValues) {
                    value = map.getDisplayValue(value);
                }
            }

            // return the value to use for sorting
            return value;
        }

        // binds the grid to the current data source.
        _bindGrid(full: boolean) {
            this.deferUpdate(() => {

                // do a full binding if we didn't have any data when we did it the first time
                if (this._lastCount == 0 && this._cv && this._cv.items && this._cv.items.length) {
                    full = true;
                }

                // save selected state
                var selItems = [];
                if (this.preserveSelectedState && this.selectionMode == SelectionMode.ListBox) {
                    for (var i = 0; i < this.rows.length; i++) {
                        var row = this.rows[i];
                        if (row.isSelected && row.dataItem) {
                            selItems.push(row.dataItem);
                        }
                    }
                }

                // save collapsed state
                var map;
                if (this.preserveOutlineState && isFunction(window['Map']) && this.rows.maxGroupLevel > -1) {
                    map = new Map();
                    for (var i = 0; i < this.rows.length; i++) {
                        var gr = <GroupRow>this.rows[i];
                        if (gr instanceof GroupRow && gr.isCollapsed && gr.dataItem) {
                            var key = gr.dataItem;
                            if (key instanceof wijmo.collections.CollectionViewGroup) {
                                key = key._path;
                            }
                            map.set(key, true);
                        }
                    }
                }

                // update columns
                if (full) {
                    this.columns.deferUpdate(() => {
                        this._bindColumns();
                    });
                }

                // update rows
                this.rows.deferUpdate(() => {
                    this._bindRows();
                });

                // restore/initialize ListBox selection
                var cnt = 0;
                if (selItems.length) {
                    for (var i = 0; i < this.rows.length && cnt < selItems.length; i++) {
                        if (selItems.indexOf(this.rows[i].dataItem) > -1) {
                            this.rows[i].isSelected = true;
                            cnt++;
                        }
                    }
                }

                // failed to restore ListBox selection by object, update by index
                if (this.selectionMode == SelectionMode.ListBox && cnt == 0) {
                    var sel = this.selection;
                    for (var i = sel.topRow; i <= sel.bottomRow && i > -1 && i < this.rows.length; i++) {
                        this.rows[i].isSelected = true;
                    }
                }

                // restore collapsed state
                if (map) {
                    this.rows.deferUpdate(() => {
                        for (var i = 0; i < this.rows.length; i++) {
                            var gr = <GroupRow>this.rows[i];
                            if (gr instanceof GroupRow) {
                                var key = gr.dataItem;
                                if (key instanceof wijmo.collections.CollectionViewGroup) {
                                    key = key._path;
                                }
                                if (map.get(key)) {
                                    gr.isCollapsed = true;
                                }
                            }
                        }
                    });
                }

                // save item count for next time
                if (!this._lastCount && this._cv && this._cv.items) {
                    this._lastCount = this._cv.items.length;
                }
            });

            // update selection
            if (this.collectionView) {
                this._cvCurrentChanged(this.collectionView, EventArgs.empty);
            }
        }

        // update grid rows to sync with data source
        /*protected*/ _cvCollectionChanged(sender, e: collections.NotifyCollectionChangedEventArgs) {

            // auto-generate if necessary
            if (this.autoGenerateColumns && this.columns.length == 0) {
                this._bindGrid(true);
                return;
            }

            // hierarchical binding: re-create all rows
            if (this.childItemsPath && e.action != collections.NotifyCollectionChangedAction.Change) {
                this._bindGrid(false);
                return;
            }

            // synchronize grid with updated CollectionView
            var index: number;
            switch (e.action) {

                // an item has changed, invalidate the grid to show the changes
                case collections.NotifyCollectionChangedAction.Change:
                    this.invalidate();
                    return;

                // an item has been added, insert a row
                case collections.NotifyCollectionChangedAction.Add:
                    if (e.index == this.collectionView.items.length - 1) {
                        index = this.rows.length;
                        if (this.rows[index - 1] instanceof _NewRowTemplate) {
                            index--;
                        }
                        this.rows.insert(index, new Row(e.item));
                        return;
                    }
                    assert(false, 'added item should be the last one.');
                    break;

                // an item has been removed, delete the row
                case collections.NotifyCollectionChangedAction.Remove:
                    var index = this._findRow(e.item);
                    if (index > -1) {
                        this.rows.removeAt(index);
                        this._cvCurrentChanged(sender, e);
                        return;
                    }
                    assert(false, 'removed item not found in grid.');
                    break;
            }

            // reset (sort, new source, etc): re-create all rows
            this._bindGrid(false);
        }

        // update selection to sync with data source
        private _cvCurrentChanged(sender, e) {
            if (this.collectionView) {

                // get grid's current item
                var sel = this.selection,
                    item = sel.row > -1 && sel.row < this.rows.length ? this.rows[sel.row].dataItem : null;

                // groups are not regular data items (TFS 142470)
                if (item instanceof wijmo.collections.CollectionViewGroup) {
                    item = null;
                }

                // if it doesn't match the view's, move the selection to match
                if (item != this.collectionView.currentItem) {
                    sel.row = sel.row2 = this._getRowIndex(this.collectionView.currentPosition);
                    this.select(sel, false);
                    if (this.selectionMode != SelectionMode.None) {
                        this.scrollIntoView(sel.row, -1);
                    }
                }
            }
        }

        // convert CollectionView index to row index
        private _getRowIndex(index: number): number {
            if (this.collectionView) {

                // look up item, then scan rows to find it
                if (index > -1) {
                    var item = this.collectionView.items[index];
                    for (; index < this.rows.length; index++) {
                        if (this.rows[index].dataItem === item) {
                            return index;
                        }
                    }
                    return -1; // item not found, shouldn't happen!
                } else {

                    // empty grid except for new row template? select that
                    if (this.rows.length == 1 && this.rows[0] instanceof _NewRowTemplate) {
                        return 0;
                    }

                    // no item to look up, so return current unbound row (group header)
                    // or -1 (no selection)
                    var index = this.selection.row,
                        row = index > -1 ? this.rows[index] : null;
                    return row && (row instanceof GroupRow || row.dataItem == null)
                        ? index
                        : -1;
                }
            }

            // not bound
            return this.selection.row;
        }

        // convert row index to CollectionView index
        _getCvIndex(index: number): number {
            if (index > -1 && this.collectionView) {
                var item = this.rows[index].dataItem;
                index = Math.min(index, this.collectionView.items.length);
                for (; index > -1; index--) {
                    if (this.collectionView.items[index] === item) {
                        return index;
                    }
                }
            }
            return -1;
        }

        // gets the index of the row that represents a given data item
        private _findRow(data: any): number {
            for (var i = 0; i < this.rows.length; i++) {
                if (this.rows[i].dataItem == data) {
                    return i;
                }
            }
            return -1;
        }

        // re-arranges the child HTMLElements within this grid.
        private _updateLayout() {

            // raise updatingLayout event
            var e = new CancelEventArgs();
            if (!this.onUpdatingLayout(e)) {
                return;
            }

            // compute content height, max height supported by browser,
            // and max offset so things match up when you scroll all the way down.
            var tlw = (this._hdrVis & HeadersVisibility.Row) ? this._hdrCols.getTotalSize() : 0,
                tlh = (this._hdrVis & HeadersVisibility.Column) ? this._hdrRows.getTotalSize() : 0,
                blh = this._ftrRows.getTotalSize(),
                heightReal = this._rows.getTotalSize() + blh;

            // make sure scrollbars are functional even if we have no rows (TFS 110441)
            if (heightReal < 1) {
                heightReal = 1;
            }

            // keep track of relevant variables
            this._rtl = this.hostElement ? (getComputedStyle(this.hostElement).direction == 'rtl') : false;
            this._heightBrowser = Math.min(heightReal, FlexGrid._getMaxSupportedCssHeight());
            this._maxOffsetY = Math.max(0, heightReal - this._heightBrowser);

            // compute default cell padding
            if (this.cells.hostElement) {
                var cell = createElement('<div class="wj-cell"></div>', this.cells.hostElement),
                    cs = getComputedStyle(cell);
                this._cellPadding = parseInt(this._rtl ? cs.paddingRight : cs.paddingLeft);
                cell.parentElement.removeChild(cell);
            }

            // top of the footer divs
            var ftrTop = this._heightBrowser + tlh - blh;

            // set sizes that do *not* depend on scrollbars being visible
            if (this._rtl) {
                setCss(this._eTL, { right: 0, top: 0, width: tlw, height: tlh });
                setCss(this._eCHdr, { right: tlw, top: 0, height: tlh });
                setCss(this._eRHdr, { right: 0, top: tlh, width: tlw });
                setCss(this._eCt, { right: tlw, top: tlh, width: this._gpCells.width, height: this._heightBrowser });
                setCss(this._fCt, { right: tlw, top: tlh });
                setCss(this._eBL, { right: 0, top: ftrTop, width: tlw, height: blh });
                setCss(this._eCFtr, { right: tlw, top: ftrTop, height: blh });
            } else {
                setCss(this._eTL, { left: 0, top: 0, width: tlw, height: tlh });
                setCss(this._eCHdr, { left: tlw, top: 0, height: tlh });
                setCss(this._eRHdr, { left: 0, top: tlh, width: tlw });
                setCss(this._eCt, { left: tlw, top: tlh, width: this._gpCells.width, height: this._heightBrowser });
                setCss(this._fCt, { left: tlw, top: tlh });
                setCss(this._eBL, { left: 0, top: ftrTop, width: tlw, height: blh });
                setCss(this._eCFtr, { left: tlw, top: ftrTop, height: blh });
            }

            // adjust for sticky headers
            if (this._stickyHdr) {
                this._updateStickyHeaders();
            }

            // update auto-sizer element
            var root = this._root,
                sbW = root.offsetWidth - root.clientWidth,
                sbH = root.offsetHeight - root.clientHeight;
            setCss(this._eSz, {
                width: tlw + sbW + this._gpCells.width,
                height: tlh + sbH + this._heightBrowser
            });

            // update star sizes and re-adjust content width to handle round-offs
            var clientWidth = null;
            if (this.columns._updateStarSizes(root.clientWidth - tlw)) {
                clientWidth = root.clientWidth;
                setCss(this._eCt, { width: this._gpCells.width });
            }

            // store control size
            this._szClient = new Size(root.clientWidth - tlw, root.clientHeight - tlh);
            this._rcBounds = null;

            // refresh content
            this._updateContent(false);

            // update auto-sizer after refreshing content
            sbW = root.offsetWidth - root.clientWidth;
            sbH = root.offsetHeight - root.clientHeight;
            setCss(this._eSz, {
                width: tlw + sbW + this._gpCells.width,
                height: tlh + sbH + this._heightBrowser
            });

            // update client size after refreshing content
            this._szClient = new Size(root.clientWidth - tlw, root.clientHeight - tlh);

            // adjust star sizes for vertical scrollbars
            if (clientWidth && clientWidth != root.clientWidth) {
                if (this.columns._updateStarSizes(root.clientWidth - tlw)) {
                    setCss(this._eCt, { width: this._gpCells.width });
                    this._updateContent(false);
                }
            }

            // set sizes that *do* depend on scrollbars being visible
            setCss([this._eCHdr, this._eCFtr, this._fCt], { width: this._szClient.width });
            setCss([this._eRHdr, this._fCt], { height: this._szClient.height });

            // adjust top of footer 
            if (blh) {
                ftrTop = Math.min(ftrTop, this._szClient.height + tlh - blh);
                setCss([this._eBL, this._eCFtr], { top: ftrTop });
            }

            // raise the event
            this.onUpdatedLayout(e);
        }

        // update the top of the header elements to remain visible 
        // when the user scrolls the window
        _updateStickyHeaders() {
            var stuck = false,
                offset = 0;

            // calculate offset
            if (this._stickyHdr) {
                var maxTop = 0,
                    thisTop = null;
                for (var el = this.hostElement; el; el = el.parentElement) {
                    var rc = el.getBoundingClientRect();
                    if (thisTop == null) {
                        thisTop = rc.top;
                    }
                    maxTop = Math.max(maxTop, rc.top);
                }
                thisTop = Math.max(0, maxTop - thisTop - 1);
                offset = -thisTop;
                stuck = thisTop > 0;
            }

            // apply offset
            this._eTL.style.top = this._eCHdr.style.top = stuck ? (-offset + 'px') : '';
            toggleClass(this._eTL, FlexGrid._WJS_STICKY, stuck);
            toggleClass(this._eCHdr, FlexGrid._WJS_STICKY, stuck);
        }

        // updates the scrollPosition property based on the element's scroll position
        // note that IE/Chrome/FF handle scrollLeft differently under RTL:
        // - Chrome reverses direction,
        // - FF uses negative values, 
        // - IE does the right thing (nothing)
        private _updateScrollPosition() : boolean {
            var root = this._root,
                top = root.scrollTop,
                left = root.scrollLeft;
            if (this._rtl && FlexGrid._getRtlMode() == 'rev') {
                left = (root.scrollWidth - root.clientWidth) - left;
            }
            var pt = new Point(-Math.abs(left), -top);

            // save new value and raise event
            if (!this._ptScrl.equals(pt)) {
                this._ptScrl = pt;
                this.onScrollPositionChanged();
                return true;
            }

            // no change...
            return false;
        }

        // updates the cell elements within this grid.
        private _updateContent(recycle: boolean, state?: boolean) {
            var focus = this.containsFocus(),
                hdrFocus = contains(this.columnHeaders.hostElement, getActiveElement());

            // raise updatingView event
            var e = new CancelEventArgs();
            if (!this.onUpdatingView(e)) {
                return;
            }

            // calculate offset to work around IE limitations
            this._offsetY = 0;
            if (this._heightBrowser > this._szClient.height) {
                var pct = Math.round((-this._ptScrl.y) / (this._heightBrowser - this._szClient.height) * 100) / 100;
                this._offsetY = Math.round(this._maxOffsetY * pct);
            }

            // update scroll position and then cells (TFS 144263, 152757)
            this._updateScrollPosition();
            this._gpCells._updateContent(recycle, state, this._offsetY);

            // update visible headers
            if (this._hdrVis & HeadersVisibility.Column) {
                if (!state || (this._ssHdr & HeadersVisibility.Column)) {
                    this._gpCHdr._updateContent(recycle, state, 0);
                }
            }
            if (this._hdrVis & HeadersVisibility.Row) {
                if (!state || (this._ssHdr & HeadersVisibility.Row)) {
                    this._gpRHdr._updateContent(recycle, state, this._offsetY);
                }
            }
            if (this._hdrVis && !state) {
                this._gpTL._updateContent(recycle, state, 0);
            }

            // update column footers
            if (this._gpCFtr.rows.length) {
                this._gpBL._updateContent(recycle, state, 0);
                this._gpCFtr._updateContent(recycle, state, 0)
            }
            // update marquee position
            if (this.showMarquee) {
                var sel = this._selHdl._sel,
                    marquee = this._eMarquee;
                if (!this.isRangeValid(sel)) { // TFS 139220
                    setCss(marquee, {
                        left: 0,
                        top: 0,
                        width: 0,
                        height: 0,
                        visibility: 'collapse'
                    });
                } else {
                    var rcm = this._getMarqueeRect(sel),
                        mc = <HTMLElement>marquee.firstChild,
                        dx = marquee.offsetWidth - mc.offsetWidth,
                        dy = marquee.offsetHeight - mc.offsetHeight;
                    setCss(marquee, {
                        left: rcm.left + this.cells.hostElement.offsetLeft - dx / 2,
                        top: rcm.top + this.cells.hostElement.offsetTop - dy  / 2,
                        width: rcm.width + dx,
                        height: rcm.height + dy,
                        visibility: rcm.width > 0 && rcm.height > 0 ? '' : 'collapse'
                    });
                }
            }

            // copy frozen cells into their own container (in IE only)
            if (isIE()) {
                setText(this._fCt, null);
                if (!this.activeEditor) {
                    if (this.frozenRows || this.frozenColumns) {
                        var frozen = this._eCt.querySelectorAll('.wj-frozen');
                        for (var i = 0; i < frozen.length; i++) {
                            this._fCt.appendChild(frozen[i].cloneNode(true));
                        }
                    }
                }
            }

            // restore focus
            if (focus && !state) {
                setTimeout(() => { // to work with Android (see HeaderFilters demo)
                    this.focus();
                }, 10);
            }

            // make sure hit-testing works
            this._rcBounds = null;

            // done updating the view
            this.onUpdatedView(e);
        }

        // get marquee rectangle (accounting for merging, freezing, RTL)
        private _getMarqueeRect(rng: CellRange): Rect {

            // get selection corner cells (accounting for merging)
            var m1 = this.getMergedRange(this.cells, rng.topRow, rng.leftCol) || new CellRange(rng.topRow, rng.leftCol),
                m2 = this.getMergedRange(this.cells, rng.bottomRow, rng.rightCol) || new CellRange(rng.bottomRow, rng.rightCol);

            // get cell client rectangles
            var rc1 = this.cells.getCellBoundingRect(m1.topRow, m1.leftCol, true),
                rc2 = this.cells.getCellBoundingRect(m2.bottomRow, m2.rightCol, true);

            // adjust for frozen rows
            if (this.rows.frozen) {
                var fzr = Math.min(this.rows.length, this.rows.frozen),
                    rcf = this.cells.getCellBoundingRect(fzr - 1, 0, true);
                if (rng.bottomRow >= fzr && rc2.bottom < rcf.bottom) {
                    return new Rect(0, 0, 0, 0);
                }
                if (rng.topRow >= fzr && rc1.top < rcf.bottom) {
                    rc1.top = rcf.bottom;
                }
            }

            // adjust for frozen columns
            if (this.columns.frozen) {
                var fzc = Math.min(this.columns.length, this.columns.frozen),
                    rcf = this.cells.getCellBoundingRect(0, fzc - 1, true);
                if (this._rtl) {
                    if (rng.rightCol >= fzc && rc2.left > rcf.left) {
                        return new Rect(0, 0, 0, 0);
                    }
                    if (rng.leftCol >= fzc && rc1.right > rcf.left) {
                        rc1.left = rcf.left - rc1.width;
                    }
                } else {
                    if (rng.rightCol >= fzc && rc2.right < rcf.right) {
                        return new Rect(0, 0, 0, 0);
                    }
                    if (rng.leftCol >= fzc && rc1.left < rcf.right) {
                        rc1.left = rcf.right;
                    }
                }
            }

            // return marquee rect
            return this._rtl
                ? new Rect(rc2.left, rc1.top, rc1.right - rc2.left, rc2.bottom - rc1.top)
                : new Rect(rc1.left, rc1.top, rc2.right - rc1.left, rc2.bottom - rc1.top);
        }

        // bind columns
        /*protected*/ _bindColumns() {

            // remove old auto-generated columns
            for (var i = 0; i < this.columns.length; i++) {
                var col = this.columns[i];
                if (col._getFlag(RowColFlags.AutoGenerated)) {
                    this.columns.removeAt(i);
                    i--;
                }
            }

            // get first item to infer data types
            var item = null,
                cv = this.collectionView;
            if (cv && cv.sourceCollection && cv.sourceCollection.length) {
                item = cv.sourceCollection[0];
            }

            // auto-generate new columns
            // (skipping unwanted types: array and object)
            if (item && this.autoGenerateColumns) {
                for (var key in item) {
                    if (isPrimitive(item[key])) {
                        col = new Column();
                        col._setFlag(RowColFlags.AutoGenerated, true);
                        col.binding = col.name = key;
                        col.header = toHeaderCase(key);
                        col.dataType = getType(item[key]);
                        if (col.dataType == DataType.Number) {
                            col.width = 80;
                        }
                        var pdesc = Object.getOwnPropertyDescriptor(item, key);
                        if (pdesc && !pdesc.writable && !isFunction(pdesc.set)) {
                            col._setFlag(RowColFlags.ReadOnly, true);
                        }
                        this.columns.push(col);
                    }
                }
            }

            // update missing column types
            this._updateColumnTypes();

            // REVIEW: add onLoading/edRows()?
        }

        // update missing column types to match data
        /*protected*/ _updateColumnTypes() {
            var cv = this.collectionView;
            if (hasItems(cv)) {
                var item = cv.items[0],
                    cols = this.columns;
                for (var i = 0; i < cols.length; i++) {
                    var col = cols[i];
                    if (col.dataType == null && col._binding) {
                        col.dataType = getType(col._binding.getValue(item));
                    }
                }
            }
        }

        // get the binding column 
        // (in the MultiRow grid, each physical column may contain several binding columns)
        /*protected*/ _getBindingColumn(p: GridPanel, r: Number, c: Column): Column {
            return c;
        }

        // bind rows
        /*protected*/ _bindRows() {

            // raise loading rows event
            var e = new CancelEventArgs();
            if (!this.onLoadingRows(e)) {
                return;
            }

            // clear rows
            this.rows.clear();

            // re-populate
            var cv = this.collectionView;
            if (cv && cv.items) {
                var list = cv.items;
                var groups = cv.groups;

                // bind to hierarchical sources (childItemsPath)
                if (this.childItemsPath) {
                    for (var i = 0; i < list.length; i++) {
                        this._addNode(list, i, 0);
                    }

                // bind to grouped sources
                } else if (groups != null && groups.length > 0 && this.showGroups) {
                    for (var i = 0; i < groups.length; i++) {
                        this._addGroup(groups[i]);
                    }

                // bind to regular sources
                } else {
                    for (var i = 0; i < list.length; i++) {
                        this._addBoundRow(list, i);
                    }
                }
            }

            // done binding rows
            this.onLoadedRows(e);
        }
        /*protected*/ _addBoundRow(items: any[], index: number) {
            this.rows.push(new Row(items[index]));
        }
        /*protected*/ _addNode(items: any[], index: number, level: number) {
            var gr = new GroupRow(),
                path = this.childItemsPath,
                prop = isArray(path) ? path[level] : path,
                item = items[index],
                children = item[prop];

            // add main node
            gr.dataItem = item;
            gr.level = level;
            this.rows.push(gr);

            // add child nodes
            if (children) {
                for (var i = 0; i < children.length; i++) {
                    this._addNode(children, i, level + 1);
                }
            }
        }
        private _addGroup(g: collections.CollectionViewGroup) {

            // add group row
            var gr = new GroupRow();
            gr.level = g.level;
            gr.dataItem = g;
            this.rows.push(gr);

            // add child rows
            if (g.isBottomLevel) {
                var items = g.items;
                for (var i = 0; i < items.length; i++) {
                    this._addBoundRow(items, i);
                }
            } else {
                for (var i = 0; i < g.groups.length; i++) {
                    this._addGroup(g.groups[i]);
                }
            }
        }

        // gets a list of the properties defined by a class and its ancestors
        // that have getters, setters, and whose names don't start with '_'.
        private static _getSerializableProperties(obj: any) {
            var arr = [];

            // travel up class hierarchy saving public properties that can be get/set.
            // NOTE: use getPrototypeOf instead of __proto__ for IE9 compatibility.
            for (obj = obj.prototype; obj != Object.prototype; obj = Object.getPrototypeOf(obj)) {
                var names = Object.getOwnPropertyNames(obj);
                for (var i = 0; i < names.length; i++) {
                    var name = names[i],
                        pd = Object.getOwnPropertyDescriptor(obj, name);
                    if (pd && pd.set && pd.get && name[0] != '_' &&
                        !name.match(/disabled|required/)) { // deprecated properties
                        arr.push(name);
                    }
                }
            }

            // done
            return arr;
        }

        // method used in JSON-style initialization
        _copy(key: string, value: any): boolean {
            if (key == 'columns') {
                this.columns.clear();
                var arr = asArray(value);
                for (var i = 0; i < arr.length; i++) {
                    var c = new Column();
                    copy(c, arr[i]);
                    this.columns.push(c);
                }
                return true;
            }
            return false;
        }

        // checked whether an object is an input element
        _isInputElement(e: any): boolean {
            if (e instanceof HTMLElement) {
                var m = (<HTMLElement>e).tagName.match(/INPUT|SELECT|TEXTAREA|BUTTON|A/);
                return m && m.length > 0;
            }
            return false;
        }

        // get max supported element height
        // IE limits it to about 1.5M, FF to 6M, Chrome to 30M
        private static _maxCssHeight: number;
        private static _getMaxSupportedCssHeight() : number {
            if (!FlexGrid._maxCssHeight) {
                var maxHeight = 1e6,
                    testUpTo = 60e6,
                    div = document.createElement('div');
                div.style.visibility = 'hidden';
                document.body.appendChild(div);
                for (var test = maxHeight; test <= testUpTo; test += 500000) {
                    div.style.height = test + 'px';
                    if (div.offsetHeight != test) {
                        break;
                    }
                    maxHeight = test;
                }
                document.body.removeChild(div);
                FlexGrid._maxCssHeight = maxHeight;
            }
            return FlexGrid._maxCssHeight;
        }

        // IE/Chrome/Safari/FireFox handle RTL differently; this function
        // returns a value that indicates the RTL mode:
        // 'rev': reverse direction so origin is +max, end is zero (Chrome/Safari)
        // 'neg': switches signs, so origin is -max, end is zero (Firefox)
        // 'std': standard, origin is zero, end is +max (IE)
        static _rtlMode: string;
        private static _getRtlMode() : string {
            if (!FlexGrid._rtlMode) {
                var el = createElement(
                    '<div dir="rtl" style="visibility:hidden;width:100px;height:100px;overflow:auto">' +
                        '<div style="width:2000px;height:2000px"></div>' +
                    '</div>');

                document.body.appendChild(el);
                var sl = el.scrollLeft;
                el.scrollLeft = -1000;
                var sln = el.scrollLeft;
                document.body.removeChild(el);

                FlexGrid._rtlMode = sln < 0 ? 'neg' : sl > 0 ? 'rev' : 'std';
                //console.log('rtlMode: ' + FlexGrid._rtlMode);
            }
            return FlexGrid._rtlMode;
        }

        //#endregion
    }
}

module wijmo.grid {
    'use strict';

    /**
     * Provides arguments for @see:CellRange events.
     */
    export class CellRangeEventArgs extends CancelEventArgs {
        _p: GridPanel;
        _rng: CellRange;
        _data: any;

        /**
         * Initializes a new instance of the @see:CellRangeEventArgs class.
         *
         * @param p @see:GridPanel that contains the range.
         * @param rng Range of cells affected by the event.
         * @param data Data related to the event.
         */
        constructor(p: GridPanel, rng: CellRange, data?: string) {
            super();
            this._p = asType(p, GridPanel);
            this._rng = asType(rng, CellRange);
            this._data = data;
        }
        /**
         * Gets the @see:GridPanel affected by this event.
         */
        get panel(): GridPanel {
            return this._p;
        }
        /**
         * Gets the @see:CellRange affected by this event.
         */
        get range(): CellRange {
            return this._rng.clone();
        }
        /**
         * Gets the row affected by this event.
         */
        get row(): number {
            return this._rng.row;
        }
        /**
         * Gets the column affected by this event.
         */
        get col(): number {
            return this._rng.col;
        }
        /**
         * Gets or sets the data associated with the event.
         */
        get data(): any {
            return this._data;
        }
        set data(value: any) {
            this._data = value;
        }
    }

    /**
     * Provides arguments for the @see:FlexGrid.formatItem event.
     */
    export class FormatItemEventArgs extends CellRangeEventArgs {
        _cell: HTMLElement;

        /**
        * Initializes a new instance of the @see:FormatItemEventArgs class.
        *
        * @param p @see:GridPanel that contains the range.
        * @param rng Range of cells affected by the event.
        * @param cell Element that represents the grid cell to be formatted.
        */
        constructor(p: GridPanel, rng: CellRange, cell: HTMLElement) {
            super(p, rng);
            this._cell = asType(cell, HTMLElement);
        }
        /**
         * Gets a reference to the element that represents the grid cell to be formatted.
         */
        get cell(): HTMLElement {
            return this._cell;
        }
   }
}
module wijmo.grid {
    'use strict';

    /**
     * Specifies constants that define the type of cell in a @see:GridPanel.
     */
    export enum CellType {
        /** Unknown or invalid cell type. */
        None,
        /** Regular data cell. */
        Cell,
        /** Column header cell. */
        ColumnHeader,
        /** Row header cell. */
        RowHeader,
        /** Top-left cell. */
        TopLeft,
        /** Column footer cell. */
        ColumnFooter,
        /** Bottom left cell (at the intersection of the row header and column footer cells). **/
        BottomLeft
    }

    /**
     * Represents a logical part of the grid, such as the column headers, row headers,
     * and scrollable data part.
     */
    export class GridPanel {
        private _g: FlexGrid;
        private _ct: CellType;
        private _e: HTMLElement;
        private _rows: RowCollection;
        private _cols: ColumnCollection;
        private _offsetY = 0;
        private _rng: CellRange; // buffered view range
        private _rowIdx: number[]; // index of the last cell element in each row
        private static _evtBlur;

        /**
         * Initializes a new instance of the @see:GridPanel class.
         *
         * @param g The @see:FlexGrid object that owns the panel.
         * @param cellType The type of cell in the panel.
         * @param rows The rows displayed in the panel.
         * @param cols The columns displayed in the panel.
         * @param element The HTMLElement that hosts the cells in the control.
         */
        constructor(g: FlexGrid, cellType: CellType, rows: RowCollection, cols: ColumnCollection, element: HTMLElement) {
            this._g = asType(g, FlexGrid);
            this._ct = asInt(cellType);
            this._rows = asType(rows, RowCollection);
            this._cols = asType(cols, ColumnCollection);
            this._e = asType(element, HTMLElement);
            this._rng = new CellRange();

            // dispatch blur event for focused cells before recycling the panel
            if (!GridPanel._evtBlur) {
                GridPanel._evtBlur = document.createEvent('HTMLEvents');
                GridPanel._evtBlur.initEvent('blur', true, false);
            }
        }
        /**
         * Gets the grid that owns the panel.
         */
        get grid(): FlexGrid {
            return this._g;
        }
        /**
         * Gets the type of cell contained in the panel.
         */
        get cellType(): CellType {
            return this._ct;
        }
        /**
         * Gets a @see:CellRange that indicates the range of cells currently visible on the panel.
         */
        get viewRange(): CellRange {
            return this._getViewRange(false);
        }
        /**
         * Gets the total width of the content in the panel.
         */
        get width(): number {
            return this._cols.getTotalSize();
        }
        /**
         * Gets the total height of the content in this panel.
         */
        get height(): number {
            return this._rows.getTotalSize();
        }
        /**
         * Gets the panel's row collection.
         */
        get rows(): RowCollection {
            return this._rows;
        }
        /**
         * Gets the panel's column collection.
         */
        get columns(): ColumnCollection {
            return this._cols;
        }
        /**
         * Gets the value stored in a cell in the panel.
         *
         * @param r The row index of the cell.
         * @param c The index, name, or binding of the column that contains the cell.
         * @param formatted Whether to format the value for display.
         */
        getCellData(r: number, c: any, formatted: boolean): any {
            var row = <Row>this._rows[asNumber(r, false, true)],
                col: Column,
                value = null;

            // get column index by name or binding
            if (isString(c)) {
                c = this._cols.indexOf(c);
                if (c < 0) {
                    throw 'Invalid column name or binding.';
                }
            }

            // get column
            col = <Column>this._cols[asNumber(c, false, true)];

            // get binding column (MultiRow grid may have multiple display columns for each physical column)
            var bcol = this._g ? this._g._getBindingColumn(this, r, col) : col;

            // get bound value from data item using binding
            if (bcol.binding && row.dataItem &&
                !(row.dataItem instanceof collections.CollectionViewGroup)) { // TFS 108841
                value = bcol._binding.getValue(row.dataItem);
            } else if (row._ubv) { // get unbound value
                value = row._ubv[col._hash];
            }

            // special values for row and column headers, aggregates
            if (value == null) {
                switch (this._ct) {
                    case CellType.ColumnHeader:
                        if (r == this._rows.length - 1 || bcol != col) {
                            value = bcol.header;
                        }
                        break;
                    case CellType.ColumnFooter:
                        if (bcol.aggregate != Aggregate.None && row instanceof GroupRow) {
                            var icv = this._g.collectionView;
                            if (icv) {
                                var cv = <collections.CollectionView>tryCast(icv, collections.CollectionView);
                                value = cv
                                    ? cv.getAggregate(bcol.aggregate, bcol.binding)
                                    : getAggregate(bcol.aggregate, icv.items, bcol.binding);
                            }
                        }
                        break;
                    case CellType.Cell:
                        if (bcol.aggregate != Aggregate.None && row instanceof GroupRow) {
                            var group = <collections.CollectionViewGroup>tryCast(row.dataItem, collections.CollectionViewGroup);
                            if (group) {
                                value = group.getAggregate(bcol.aggregate, bcol.binding, this._g.collectionView);
                            }
                        }
                        break;
                }
            }

            // format value if requested, never return null
            if (formatted) {
                if (this.cellType == CellType.Cell && bcol.dataMap) {
                    value = bcol.dataMap.getDisplayValue(value);
                }
                value = value != null ? Globalize.format(value, bcol.format) : '';
            }

            // done
            return value;
        }
        /**
         * Sets the content of a cell in the panel.
         *
         * @param r The index of the row that contains the cell.
         * @param c The index, name, or binding of the column that contains the cell.
         * @param value The value to store in the cell.
         * @param coerce Whether to change the value automatically to match the column's data type.
         * @param invalidate Whether to invalidate the grid to show the change.
         * @return Returns true if the value is stored successfully, false otherwise (failed cast).
         */
        setCellData(r: number, c: any, value: any, coerce = true, invalidate = true): boolean {
            var row = <Row>this._rows[asNumber(r, false, true)],
                col: Column;

            // get column index by name or binding
            if (isString(c)) {
                c = this._cols.indexOf(c);
                if (c < 0) {
                    throw 'Invalid column name or binding.';
                }
            }

            // get column
            col = <Column>this._cols[asNumber(c, false, true)];

            // get binding column (MultiRow grid may have multiple display columns for each physical column)
            var bcol = this._g ? this._g._getBindingColumn(this, r, col) : col;

            // handle dataMap, coercion, type-checking
            if (this._ct == CellType.Cell) {

                // honor dataMap
                if (bcol.dataMap && value != null) {
                    if (bcol.isRequired || (value != '' && value != null)) { // TFS 107058
                        var map = bcol.dataMap,
                            key = map.getKeyValue(value);
                        if (key == null) {
                            if (!map.isEditable || map.displayMemberPath != map.selectedValuePath) {
                                return false; // not on map, not editable? cancel edits
                            }
                        } else {
                            value = key; // got the key, use it instead of the value
                        }
                    }
                }

                // get target type
                var targetType = DataType.Object;
                if (bcol.dataType) {
                    targetType = bcol.dataType;
                } else {
                    var current = this.getCellData(r, c, false);
                    targetType = getType(current);
                }

                // honor 'isRequired' property
                if (isBoolean(bcol.isRequired)) {
                    if (!bcol.isRequired && (value === '' || value === null)) {
                        value = null; // setting to null
                        coerce = false;
                    } else if (bcol.isRequired && (value === '' || value === null)) {
                        return false; // value is required
                    }
                }

                // coerce type if required
                if (coerce) {
                    value = changeType(value, targetType, bcol.format);
                    if (targetType != DataType.Object && getType(value) != targetType) {
                        return false; // wrong data type
                    }
                }
            }

            // store value
            if (row.dataItem && bcol.binding) {
                var binding = bcol._binding,
                    item = row.dataItem,
                    oldValue = binding.getValue(item);
                if (value !== oldValue && !DateTime.equals(value, oldValue)) {

                    // set the value
                    binding.setValue(item, value);

                    // track changes in CollectionView if this is not the current edit item (e.g. when pasting)
                    var view = <collections.CollectionView>this._g.collectionView;
                    if (view instanceof collections.CollectionView && item != view.currentEditItem) {
                        var e = new collections.NotifyCollectionChangedEventArgs(collections.NotifyCollectionChangedAction.Change, item, view.items.indexOf(item));
                        view.onCollectionChanged(e);
                    }
                }
            } else {
                if (!row._ubv) row._ubv = {};
                row._ubv[col._hash] = value;
            }

            // invalidate
            if (invalidate && this._g) {
                this._g.invalidate();
            }

            // done
            return true;
        }
        /**
         * Gets a cell's bounds in viewport coordinates.
         *
         * The returned value is a @see:Rect object which contains the position and dimensions 
         * of the cell in viewport coordinates.
         * The viewport coordinates are the same as those used by the 
         * <a href="https://developer.mozilla.org/en-US/docs/Web/API/Element.getBoundingClientRect"
         * target="_blank">getBoundingClientRect</a> method.
         *
         * @param r The index of the row that contains the cell.
         * @param c The index of the column that contains the cell.
         * @param raw Whether to return the rectangle in raw panel coordinates as opposed to viewport coordinates.
         */
        getCellBoundingRect(r: number, c: number, raw?: boolean): Rect {

            // get rect in panel coordinates
            var row = this.rows[r],
                col = this.columns[c],
                rc = new Rect(col.pos, row.pos, col.renderSize, row.renderSize);

            // adjust for rtl
            if (this._g._rtl) {
                rc.left = this.hostElement.clientWidth - rc.right;

                // account for scrollbars in non-ie browsers
                if (!isIE()) { 
                    var root = this.hostElement.parentElement;
                    rc.left -= root.offsetWidth - root.clientWidth;
                }
            }

            // adjust for panel position
            if (!raw) {
                var rcp = this.hostElement.getBoundingClientRect();
                rc.left += rcp.left;
                rc.top += rcp.top - this._offsetY;
            }

            // account for frozen rows/columns (TFS 105593)
            if (r < this.rows.frozen) {
                rc.top -= this._g.scrollPosition.y;
            }
            if (c < this.columns.frozen) {
                rc.left -= this._g.scrollPosition.x * (this._g._rtl ? -1 : +1);
            }

            // done
            return rc;
        }
        /**
         * Gets a @see:SelectedState value that indicates the selected state of a cell.
         *
         * @param r Row index of the cell to inspect.
         * @param c Column index of the cell to inspect.
         * @param rng @see:CellRange that contains the cell to inspect.
         */
        getSelectedState(r: number, c: number, rng: CellRange): SelectedState {
            var g = this._g,
                mode = g.selectionMode,
                sel = g._selHdl._sel;
            if (mode != SelectionMode.None) {
                switch (this._ct) {

                    // regular cells
                    case CellType.Cell:

                        // handle merged ranges
                        if (!rng) {
                            rng = g.getMergedRange(this, r, c);
                        }
                        if (rng) {
                            if (rng.contains(sel.row, sel.col)) {
                                return g.showMarquee ? SelectedState.None : SelectedState.Cursor;
                            } else if (rng.intersects(sel)) {
                                return SelectedState.Selected;
                            }
                        }

                        // cursor (if not showing marquee)
                        if (sel.row == r && sel.col == c) {
                            return g.showMarquee ? SelectedState.None : SelectedState.Cursor;
                        }

                        // special case: row/col selected property
                        if (g.rows[r].isSelected || g.columns[c].isSelected) {
                            return SelectedState.Selected;
                        }

                        // adjust for selection mode
                        sel = g._selHdl._adjustSelection(sel);

                        // ListBox mode (already checked for selected rows/cols)
                        if (mode == SelectionMode.ListBox) {
                            return SelectedState.None;
                        }

                        // regular ranges
                        return sel.containsRow(r) && sel.containsColumn(c)
                            ? SelectedState.Selected
                            : SelectedState.None;

                    // column headers
                    case CellType.ColumnHeader:
                        if (g.showSelectedHeaders & HeadersVisibility.Column) {
                            if (g.columns[c].isSelected || sel.containsColumn(c) || sel.intersectsColumn(rng)) {
                                if (rng) r = rng.bottomRow;
                                if (r == this.rows.length - 1) {
                                    return SelectedState.Selected;
                                }
                            }
                        }
                        break;

                    // row headers
                    case CellType.RowHeader:
                        if (g.showSelectedHeaders & HeadersVisibility.Row) {
                            if (g.rows[r].isSelected || sel.containsRow(r) || sel.intersectsRow(rng)) {
                                if (rng) c = rng.rightCol;
                                if (c == this.columns.length - 1) {
                                    return SelectedState.Selected;
                                }
                            }
                        }
                        break;
                }
            }

            // not selected
            return SelectedState.None;
        }
        /**
         * Gets the host element for the panel.
         */
        get hostElement(): HTMLElement {
            return this._e;
        }

        // ** implementation

        /* -- do not document, this is internal --
         * Gets the Y offset for cells in the panel.
         */
        _getOffsetY(): number {
            return this._offsetY;
        }

        /* -- do not document, this is internal --
         * Updates the cell elements in the panel.
         * @param recycle Whether to recycle existing elements or start from scratch.
         * @param state Whether to keep existing elements and update their state.
         * @param offsetY Scroll position to use when updating the panel.
         */
        _updateContent(recycle: boolean, state: boolean, offsetY: number) {
            var r: number, c: number,
                ctr: number,
                cell: HTMLElement,
                g = this._g,
                rows = this._rows,
                cols = this._cols,
                ct = this._ct;

            // scroll headers into position
            if (ct == CellType.ColumnHeader || ct == CellType.ColumnFooter || ct == CellType.RowHeader) {
                var sp = g._ptScrl,
                    s = this._e.style;
                if (ct == CellType.RowHeader) { // scroll row headers vertically
                    s.top = sp.y + 'px';
                } else { // scroll column headers/footers horizontally
                    if (g._rtl) {
                        s.right = sp.x + 'px';
                    } else {
                        s.left = sp.x + 'px';
                    }
                }
            }

            // update offset (and don't recycle if it changed!)
            if (this._offsetY != offsetY) {
                recycle = false;
                this._offsetY = offsetY;
            }

            // calculate new view range and buffered view range
            var vrng = this._getViewRange(false),
                rng = this._getViewRange(true); // (recycle && g.isTouching) ? this._getViewRange(true) : vrng;

            // done if recycling, not updating state, and old range contains new (unbuffered)
            // this happens a lot while scrolling by small amounts (< 1 cell)
            if (recycle && !state && !rows.frozen && !cols.frozen && this._rng.contains(vrng)) {
                return;
            }

            // if not recycling or if the range changed, ignore 'cells' refresh list
            if (!recycle || !rng.equals(this._rng)) {
                state = false;
            }

            // clear content if not recycling
            if (!recycle) {
                var ae = getActiveElement(),
                    eFocus = contains(this._e, ae) ? ae : null,
                    cf = this._g.cellFactory;

                // give cell factory a chance to dispose of the cells
                for (var i = 0; i < this._e.childElementCount; i++) {
                    cf.disposeCell(<HTMLElement>this._e.children[i]);
                }

                // clear content
                setText(this._e, null);

                // update focus state in case the editor was disposed
                if (eFocus) {
                    eFocus.dispatchEvent(GridPanel._evtBlur);
                }

                // clear row indices
                this._rowIdx = [];
            }

            // reorder cells to optimize scrolling (headers too)
            if (recycle && this._ct != CellType.TopLeft) {
                this._reorderCells(rng, this._rng);
            }

            // save new ranges
            this._rng = rng;

            // go create/update the cells
            // (render frozen cells last so we don't need z-index!)
            ctr = 0;
            this._rowIdx = [];
            for (r = rng.topRow; r <= rng.bottomRow && r > -1; r++) { // not frozen
                this._rowIdx.push(ctr);
                ctr = this._renderRow(r, rng, false, state, ctr);
            }
            this._rowIdx.push(ctr); // one past last cell
            for (r = rng.topRow; r <= rng.bottomRow && r > -1; r++) { // frozen col
                ctr = this._renderRow(r, rng, true, state, ctr);
            }
            for (r = 0; r < rows.frozen && r < rows.length; r++) { // frozen row
                ctr = this._renderRow(r, rng, false, state, ctr);
            }
            for (r = 0; r < rows.frozen && r < rows.length; r++) { // frozen row/col
                ctr = this._renderRow(r, rng, true, state, ctr);
            }

            // show the cells we are using, hide the others
            var cnt = this._e.childElementCount;
            for (var i = ctr; i < cnt; i++) {
                cell = <HTMLElement>this._e.children[i];
                cell.style.display = 'none';
            }
        }

        // reorder cells within the panel to optimize scrolling performance
        _reorderCells(rngNew: CellRange, rngOld: CellRange) {

            // sanity
            if (!this._rowIdx || this._rows.frozen > 0 || this._cols.frozen > 0 ||
                rngNew.columnSpan != rngOld.columnSpan || rngNew.rowSpan != rngOld.rowSpan ||
                !rngOld.isValid || !rngNew.isValid || !rngNew.intersects(rngOld)) {
                return;
            }

            // vertical scrolling
            if (rngNew.row != rngOld.row) {
                var delta = rngNew.row - rngOld.row,
                    limit = Math.max(1, rngNew.rowSpan - 1);
                if (delta != 0 && Math.abs(delta) < limit) {
                
                    // down:
                    // remove cells from the top and append to bottom
                    if (delta > 0) {
                        var rng = this._createRange(0, this._rowIdx[delta]);
                        this._e.appendChild(rng.extractContents());
                    }

                    // up:
                    // remove cells from the bottom and insert at the top
                    if (delta < 0) {
                        var last = this._rowIdx.length - 1,
                            rng = this._createRange(this._rowIdx[last + delta], this._rowIdx[last]);
                            this._e.insertBefore(rng.extractContents(), this._e.firstChild);
                        }
                    }
                }

            // horizontal scrolling
            if (rngNew.col != rngOld.col) {
                var delta = rngNew.col - rngOld.col,
                    limit = Math.max(1, rngNew.columnSpan - 1);
                if (delta != 0 && Math.abs(delta) < limit) {
    
                    // scan each row in the view and move the elements within
                    var cnt = this._e.childElementCount;
                    for (var i = 0; i < this._rowIdx.length - 1; i++) {
                        var row = this.rows[rngNew.topRow + i];
                        if (!(row instanceof GroupRow)) {

                            // get cell range for this row
                            var start = this._rowIdx[i],
                                end = this._rowIdx[i + 1];

                            // right:
                            // move first 'delta' elements to the end of the row range
                            if (delta > 0 && start + delta <= cnt) { // TFS 201620
                                var rng = this._createRange(start, start + delta);
                                this._e.insertBefore(rng.extractContents(), this._e.children[end]);
                            }

                            // left:
                            // move last 'delta' elements to the start of the row range
                            if (delta < 0 && end + delta - 1 >= 0) { // TFS 152906
                                var rng = this._createRange(end + delta - 1, end - 1);
                                this._e.insertBefore(rng.extractContents(), this._e.children[start]);
                            }
                        }
                    }
                }
            }
        }

        // creates a range of cells that can be moved to optimize rendering
        _createRange(start: number, end: number) {
            var rng = document.createRange();
            rng.setStart(this._e, start);
            rng.setEnd(this._e, end);
            return rng;
        }

        // renders a row
        _renderRow(r: number, rng: CellRange, frozen: boolean, state: boolean, ctr: number): number {

            // skip hidden rows
            if ((<Row>this.rows[r]).renderSize <= 0) {
                return ctr;
            }

            // render each cell in the row
            if (frozen) {
                for (var c = 0; c < this.columns.frozen && c < this.columns.length; c++) {
                    ctr = this._renderCell(r, c, rng, state, ctr);
                }
            } else {
                for (var c = <number>rng.leftCol; c <= rng.rightCol && c > -1; c++) {
                    ctr = this._renderCell(r, c, rng, state, ctr);
                }
            }

            // return updated counter
            return ctr;
        }

        // renders a cell
        _renderCell(r: number, c: number, rng: CellRange, state: boolean, ctr: number): number {

            // skip over cells that have been merged over
            var g = this._g,
                mrng = g.getMergedRange(this, r, c);
            if (mrng) {
                for (var over = Math.max(rng.row, mrng.row); over < r; over++) {
                    if (this.rows[over].isVisible) return ctr;
                }
                for (var over = Math.max(rng.col, mrng.col); over < c; over++) {
                    if (this.columns[over].isVisible) return ctr;
                }
            }

            // skip hidden and non-merged columns
            if ((<Column>this.columns[c]).renderSize <= 0) {
                if (!mrng || mrng.getRenderSize(this).width <= 0) {
                    return ctr;
                }
            }

            // try recycling a cell
            var cell = <HTMLElement>this._e.childNodes[ctr++];

            // update selected state
            if (cell && state) {
                var selState = this.getSelectedState(r, c, mrng);
                toggleClass(cell, 'wj-state-selected', selState == SelectedState.Cursor);
                toggleClass(cell, 'wj-state-multi-selected', selState == SelectedState.Selected);
                return ctr;
            }

            // create or recycle cell
            // NOTE: make cells focusable so we can give the focus to them rather
            // than to the grid, which causes browsers to try and scroll the whole
            // grid into view.
            if (!cell) {
                cell = document.createElement('div');
                cell.tabIndex = 0;
                this._e.appendChild(cell);
            }

            // set/update cell content/style
            g.cellFactory.updateCell(this, r, c, cell, mrng);

            // return updated counter
            return ctr;
        }

        // gets the range of cells currently visible,
        // optionally adding a buffer for inertial scrolling
        _getViewRange(buffer: boolean): CellRange {
            var g = this._g,
                sp = g._ptScrl,
                rows = this._rows,
                cols = this._cols,
                rng = new CellRange(0, 0, rows.length - 1, cols.length - 1);

            // calculate range
            if (this._ct == CellType.Cell || this._ct == CellType.RowHeader) {
                var y = -sp.y + this._offsetY,
                    h = g._szClient.height,
                    fz = Math.min(rows.frozen, rows.length - 1);

                // account for frozen rows
                if (fz > 0) {
                    var fzs = rows[fz - 1].pos;
                    y += fzs;
                    h -= fzs;
                }

                // set row range
                rng.row = Math.min(rows.length - 1, Math.max(rows.frozen, rows.getItemAt(y + 1)));
                rng.row2 = rows.getItemAt(y + h);
            }
            if (this._ct == CellType.Cell || this._ct == CellType.ColumnHeader) {
                var x = -sp.x,
                    w = g._szClient.width,
                    fz = Math.min(cols.frozen, cols.length - 1);

                // account for frozen columns
                if (fz > 0) {
                    var fzs = cols[fz - 1].pos;
                    x += fzs;
                    w -= fzs;
                }

                // set column range
                rng.col = Math.min(cols.length - 1, Math.max(cols.frozen, cols.getItemAt(x + 1)));
                rng.col2 = cols.getItemAt(x + w);
            }

            // add buffer
            if (buffer && this._ct == CellType.Cell) {
                var sz = g.isTouching ? 6 : 1;
                rng.row = Math.max(rng.row - sz, 0);
                rng.row2 = Math.min(rng.row2 + sz, rows.length - 1);
                rng.col = Math.max(rng.col - 1, 0);
                rng.col2 = Math.min(rng.col2 + 1, cols.length - 1);
            }

            // handle case where all rows/cols are frozen
            if (rows.length <= rows.frozen) {
                rng.row = rng.row2 = -1;
            }
            if (cols.length <= cols.frozen) {
                rng.col = rng.col2 = -1;
            }

            // return viewrange
            return rng;
        }

        // gets the point where the frozen area ends
        _getFrozenPos(): Point {
            var fzr = this._rows.frozen,
                fzc = this._cols.frozen,
                fzrow = fzr > 0 ? this._rows[fzr - 1] : null,
                fzcol = fzc > 0 ? this._cols[fzc - 1] : null,
                fzy = fzrow ? fzrow.pos + fzrow.renderSize : 0,
                fzx = fzcol ? fzcol.pos + fzcol.renderSize : 0;
            return new Point(fzx, fzy);
        }
    }
}

module wijmo.grid {
    'use strict';

    /**
     * Creates HTML elements that represent cells within a @see:FlexGrid control.
     */
    export class CellFactory {
        static _WJC_COLLAPSE = 'wj-elem-collapse';
        static _WJC_DROPDOWN = 'wj-elem-dropdown';
        static _ddIcon: HTMLElement;
        static _fmtRng: CellRange;

        /**
         * Creates or updates a cell in the grid.
         *
         * @param p The @see:GridPanel that contains the cell.
         * @param r The index of the row that contains the cell.
         * @param c The index of the column that contains the cell.
         * @param cell The element that represents the cell.
         * @param rng The @see:CellRange object that contains the cell's 
         * merged range, or null if the cell is not merged.
         * @param updateContent Whether to update the cell's content as
         * well as its position and style.
         */
        public updateCell(p: GridPanel, r: number, c: number, cell: HTMLElement, rng?: CellRange, updateContent?: boolean) {
            var g = p.grid,
                ct = p.cellType,
                rows = p.rows,
                cols = p.columns,
                row = rows[r],
                col = cols[c],
                r2 = r,
                c2 = c,
                gr = <GroupRow>tryCast(row, GroupRow),
                nr = tryCast(row, _NewRowTemplate),
                cellWidth = col.renderWidth,
                cellHeight = row.renderHeight,
                cl = 'wj-cell',
                css: any = { display: '' },
                canSkip = (updateContent != false); // don't skip if not updating content

            // clear cells that have child elements before re-using them
            // this is a workaround for a bug in IE that affects templates
            // strangely, setting the cell's innerHTML to '' doesn't help...
            if (updateContent != false && cell.firstElementChild &&
                (cell.childNodes.length != 1 || (<HTMLInputElement>cell.firstElementChild).type != 'checkbox')) {
                setText(cell, null);
                canSkip = false;
            }

            // adjust for merged ranges
            if (rng && !rng.isSingleCell) {
                r = rng.row;
                c = rng.col;
                r2 = rng.row2;
                c2 = rng.col2;
                row = rows[r];
                col = cols[c];
                gr = <GroupRow>tryCast(row, GroupRow);
                var sz = rng.getRenderSize(p);
                cellHeight = sz.height;
                cellWidth = sz.width;
            }

            // get column to use for binding (usually the same as col, but not on MultiRow)
            var bcol = g._getBindingColumn(p, r, col);

            // get cell position accounting for frozen rows/columns
            // in IE, frozen cells: will be moved to a fixed div, so adjust the scroll position for that
            // in other browsers, they will remain in the cells div, so adjust for that instead
            // (not when editing...)
            var cpos = col.pos,
                rpos = row.pos;
            if (isIE() && ct == CellType.Cell && !g.editRange) {
                if (r < rows.frozen && c >= cols.frozen) {
                    cpos += g._ptScrl.x;
                }
                if (c < cols.frozen && r >= rows.frozen) {
                    rpos += g._ptScrl.y;
                }
            } else { // header cells: just remove the scroll position
                if (r < rows.frozen) {
                    rpos -= g._ptScrl.y;
                }
                if (c < cols.frozen) {
                    cpos -= g._ptScrl.x;
                }
            }

            // size and position
            if (g._rtl) {
                css.right = cpos + 'px';
            } else {
                css.left = cpos + 'px';
            }
            css.top = (rpos - p._getOffsetY()) + 'px';
            css.width = cellWidth + 'px';
            css.height = cellHeight + 'px';

            // selector classes that only apply to regular cells
            if (ct == CellType.Cell) {
                if (gr) {
                    cl += ' wj-group';
                }
                if (g.showAlternatingRows && r % 2 != 0) {
                    cl += ' wj-alt';
                }
                if (r < rows.frozen || c < cols.frozen) {
                    cl += ' wj-frozen';
                }
                if (nr) {
                    cl += ' wj-new';
                }
                if (row.cssClass) {
                    cl += ' ' + row.cssClass;
                }
                if (bcol.cssClass) {
                    cl += ' ' + bcol.cssClass;
                }
            } else {
                cl += ' wj-header';
                if (g.showAlternatingRows && r % 2 != 0) {
                    cl += ' wj-header-alt';
                }
            }

            // selected state
            var selState = p.getSelectedState(r, c, rng);
            if (selState != SelectedState.None && ct == CellType.Cell &&
                col.dataType != DataType.Boolean && g.editRange && g.editRange.contains(r, c)) {
                selState = SelectedState.None;
            }
            switch (selState) {
                case SelectedState.Cursor:
                    cl += ' wj-state-selected';
                    break;
                case SelectedState.Selected:
                    cl += ' wj-state-multi-selected';
                    break;
            }

            // frozen area boundary
            if (r2 == rows.frozen - 1) {
                cl += ' wj-frozen-row';
            }
            if (c2 == cols.frozen - 1) {
                cl += ' wj-frozen-col';
            }

            // word-wrapping
            if (col.wordWrap || row.wordWrap) {
                cl += ' wj-wrap';
            }

            // optimization: skip cell update if possible
            if (canSkip && cl == cell.className) {
                var s = cell.style;
                if (s.top == css.top && s.width == css.width && s.height == css.height) {
                    if ((g._rtl && s.right == css.right) || (!g._rtl && s.left == css.left)) {
                        if (s.display) {
                            cell.style.display = '';
                            //s.display == ''; // this doesn't work... weird
                        }
                        return;
                    }
                }
            }

            // alignment
            css.textAlign = bcol.getAlignment();
            // TODO: vertical alignment?

            // group row indentation
            if (ct == CellType.Cell) {
                if (g.rows.maxGroupLevel > -1) {
                    css.paddingLeft = css.paddingRight = '';
                    if (c == g.columns.firstVisibleIndex && g.treeIndent) {
                        var level = gr ? Math.max(0, gr.level) : (g.rows.maxGroupLevel + 1),
                            indent = g.treeIndent * level + g._cellPadding;
                        if (g._rtl) {
                            css.paddingRight = indent;
                        } else {
                            css.paddingLeft = indent;
                        }
                    }
                }
            }

            // cell content
            if (updateContent != false) {
                var data = p.getCellData(r, c, false),
                    content = p.getCellData(r, c, true);
                if (ct == CellType.Cell && c == g.columns.firstVisibleIndex &&
                    gr && gr.hasChildren && !this._isEditingCell(g, r, c)) {

                    // collapse/expand outline
                    if (!content) {
                        content = gr.getGroupHeader();
                    }
                    cell.innerHTML = this._getTreeIcon(gr) + ' ' + content;
                    css.textAlign = '';

                } else if (ct == CellType.ColumnHeader && bcol.currentSort && g.showSort && (r2 == g._getSortRowIndex() || bcol != col)) {

                    // add sort class names to allow easier customization
                    cl += ' wj-sort-' + (bcol.currentSort == '+' ? 'asc' : 'desc');

                    // column header with sort sign
                    cell.innerHTML = escapeHtml(content) + '&nbsp;' + this._getSortIcon(bcol);

                } else if (ct == CellType.RowHeader && c == g.rowHeaders.columns.length - 1 && !content) {

                    // edit/new item template indicators
                    var ecv = <collections.IEditableCollectionView>g.collectionView,
                        editItem = ecv ? ecv.currentEditItem : null;
                    if (editItem && row.dataItem == editItem) {
                        content = '\u270E'; // pencil icon indicates item being edited
                    } else if (tryCast(row, _NewRowTemplate)) {
                        content = '*'; // asterisk indicates new row template
                    }
                    setText(cell, content);

                } else if (ct == CellType.Cell && bcol.dataType == DataType.Boolean && (!gr || isBoolean(data))) { // TFS 122709

                    // re-use/create checkbox
                    // (re-using allows selecting and checking/unchecking with a single click)
                    var chk = <HTMLInputElement>cell.firstChild;
                    if (!(chk instanceof HTMLInputElement) || chk.type != 'checkbox') {
                        cell.innerHTML = '<input type="checkbox"/>';
                        chk = <HTMLInputElement>cell.firstChild;
                    }

                    // initialize/update checkbox value
                    chk.checked = data == true ? true : false;
                    chk.indeterminate = data == null;

                    // disable checkbox if it is not editable (so user can't click it)
                    chk.disabled = !g._edtHdl._allowEditing(r, c);
                    if (chk.disabled) {
                        chk.style.cursor = 'default';
                    }

                    // assign editor to grid
                    if (g.editRange && g.editRange.contains(r, c)) {
                        g._edtHdl._edt = chk;
                    }

                } else if (ct == CellType.Cell && this._isEditingCell(g, r, c)) {

                    // select input type (important for mobile devices)
                    var inpType = bcol.inputType;
                    if (!bcol.inputType) {
                        inpType = bcol.dataType == DataType.Number && !bcol.dataMap ? 'tel' : 'text';
                    }

                    // get editor value (use full precision when editing floating point values)
                    // this is a little tricky: TFS 123276, 134218, 135336
                    if (!bcol.dataMap && !bcol.mask) {
                        var val = <number>p.getCellData(r, c, false);
                        if (isNumber(val)) {
                            var fmt = bcol.format;
                            if (fmt && val != Math.round(val)) {
                                fmt = bcol.format.replace(/([a-z])(\d*)(.*)/ig, '$0112$3');
                            }
                            content = Globalize.formatNumber(val, fmt, true);
                        }
                    }

                    // create/initialize editor
                    cell.innerHTML = '<input type="' + inpType + '" class="wj-grid-editor wj-form-control">';
                    var edt = <HTMLInputElement>cell.children[0];
                    edt.value = content;
                    edt.style.textAlign = bcol.getAlignment(); // right-align numbers when editing
                    css.padding = '0px'; // no padding on cell div (the editor has it)

                    // apply mask, if any
                    if (bcol.mask) {
                        var mp = new _MaskProvider(edt, bcol.mask);
                    }

                    // assign editor to grid
                    g._edtHdl._edt = edt;

                } else {

                    // regular content
                    if (ct == CellType.Cell && (row.isContentHtml || bcol.isContentHtml)) {
                        cell.innerHTML = content;
                    } else {
                        setText(cell, content);
                    }
                }

                // add drop-down element to the cell if the column:
                // a) has a dataMap,
                // b) has showDropDown set to not false (null or true)
                // c) is editable
                if (ct == CellType.Cell && wijmo.input &&
                    bcol.dataMap && bcol.showDropDown !== false && g._edtHdl._allowEditing(r, c)) {

                    // create icon once
                    if (!CellFactory._ddIcon) {
                        CellFactory._ddIcon = createElement(
                            '<div class="' + CellFactory._WJC_DROPDOWN + '"><span class="wj-glyph-down"></span></div>'
                        );
                    }

                    // clone icon and add clone to cell
                    var dd = <HTMLElement>CellFactory._ddIcon.cloneNode(true);
                    cell.appendChild(dd);
                }
            }

            // make row/col headers draggable
            switch (ct) {
                case CellType.RowHeader:
                    cell.removeAttribute('draggable');
                    if (!gr && !nr && (g.allowDragging & AllowDragging.Rows) != 0) {
                        cell.setAttribute('draggable', 'true');
                    }
                    break;
                case CellType.ColumnHeader:
                    cell.removeAttribute('draggable');
                    if ((g.allowDragging & AllowDragging.Columns) != 0) {
                        cell.setAttribute('draggable', 'true');
                    }
                    break;
            }

            // apply class specifier to cell
            if (cell.className != cl) {
                cell.className = cl;
            }

            // apply style to cell
            setCss(cell, css);

            // customize the cell
            if (g.itemFormatter) {
                g.itemFormatter(p, r, c, cell);
            }
            if (g.formatItem.hasHandlers) {
                var rng = CellFactory._fmtRng;
                if (!rng) { // avoid allocating a new CellRange each time (this may get called a lot!)
                    rng = CellFactory._fmtRng = new CellRange(r, c, r2, c2);
                } else {
                    rng.setRange(r, c, r2, c2);
                }
                var e = new FormatItemEventArgs(p, rng, cell);
                g.onFormatItem(e);
            }
        }
        /**
         * Disposes of a cell element and releases all resources associated with it.
         *
         * @param cell The element that represents the cell.
         */
        public disposeCell(cell: HTMLElement) {
            // no action needed for standard cells...
        }

        // ** implementation

        // determines whether the grid is currently editing a cell
        private _isEditingCell(g: FlexGrid, r: number, c: number): boolean {
            return g.editRange && g.editRange.contains(r, c);
        }

        // get an element to create a collapse/expand icon
        // NOTE: the _WJC_COLLAPSE class is used by the mouse handler to identify
        // the collapse/expand button/element.
        private _getTreeIcon(gr: GroupRow): string {
            var glyph = 'wj-glyph-' +
                (gr.isCollapsed ? '' : 'down-') +
                (gr.grid._rtl ? 'left' : 'right');
            return '<span class="' + CellFactory._WJC_COLLAPSE + ' ' + glyph + '"></span>';
        }

        // get an element to create a sort up/down icon
        private _getSortIcon(col: Column): string {
            return '<span class="wj-glyph-' + (col.currentSort == '+' ? 'up' : 'down') + '"></span>';
        }
    }
}

module wijmo.grid {
    'use strict';

    /**
     * Represents a rectangular group of cells defined by two row indices and
     * two column indices.
     */
    export class CellRange {
        _row: number;
        _col: number;
        _row2: number;
        _col2: number;

        /**
         * Initializes a new instance of the @see:CellRange class.
         *
         * @param r The index of the first row in the range (defaults to -1).
         * @param c The index of the first column in the range (defaults to -1).
         * @param r2 The index of the last row in the range (defaults to <b>r</b>).
         * @param c2 The index of the last column in the range (defaults to <b>c</b>).
         */
        constructor(r = -1, c = -1, r2 = r, c2 = c) {
            this.setRange(r, c, r2, c2);
        }
        /**
         * Initializes an existing @see:CellRange.
         *
         * @param r The index of the first row in the range (defaults to -1).
         * @param c The index of the first column in the range (defaults to -1).
         * @param r2 The index of the last row in the range (defaults to <b>r</b>).
         * @param c2 The index of the last column in the range (defaults to <b>c</b>).
         */
        setRange(r = -1, c = -1, r2 = r, c2 = c) {
            this._row = asInt(r);
            this._col = asInt(c);
            this._row2 = asInt(r2);
            this._col2 = asInt(c2);
        }
        /**
         * Gets or sets the index of the first row in the range.
         */
        get row(): number {
            return this._row;
        }
        set row(value: number) {
            this._row = asInt(value);
        }
        /**
         * Gets or sets the index of the first column in the range.
         */
        get col(): number {
            return this._col;
        }
        set col(value: number) {
            this._col = asInt(value);
        }
        /**
         * Gets or sets the index of the second row in the range.
         */
        get row2(): number {
            return this._row2;
        }
        set row2(value: number) {
            this._row2 = asInt(value);
        }
        /**
         * Gets or sets the index of the second column in the range.
         */
        get col2(): number {
            return this._col2;
        }
        set col2(value: number) {
            this._col2 = asInt(value);
        }
        /**
         * Creates a copy of the range.
         */
        clone(): CellRange {
            return new CellRange(this._row, this._col, this._row2, this._col2);
        }
        /**
         * Gets the number of rows in the range.
         */
        get rowSpan(): number {
            return Math.abs(this._row2 - this._row) + 1;
        }
        /**
         * Gets the number of columns in the range.
         */
        get columnSpan(): number {
            return Math.abs(this._col2 - this._col) + 1;
        }
        /**
         * Gets the index of the top row in the range.
         */
        get topRow(): number {
            return Math.min(this._row, this._row2);
        }
        /**
         * Gets the index of the bottom row in the range.
         */
        get bottomRow(): number {
            return Math.max(this._row, this._row2);
        }
        /**
         * Gets the index of the leftmost column in the range.
         */
        get leftCol(): number {
            return Math.min(this._col, this._col2);
        }
        /**
         * Gets the index of the rightmost column in the range.
         */
        get rightCol(): number {
            return Math.max(this._col, this._col2);
        }
        /**
         * Checks whether the range contains valid row and column indices 
         * (row and column values are zero or greater).
         */
        get isValid(): boolean {
            return this._row > -1 && this._col > -1 && this._row2 > -1 && this._col2 > -1;
        }
        /**
         * Checks whether this range corresponds to a single cell (beginning and ending rows have 
         * the same index, and beginning and ending columns have the same index).
         */
        get isSingleCell(): boolean {
            return this._row == this._row2 && this._col == this._col2;
        }
        /**
         * Checks whether the range contains another range or a specific cell.
         *
         * @param r The CellRange object or row index to find.
         * @param c The column index (required if the r parameter is not a CellRange object).
         */
        contains(r: any, c?: number): boolean {

            // first parameter may be a cell range
            var rng = <CellRange>tryCast(r, CellRange);
            if (rng) {
                return rng.topRow >= this.topRow && rng.bottomRow <= this.bottomRow &&
                    rng.leftCol >= this.leftCol && rng.rightCol <= this.rightCol;
            }

            // check specific cell
            if (isInt(r) && isInt(c)) {
                return r >= this.topRow && r <= this.bottomRow &&
                       c >= this.leftCol && c <= this.rightCol;
            }

            // anything else is an error
            throw 'contains expects a CellRange or row/column indices.';
        }
        /**
         * Checks whether the range contains a given row.
         *
         * @param r The index of the row to find.
         */
        containsRow(r: number): boolean {
            return asInt(r) >= this.topRow && r <= this.bottomRow;
        }
        /**
         * Checks whether the range contains a given column.
         *
         * @param c The index of the column to find.
         */
        containsColumn(c: number): boolean {
            return asInt(c) >= this.leftCol && c <= this.rightCol;
        }
        /**
         * Checks whether the range intersects another range.
         *
         * @param rng The CellRange object to check.
         */
        intersects(rng: CellRange): boolean {
            return this.intersectsRow(rng) && this.intersectsColumn(rng);
        }
        /**
         * Checks whether the range intersects the rows in another range.
         *
         * @param rng The CellRange object to check.
         */
        intersectsRow(rng: CellRange): boolean {
            return rng && !(this.bottomRow < rng.topRow || this.topRow > rng.bottomRow);
        }
        /**
         * Checks whether the range intersects the columns in another range.
         *
         * @param rng The CellRange object to check.
         */
        intersectsColumn(rng: CellRange): boolean {
            return rng && !(this.rightCol < rng.leftCol || this.leftCol > rng.rightCol);
        }
        /**
         * Gets the rendered size of this range.
         *
         * @param p The @see:GridPanel object that contains the range.
         * @return A @see:Size object that represents the sum of row heights and column widths in the range.
         */
        getRenderSize(p: GridPanel): Size {
            var sz = new Size(0, 0);
            if (this.isValid) {
                for (var r = this.topRow; r <= this.bottomRow; r++) {
                    sz.height += p.rows[r].renderSize;
                }
                for (var c = this.leftCol; c <= this.rightCol; c++) {
                    sz.width += p.columns[c].renderSize;
                }
            }
            return sz;
        }
        /**
         * Checks whether the range equals another range.
         *
         * @param rng The CellRange object to compare to this range.
         */
        equals(rng: CellRange): boolean {
            return (rng instanceof CellRange) &&
                this._row == rng._row && this._col == rng._col &&
                this._row2 == rng._row2 && this._col2 == rng._col2;
        }
    }
}
module wijmo.grid {
    'use strict';

    /**
     * Specifies flags that represent the state of a grid row or column.
     */
    export enum RowColFlags {
        /** The row or column is visible. */
        Visible = 1,
        /** The row or column can be resized. */
        AllowResizing = 2,
        /** The row or column can be dragged to a new position with the mouse. */
        AllowDragging = 4,
        /** The row or column can contain merged cells. */
        AllowMerging = 8,
        /** The column can be sorted by clicking its header with the mouse. */
        AllowSorting = 16,
        /** The column was generated automatically. */
        AutoGenerated = 32,
        /** The group row is collapsed. */
        Collapsed = 64,
        /** The row has a parent group that is collapsed. */
        ParentCollapsed = 128,
        /** The row or column is selected. */
        Selected = 256,
        /** The row or column is read-only (cannot be edited). */
        ReadOnly = 512,
        /** Cells in this row or column contain HTML text. */
        HtmlContent = 1024,
        /** Cells in this row or column may contain wrapped text. */
        WordWrap = 2048,
        /** Default settings for new rows. */
        RowDefault = Visible | AllowResizing,
        /** Default settings for new columns. */
        ColumnDefault = Visible | AllowDragging | AllowResizing | AllowSorting
    }

    /**
     * An abstract class that serves as a base for the @see:Row and @see:Column classes.
     */
    export class RowCol {
        _sz: number; // null or < 0 means use default
        _cssClass: string;
        _szMin: number;
        _szMax: number;
        _list = null;
        _f: RowColFlags;
        _pos = 0;
        _idx = -1;

        /**
         * Gets or sets a value that indicates whether the row or column is visible.
         */
        get visible(): boolean {
            return this._getFlag(RowColFlags.Visible);
        }
        set visible(value: boolean) {
            this._setFlag(RowColFlags.Visible, value);
        }
        /**
         * Gets a value that indicates whether the row or column is visible and not collapsed.
         *
         * This property is read-only. To change the visibility of a
         * row or column, use the @see:visible property instead.
         */
        get isVisible(): boolean {

            // if visible is false, we're not visible
            if (!this._getFlag(RowColFlags.Visible)) {
                return false;
            }

            // if the parent node is collapsed and this is not a new row, we're not visible
            if (this._getFlag(RowColFlags.ParentCollapsed) && !(this instanceof _NewRowTemplate)) {
                return false;
            }

            // looks like we're visible
            return true;
        }
        /**
         * Gets the position of the row or column.
         */
        get pos(): number {
            if (this._list) this._list._update();
            return this._pos;
        }
        /**
         * Gets the index of the row or column in the parent collection.
         */
        get index(): number {
            if (this._list) this._list._update();
            return this._idx;
        }
        /**
         * Gets or sets the size of the row or column.
         * Setting this property to null or negative values causes the element to use the 
         * parent collection's default size.
         */
        get size(): number {
            return this._sz;
        }
        set size(value: number) {
            if (value != this._sz) {
                this._sz = asNumber(value, true);
                this.onPropertyChanged();
            }
        }
        /**
         * Gets the render size of the row or column.
         * This property accounts for visibility, default size, and min and max sizes.
         */
        get renderSize(): number {
            if (!this.isVisible) {
                return 0;
            }
            var sz = this._sz,
                list = this._list;

            // default size
            if ((sz == null || sz < 0) && list != null) {
                return Math.round((<RowColCollection>(list)).defaultSize);
            }
            
            // min/max
            if (list != null) {
                if (list.minSize != null && sz < list.minSize) {
                    sz = list.minSize;
                }
                if (list.maxSize != null && sz > list.maxSize) {
                    sz = list.maxSize;
                }
            }
            if (this._szMin != null && sz < this._szMin) {
                sz = this._szMin;
            }
            if (this._szMax != null && sz > this._szMax) {
                sz = this._szMax;
            }

            // done
            return Math.round(sz);
        }
        /**
         * Gets or sets a value that indicates whether the user can resize the row or column with the mouse.
         */
        get allowResizing(): boolean {
            return this._getFlag(RowColFlags.AllowResizing);
        }
        set allowResizing(value: boolean) {
            this._setFlag(RowColFlags.AllowResizing, value);
        }
        /**
         * Gets or sets a value that indicates whether the user can move the row or column to a new position with the mouse.
         */
        get allowDragging(): boolean {
            return this._getFlag(RowColFlags.AllowDragging);
        }
        set allowDragging(value: boolean) {
            this._setFlag(RowColFlags.AllowDragging, value);
        }
        /**
         * Gets or sets a value that indicates whether cells in the row or column can be merged.
         */
        get allowMerging(): boolean {
            return this._getFlag(RowColFlags.AllowMerging);
        }
        set allowMerging(value: boolean) {
            this._setFlag(RowColFlags.AllowMerging, value);
        }
        /**
         * Gets or sets a value that indicates whether the row or column is selected.
         */
        get isSelected(): boolean {
            return this._getFlag(RowColFlags.Selected);
        }
        set isSelected(value: boolean) {
            //this._setFlag(RowColFlags.Selected, value);

            // set flag quietly, then update selection (faster than full invalidation)
            if (this._setFlag(RowColFlags.Selected, value, true)) {
                var g = this.grid;
                if (g) {
                    g.refreshCells(false, true, true);
                }
            }
        }
        /**
         * Gets or sets a value that indicates whether cells in the row or column can be edited.
         */
        get isReadOnly(): boolean {
            return this._getFlag(RowColFlags.ReadOnly);
        }
        set isReadOnly(value: boolean) {
            this._setFlag(RowColFlags.ReadOnly, value);
        }
        /**
         * Gets or sets a value that indicates whether cells in this row or column 
         * contain HTML content rather than plain text.
         */
        get isContentHtml(): boolean {
            return this._getFlag(RowColFlags.HtmlContent);
        }
        set isContentHtml(value: boolean) {
            if (this.isContentHtml != value) {
                this._setFlag(RowColFlags.HtmlContent, value);
                if (this.grid) {
                    this.grid.invalidate();
                }
            }
        }
        /**
         * Gets or sets a value that indicates whether cells in the row or column wrap their content.
         */
        get wordWrap(): boolean {
            return this._getFlag(RowColFlags.WordWrap);
        }
        set wordWrap(value: boolean) {
            this._setFlag(RowColFlags.WordWrap, value);
        }
        /**
         * Gets or sets a CSS class name to use when rendering 
         * non-header cells in the row or column.
         */
        get cssClass(): string {
            return this._cssClass;
        }
        set cssClass(value: string) {
            if (value != this._cssClass) {
                this._cssClass = asString(value);
                if (this.grid) {
                    this.grid.invalidate(false);
                }
            }
        }
        /**
         * Gets the @see:FlexGrid that owns the row or column.
         */
        get grid(): FlexGrid {
            return this._list? (<RowColCollection>this._list)._g: null;
        }
        /**
         * Gets the @see:ICollectionView bound to this row or column.
         */
        get collectionView(): collections.ICollectionView {
            return this.grid ? this.grid.collectionView : null;
        }
        /**
         * Marks the owner list as dirty and refreshes the owner grid.
         */
        onPropertyChanged() {
            if (this._list) {
                this._list._dirty = true;
                this.grid.invalidate();
            }
        }

        // Gets the value of a flag.
        _getFlag(flag: RowColFlags): boolean {
            return (this._f & flag) != 0;
        }

        // Sets the value of a flag, with optional notification.
        _setFlag(flag: RowColFlags, value: boolean, quiet?: boolean): boolean {
            if (value != this._getFlag(flag)) {
                this._f = value ? (this._f | flag) : (this._f & ~flag);
                if (!quiet) {
                    this.onPropertyChanged();
                }
                return true;
            }
            return false;
        }
    }

    /**
     * Represents a column on the grid.
     */
    export class Column extends RowCol {
        private static _ctr = 0;
        private _hdr: string;
        private _name: string;
        private _type: DataType;
        private _align: string;
        private _map: DataMap;
        private _fmt: string;
        private _agg: Aggregate;
        private _inpType: string;
        private _mask: string;
        private _required: boolean;
        private _showDropDown: boolean;
        private _ddCssClass: string;

        /*private*/ _binding: Binding;
        /*private*/ _bindingSort: Binding;
        /*private*/ _szStar: string;
        /*private*/ _hash: string; // unique column id

        /**
         * Initializes a new instance of the @see:Column class.
         *
         * @param options Initialization options for the column.
         */
        constructor(options? : any) {
            super();
            this._f = RowColFlags.ColumnDefault;
            this._hash = Column._ctr.toString(36); // unique column key (used for unbound rows)
            Column._ctr++;
            if (options) {
                copy(this, options);
            }
        }
        /**
         * Gets or sets the name of the column.
         *
         * The column name can be used to retrieve the column using the
         * @see:FlexGrid.getColumn method.
         */
        get name(): string {
            return this._name;
        }
        set name(value: string) {
            this._name = value;
        }
        /**
         * Gets or sets the type of value stored in the column.
         *
         * Values are coerced into the proper type when editing the grid.
         */
        get dataType(): DataType {
            return this._type;
        }
        set dataType(value: DataType) {
            if (this._type != value) {
                this._type = asEnum(value, DataType);
                if (this.grid) {
                    this.grid.invalidate();
                }
            }
        }
        /**
         * Gets or sets a value that determines whether values in the column 
         * are required.
         *
         * By default, this property is set to null, which means values 
         * are required, but string columns may contain empty strings.
         *
         * When set to true, values are required and empty strings are 
         * not allowed.
         *
         * When set to false, null values and empty strings are allowed.
         */
        get isRequired(): boolean {
            return this._required;
        }
        set isRequired(value: boolean) {
            this._required = asBoolean(value, true);
        }
        // Deprecated: use 'isRequired' instead to avoid confusion with 'required' HTML attribute.
        get required(): boolean {
            _deprecated('required', 'isRequired');
            return this.isRequired;
        }
        set required(value: boolean) {
            _deprecated('required', 'isRequired');
            this.isRequired = value;
        }
        /**
         * Gets or sets a value that indicates whether the grid adds drop-down buttons to the
         * cells in this column.
         *
         * The drop-down buttons are shown only if the column has a @see:dataMap
         * set and is editable. Clicking on the drop-down buttons causes the grid
         * to show a list where users can select the value for the cell.
         *
         * Cell drop-downs require the wijmo.input module to be loaded.
         */
        get showDropDown(): boolean {
            return this._showDropDown;
        }
        set showDropDown(value: boolean) {
            if (value != this._showDropDown) {
                this._showDropDown = asBoolean(value, true);
                if (this.grid) {
                    this.grid.invalidate();
                }
            }
        }
        /**
         * Gets or sets a CSS class name to add to drop-downs in this column.
         *
         * The drop-down buttons are shown only if the column has a @see:dataMap
         * set and is editable. Clicking on the drop-down buttons causes the grid
         * to show a list where users can select the value for the cell.
         *
         * Cell drop-downs require the wijmo.input module to be loaded.
         */
        get dropDownCssClass(): string {
            return this._ddCssClass;
        }
        set dropDownCssClass(value: string) {
            this._ddCssClass = asString(value);
        }
        /**
         * Gets or sets the "type" attribute of the HTML input element used to edit values
         * in this column.
         *
         * By default, this property is set to "tel" for numeric columns, and to "text" for
         * all other non-boolean column types. The "tel" input type causes mobile devices 
         * to show a numeric keyboard that includes a negative sign and a decimal separator.
         *
         * Use this property to change the default setting if the default does not work well
         * for the current culture, device, or application. In these cases, try setting the 
         * property to "number" or simply "text."
         */
        get inputType(): string {
            return this._inpType;
        }
        set inputType(value: string) {
            this._inpType = asString(value, true);
        }
        /**
         * Gets or sets a mask to use while editing values in this column.
         *
         * The mask format is the same used by the @see:wijmo.input.InputMask
         * control.
         *
         * If specified, the mask must be compatible with the value of
         * the @see:format property. For example, the mask '99/99/9999' can 
         * be used for entering dates formatted as 'MM/dd/yyyy'.
         */
        get mask(): string {
            return this._mask;
        }
        set mask(value: string) {
            this._mask = asString(value, true);
        }
        /**
         * Gets or sets the name of the property the column is bound to.
         */
        get binding(): string {
            return this._binding ? this._binding.path : null;
        }
        set binding(value: string) {
            if (value != this.binding) {
                var path = asString(value);
                this._binding = path ? new Binding(path) : null;
                if (!this._type && this.grid && this._binding) {
                    var cv = this.grid.collectionView;
                    if (cv && cv.sourceCollection && cv.sourceCollection.length) {
                        var item = cv.sourceCollection[0];
                        this._type = getType(this._binding.getValue(item));
                    }
                }
                this.onPropertyChanged();
            }
        }
        /**
         * Gets or sets the name of the property to use when sorting this column.
         *
         * Use this property in cases where you want the sorting to be performed
         * based on values other than the ones specified by the @see:binding property.
         *
         * Setting this property is null causes the grid to use the value of the
         * @see:binding property to sort the column.
         */
        get sortMemberPath(): string {
            return this._bindingSort ? this._bindingSort.path : null;
        }
        set sortMemberPath(value: string) {
            if (value != this.sortMemberPath) {
                var path = asString(value);
                this._bindingSort = path ? new Binding(path) : null;
                this.onPropertyChanged();
            }
        }
        /**
         * Gets or sets the width of the column.
         *
         * Column widths may be positive numbers (sets the column width in pixels), 
         * null or negative numbers (uses the collection's default column width), or
         * strings in the format '{number}*' (star sizing).
         *
         * The star-sizing option performs a XAML-style dynamic sizing where column 
         * widths are proportional to the number before the star. For example, if
         * a grid has three columns with widths "100", "*", and "3*", the first column
         * will be 100 pixels wide, the second will take up 1/4th of the remaining
         * space, and the last will take up the remaining 3/4ths of the remaining space.
         *
         * Star-sizing allows you to define columns that automatically stretch to fill
         * the width available. For example, set the width of the last column to "*"
         * and it will automatically extend to fill the entire grid width so there's
         * no empty space. You may also want to set the column's @see:minWidth property
         * to prevent the column from getting too narrow.
         */
        get width() : any {
            if (this._szStar != null) {
                return this._szStar;
            } else {
                return this.size;
            }
        }
        set width(value: any) {
            if (Column._parseStarSize(value) != null) {
                this._szStar = value;
                this.onPropertyChanged();
            } else {
                this._szStar = null;
                this.size = asNumber(value, true);
            }
        }
        /**
         * Gets or sets the minimum width of the column.
         */
        get minWidth(): number {
            return this._szMin;
        }
        set minWidth(value: number) {
            if (value != this._szMin) {
                this._szMin = asNumber(value, true, true);
                this.onPropertyChanged();
            }
        }
        /**
         * Gets or sets the maximum width of the column.
         */
        get maxWidth(): number {
            return this._szMax;
        }
        set maxWidth(value: number) {
            if (value != this._szMax) {
                this._szMax = asNumber(value, true, true);
                this.onPropertyChanged();
            }
        }
        /**
         * Gets the render width of the column.
         *
         * The value returned takes into account the column's visibility, default size, and min and max sizes.
         */
        get renderWidth(): number {
            return this.renderSize;
        }
        /**
         * Gets or sets the horizontal alignment of items in the column.
         *
         * The default value for this property is null, which causes the grid to select
         * the alignment automatically based on the column's @see:dataType (numbers are
         * right-aligned, Boolean values are centered, and other types are left-aligned).
         *
         * If you want to override the default alignment, set this property
         * to 'left,' 'right,' or 'center,'
         */
        get align(): string {
            return this._align;
        }
        set align(value: string) {
            if (this._align != value) {
                this._align = value;
                this.onPropertyChanged();
            }
        }
        /**
         * Gets the actual column alignment. 
         *
         * Returns the value of the @see:align property if it is not null, or
         * selects the alignment based on the column's @see:dataType.
         */
        getAlignment(): string {
            var value = this._align;
            if (value == null) {
                value = '';
                if (!this._map) {
                    switch (this._type) {
                        case DataType.Boolean:
                            value = 'center';
                            break;
                        case DataType.Number:
                            value = 'right';
                            break;
                    }
                }
            }
            return value;
        }
        /**
         * Gets or sets the text displayed in the column header.
         */
        get header(): string {
            return this._hdr ? this._hdr : this.binding;
        }
        set header(value: string) {
            if (this._hdr != value) {
                this._hdr = value;
                this.onPropertyChanged();
            }
        }
        /**
         * Gets or sets the @see:DataMap used to convert raw values into display
         * values for the column.
         *
         * Columns with an associated @see:dataMap show drop-down buttons that
         * can be used for quick editing. If you do not want to show the drop-down 
         * buttons, set the column's @see:showDropDown property to false.
         *
         * Cell drop-downs require the wijmo.input module to be loaded.
         */
        get dataMap(): DataMap {
            return this._map;
        }
        set dataMap(value: DataMap) {
            if (this._map != value) {

                // disconnect old map
                if (this._map) {
                    this._map.mapChanged.removeHandler(this.onPropertyChanged, this);
                }

                // convert arrays into DataMaps
                if (isArray(value)) {
                    value = new DataMap(value, null, null);
                }

                // set new map
                this._map = asType(value, DataMap, true);

                // connect new map
                if (this._map) {
                    this._map.mapChanged.addHandler(this.onPropertyChanged, this);
                }
                this.onPropertyChanged();
            }
        }
        /**
         * Gets or sets the format string used to convert raw values into display 
         * values for the column (see @see:Globalize).
         */
        get format(): string {
            return this._fmt;
        }
        set format(value: string) {
            if (this._fmt != value) {
                this._fmt = value;
                this.onPropertyChanged();
            }
        }
        /**
         * Gets or sets a value that indicates whether the user can sort the column by clicking its header.
         */
        get allowSorting(): boolean {
            return this._getFlag(RowColFlags.AllowSorting);
        }
        set allowSorting(value: boolean) {
            this._setFlag(RowColFlags.AllowSorting, value);
        }
        /**
         * Gets a string that describes the current sorting applied to the column.
         * Possible values are '+' for ascending order, '-' for descending order, or 
         * null for unsorted columns.
         */
        get currentSort(): string {
            if (this.grid && this.grid.collectionView && this.grid.collectionView.canSort) {
                var sds = this.grid.collectionView.sortDescriptions;
                for (var i = 0; i < sds.length; i++) {
                    if (sds[i].property == this._getBindingSort()) {
                        return sds[i].ascending ? '+' : '-';
                    }
                }
            }
            return null;
        }
        /**
         * Gets or sets the @see:Aggregate to display in the group header rows 
         * for the column.
         */
        get aggregate(): Aggregate {
            return this._agg != null ? this._agg : Aggregate.None;
        }
        set aggregate(value: Aggregate) {
            if (value != this._agg) {
                this._agg = asEnum(value, Aggregate);
                this.onPropertyChanged();
            }
        }

        // gets the binding used for sorting (sortMemberPath if specified, binding ow)
        /*protected*/ _getBindingSort() : string {
            return this.sortMemberPath ? this.sortMemberPath :
                this.binding ? this.binding :
                null;
        }

        // parses a string in the format '<number>*' and returns the number (or null if the parsing fails).
        static _parseStarSize(value: any) {
            if (isString(value) && value.length > 0 && value[value.length - 1] == '*') {
                var sz = value.length == 1 ? 1 : value.substr(0, value.length - 1) * 1;
                if (sz > 0 && !isNaN(sz)) {
                    return sz;
                }
            }
            return null;
        }
    }

    /**
     * Represents a row in the grid.
     */
    export class Row extends RowCol {
        private _data: any;
        /*private*/ _ubv: any; // unbound value storage

        /**
         * Initializes a new instance of the @see:Row class.
         *
         * @param dataItem The data item that this row is bound to.
         */
        constructor(dataItem?: any) {
            super();
            this._f = RowColFlags.ColumnDefault;
            this._data = dataItem;
        }
        /**
         * Gets or sets the item in the data collection that the item is bound to.
         */
        get dataItem(): any {
            return this._data;
        }
        set dataItem(value: any) {
            this._data = value;
        }
        /**
         * Gets or sets the height of the row.
         * Setting this property to null or negative values causes the element to use the 
         * parent collection's default size.
         */
        get height(): number {
            return this.size;
        }
        set height(value: number) {
            this.size = value;
        }
        /**
         * Gets the render height of the row.
         *
         * The value returned takes into account the row's visibility, default size, and min and max sizes.
         */
        get renderHeight(): number {
            return this.renderSize;
        }
    }

    /**
     * Represents a row that serves as a header for a group of rows.
     */
    export class GroupRow extends Row {
        _level = -1;

        /**
         * Initializes a new instance of the @see:GroupRow class.
         */
        constructor() {
            super();
            this.isReadOnly = true; // group rows are read-only by default
        }
        /**
         * Gets or sets the hierarchical level of the group associated with the GroupRow.
         */
        get level(): number {
            return this._level;
        }
        set level(value: number) {
            asInt(value);
            if (value != this._level) {
                this._level = value;
                this.onPropertyChanged();
            }
        }
        /**
         * Gets a value that indicates whether the group row has child rows.
         */
        get hasChildren(): boolean {
            if (this.grid != null && this._list != null) {

                // get the next row
                this._list._update();
                var rNext = this.index < this._list.length - 1
                    ? this._list[this.index + 1]
                    : null;

                // check if it's a group row or a new row template
                var gr = tryCast(rNext, GroupRow),
                    nr = tryCast(rNext, _NewRowTemplate);

                // return true if there is a next row and it's a data row or a deeper group row
                return rNext && nr == null && (gr == null || gr.level > this.level);
            }
            return true;
        }
        /**
         * Gets or sets a value that indicates whether the GroupRow is collapsed 
         * (child rows are hidden) or expanded (child rows are visible).
         */
        get isCollapsed(): boolean {
            return this._getFlag(RowColFlags.Collapsed);
        }
        set isCollapsed(value: boolean) {
            asBoolean(value);
            if (value != this.isCollapsed && this._list != null) {
                this._setCollapsed(value);
            }
        }
        /**
         * Gets the header text for this @see:GroupRow.
         */
        getGroupHeader(): string {
            var
                grid = this.grid,
                fmt = grid.groupHeaderFormat ? grid.groupHeaderFormat : wijmo.culture.FlexGrid.groupHeaderFormat,
                group = <collections.CollectionViewGroup>tryCast(this.dataItem, collections.CollectionViewGroup);
            if (group && fmt) {

                // get group info
                var propName = group.groupDescription['propertyName'],
                    value = group.name,
                    col = grid.columns.getColumn(propName);

                // customize with column info if possible
                var isHtml = this.isContentHtml; // TFS 114902
                if (col) {
                    isHtml = isHtml || col.isContentHtml;
                    if (col.header) {
                        propName = col.header;
                    }
                    if (col.dataMap) {
                        value = col.dataMap.getDisplayValue(value);
                    } else if (col.format) {
                        value = Globalize.format(value, col.format);
                    }
                }

                // get count including all items (including items not on the current page,
                // as calculated when setting Column.Aggregate TFS 195467)
                var count = group.getAggregate(Aggregate.CntAll, null, grid.collectionView);
                //var count = group.items.length;

                // build header text
                return format(fmt, {
                    name: escapeHtml(propName),
                    value: isHtml ? value : escapeHtml(value),
                    level: group.level,
                    count: count
                });
            }
            return '';
        }

        // sets the collapsed/expanded state of a group row
        _setCollapsed(collapsed: boolean) {
            var g = this.grid,
                rows = g.rows,
                rng = this.getCellRange(),
                e = new CellRangeEventArgs(g.cells, new CellRange(this.index, -1)),
                gr: GroupRow;

            // fire GroupCollapsedChanging
            g.onGroupCollapsedChanging(e);

            // if user canceled, or edits failed, bail out
            if (e.cancel) { // && TODO: grid.FinishEditing()) {
                return;
            }

            // apply new value
            g.deferUpdate(() => {

                // collapse/expand this group
                this._setFlag(RowColFlags.Collapsed, collapsed);
                for (var r = rng.topRow + 1; r <= rng.bottomRow && r > -1 && r < rows.length; r++) {

                    // apply state to this row
                    rows[r]._setFlag(RowColFlags.ParentCollapsed, collapsed);

                    // if this is a group, skip range to preserve the original state
                    gr = tryCast(rows[r], GroupRow);
                    if (gr != null && gr.isCollapsed) {
                        r = gr.getCellRange().bottomRow;
                    }
                }
            });

            // fire GroupCollapsedChanged
            g.onGroupCollapsedChanged(e);
        }

        /**
         * Gets a CellRange object that contains all of the rows in the group represented 
         * by the GroupRow and all of the columns in the grid.
         */
        getCellRange(): CellRange {
            var rows = this._list,
                top = this.index,
                bottom = rows.length - 1;
            for (var r = top + 1; r <= bottom; r++) {
                var gr = tryCast(rows[r], GroupRow);
                if (gr != null && gr.level <= this.level) {
                    bottom = r - 1;
                    break;
                }
            }
            return new CellRange(top, 0, bottom, this.grid.columns.length - 1);
        }
    }

    /**
     * Abstract class that serves as a base for row and column collections.
     */
    export class RowColCollection extends collections.ObservableArray {
        _g: FlexGrid;
        _frozen = 0;
        _szDef = 28;
        _szTot = 0;
        _dirty = false;
        _szMin: number;
        _szMax: number;

        /**
         * Initializes a new instance of the @see:RowColCollection class.
         *
         * @param g The @see:FlexGrid that owns the collection.
         * @param defaultSize The default size of the elements in the collection.
         */
        constructor(g: FlexGrid, defaultSize: number) {
            super();
            this._g = asType(g, FlexGrid);
            this._szDef = asNumber(defaultSize, false, true);
        }
        /**
         * Gets or sets the default size of elements in the collection.
         */
        get defaultSize(): number {
            return this._szDef;
        }
        set defaultSize(value: number) {
            if (this._szDef != value) {
                this._szDef = asNumber(value, false, true);
                this._dirty = true;
                this._g.invalidate();
            }
        }
        /**
         * Gets or sets the number of frozen rows or columns in the collection.
         *
         * Frozen rows and columns do not scroll, and instead remain at the top or left of
         * the grid, next to the fixed cells. Unlike fixed cells, however, frozen
         * cells may be selected and edited like regular cells.
         */
        get frozen(): number {
            return this._frozen;
        }
        set frozen(value: number) {
            if (value != this._frozen) {
                this._frozen = asNumber(value, false, true);
                this._dirty = true;
                this._g.invalidate();
            }
        }
        /**
         * Checks whether a row or column is frozen.
         *
         * @param index The index of the row or column to check.
         */
        isFrozen(index: number): boolean {
            return index < this.frozen;
        }
        /**
         * Gets or sets the minimum size of elements in the collection.
         */
        get minSize(): number {
            return this._szMin;
        }
        set minSize(value: number) {
            if (value != this._szMin) {
                this._szMin = asNumber(value, true, true);
                this._dirty = true;
                this._g.invalidate();
            }
        }
        /**
         * Gets or sets the maximum size of elements in the collection.
         */
        get maxSize(): number {
            return this._szMax;
        }
        set maxSize(value: number) {
            if (value != this._szMax) {
                this._szMax = asNumber(value, true, true);
                this._dirty = true;
                this._g.invalidate();
            }
        }
        /**
         * Gets the total size of the elements in the collection.
         */
        getTotalSize(): number {
            this._update();
            return this._szTot;
        }
        /**
         * Gets the index of the element at a given physical position.
         * @param position Position of the item in the collection, in pixels.
         */
        getItemAt(position: number): number {

            // update if necessary
            this._update();

            // shortcut for common case
            if (position <= 0 && this.length > 0) {
                return 0;
            }

            // binary search
            // REVIEW: is this worth it? might be better to use a simpler method?
            var min = 0,
                max = this.length - 1,
                cur, item;
            while (min <= max) {
                cur = (min + max) >>> 1;
                item = <RowCol>this[cur];
                if (item._pos > position) {
                    max = cur - 1;
                } else if (item._pos + item.renderSize < position) {
                    min = cur + 1;
                }
                else {

                    // skip invisible elements
                    while (cur > 0 && !this[cur].visible) {
                        cur--;
                    }
                    while (cur < this.length - 1 && !this[cur].visible) {
                        cur++;
                    }

                    // done
                    return cur;
                }
            }

            // not found, return max
            return max;
        }
        /**
         * Finds the next visible cell for a selection change.
         * @param index Starting index for the search.
         * @param move Type of move (size and direction).
         * @param pageSize Size of a page (in case the move is a page up/down).
         */
        getNextCell(index: number, move: SelMove, pageSize: number) {
            var i, item;
            switch (move) {
                case SelMove.Next:
                    for (i = index + 1; i < this.length; i++) {
                        if (this[i].renderSize > 0) return i;
                    }
                    break;
                case SelMove.Prev:
                    for (i = index - 1; i >= 0; i--) {
                        if (this[i].renderSize > 0) return i;
                    }
                    break;
                case SelMove.End:
                    for (i = this.length - 1; i >= 0; i--) {
                        if (this[i].renderSize > 0) return i;
                    }
                    break;
                case SelMove.Home:
                    for (i = 0; i < this.length; i++) {
                        if (this[i].renderSize > 0) return i;
                    }
                    break;
                case SelMove.NextPage:
                    item = this.getItemAt(this[index].pos + pageSize);
                    return item < 0
                        ? this.getNextCell(index, SelMove.End, pageSize)
                        : item;
                case SelMove.PrevPage:
                    item = this.getItemAt(this[index].pos - pageSize);
                    return item < 0
                        ? this.getNextCell(index, SelMove.Home, pageSize)
                        : item;
            }
            return index;
        }
        /**
         * Checks whether an element can be moved from one position to another.
         *
         * @param src The index of the element to move.
         * @param dst The position to which to move the element, or specify -1 to append the element.
         * @return Returns true if the move is valid, false otherwise.
         */
        canMoveElement(src: number, dst: number): boolean {

            // no move?
            if (dst == src) {
                return false;
            }

            // invalid move?
            if (src < 0 || src >= this.length || dst >= this.length) {
                return false;
            }

            // illegal move?
            if (dst < 0) dst = this.length - 1;
            var start = Math.min(src, dst),
                end = Math.max(src, dst);
            for (var i = start; i <= end; i++) {
                if (!this[i].allowDragging) {
                    return false;
                }
            }

            // can't move anything past the new row template (TFS 109012)
            if (this[dst] instanceof _NewRowTemplate) {
                return false;
            }

            // all seems OK
            return true;
        }
        /**
         * Moves an element from one position to another.
         * @param src Index of the element to move.
         * @param dst Position where the element should be moved to (-1 to append).
         */
        moveElement(src: number, dst: number) {
            if (this.canMoveElement(src, dst)) {
                var e = this[src];
                this.removeAt(src);
                if (dst < 0) dst = this.length;
                this.insert(dst, e);
            }
        }
        /**
         * Keeps track of dirty state and invalidate grid on changes.
         */
        onCollectionChanged(e = collections.NotifyCollectionChangedEventArgs.reset) {
            this._dirty = true;
            this._g.invalidate();
            super.onCollectionChanged(e);
        }
        /**
         * Appends an item to the array.
         *
         * @param item Item to add to the array.
         * @return The new length of the array.
         */
        push(item: any): number {
            item._list = this;
            return super.push(item);
        }
        /**
         * Removes or adds items to the array.
         *
         * @param index Position where items are added or removed.
         * @param count Number of items to remove from the array.
         * @param item Item to add to the array.
         * @return An array containing the removed elements.
         */
        splice(index: number, count: number, item?: any): any[]{
            if (item) {
                item._list = this;
            }
            return super.splice(index, count, item);
        }
        /**
         * Suspends notifications until the next call to @see:endUpdate.
         */
        beginUpdate() {

            // make sure we're up-to-date before suspending the updates
            this._update();

            // OK, now it's OK to suspend things
            super.beginUpdate();
        }

        // updates the index, size and position of the elements in the array.
        _update(): boolean {

            // update only if we're dirty *and* if the collection is not in an update block.
            // this is important for performance, especially when expanding/collapsing nodes.
            if (this._dirty && !this.isUpdating) {
                this._dirty = false;
                var
                    pos = 0,
                    rc: RowCol;
                for (var i = 0; i < this.length; i++) {
                    rc = this[i];
                    rc._idx = i;
                    rc._list = this;
                    rc._pos = pos;
                    pos += rc.renderSize;
                }
                this._szTot = pos;
                return true;
            }
            return false;
        }
    }

    /**
     * Represents a collection of @see:Column objects in a @see:FlexGrid control.
     */
    export class ColumnCollection extends RowColCollection {
        _firstVisible = -1;

        /**
         * Gets a column by name or by binding.
         *
         * The method searches the column by name. If a column with the given name 
         * is not found, it searches by binding. The searches are case-sensitive.
         *
         * @param name The name or binding to find.
         * @return The column with the specified name or binding, or null if not found.
         */
        getColumn(name: string): Column {
            var index = this.indexOf(name);
            return index > -1 ? this[index] : null;
        }
        /**
         * Gets the index of a column by name or binding.
         *
         * The method searches the column by name. If a column with the given name 
         * is not found, it searches by binding. The searches are case-sensitive.
         *
         * @param name The name or binding to find.
         * @return The index of column with the specified name or binding, or -1 if not found.
         */
        indexOf(name: any): number {

            // direct lookup
            if (name instanceof Column) {
                return super.indexOf(name);
            }

            // by name
            for (var i = 0; i < this.length; i++) {
                if ((<Column>this[i]).name == name) {
                    return i;
                }
            }

            // by binding
            for (var i = 0; i < this.length; i++) {
                if ((<Column>this[i]).binding == name) {
                    return i;
                }
            }
            return -1;
        }
        /**
         * Gets the index of the first visible column (where the outline tree is displayed).
         */
        get firstVisibleIndex() {
            this._update();
            return this._firstVisible;
        }

        // override to keep track of first visible column (and later to handle star sizes)
        _update(): boolean {
            if (super._update()) {
                this._firstVisible = -1;
                for (var i = 0; i < this.length; i++) {
                    if (<Column>(this[i]).visible) {
                        this._firstVisible = i;
                        break;
                    }
                }
                return true;
            }
            return false;
        }

        // update the width of the columns with star sizes
        _updateStarSizes(szAvailable: number): boolean {
            var starCount = 0,
                lastStarCol: Column;

            // count stars, remove fixed size columns from available size
            for (var i = 0; i < this.length; i++) {
                var col = this[i];
                if (col.isVisible) {
                    if (col._szStar) {
                        starCount += Column._parseStarSize(col._szStar);
                        lastStarCol = col;
                    } else {
                        szAvailable -= col.renderWidth;
                    }
                }
            }

            // update width of star columns
            if (lastStarCol) {
                var lastWidth = szAvailable;
                for (var i = 0; i < this.length; i++) {
                    var col = this[i];
                    if (col.isVisible) {
                        if (col._szStar) {
                            if (col == lastStarCol && lastWidth > 0) { // TFS 142608
                                col._sz = lastWidth; // to avoid round-off errors...
                            } else {
                                col._sz = Math.max(0, Math.round(Column._parseStarSize(col._szStar) / starCount * szAvailable));
                                lastWidth -= col.renderWidth;
                            }
                        }
                    }
                }
                this._dirty = true;
                this._update();
                return true;
            }

            // no star sizes...
            return false;
        }
    }

    /**
     * Represents a collection of @see:Row objects in a @see:FlexGrid control.
     */
    export class RowCollection extends RowColCollection {
        _maxLevel = -1;

        /**
         * Gets the maximum group level in the grid.
         *
         * @return The maximum group level or -1 if the grid has no group rows.
         */
        get maxGroupLevel(): number {
            this._update();
            return this._maxLevel;
        }

       // override to keep track of the maximum group level
        _update(): boolean {
            if (super._update()) {
                this._maxLevel = -1;
                for (var i = 0; i < this.length; i++) {
                    var gr = tryCast(this[i], GroupRow);
                    if (gr && gr.level > this._maxLevel) {
                        this._maxLevel = gr.level;
                    }
                }
                return true;
            }
            return false;
        }
    }
}
module wijmo.grid {
    'use strict';

    /**
     * Contains information about the part of a @see:FlexGrid control that exists at 
     * a specified page coordinate.
     */
    export class HitTestInfo {
        _g: FlexGrid;
        _p: GridPanel;
        _pt: Point;
        _row = -1;
        _col = -1;
        _edge = 0; // left, top, right, bottom: 1, 2, 4, 8
        static _SZEDGE = [5, 30]; // distance to cell border (mouse, touch)
        
        /**
         * Initializes a new instance of the @see:HitTestInfo class.
         *
         * @param grid The @see:FlexGrid control or @see:GridPanel to investigate.
         * @param pt The @see:Point object in page coordinates to investigate.
         */
        constructor(grid: any, pt: any) {
            var g: FlexGrid;

            // check parameters
            if (grid instanceof FlexGrid) {
                g = this._g = grid;
            } else if (grid instanceof GridPanel) {
                this._p = grid;
                g = this._g = this._p.grid;
            } else {
                throw 'First parameter should be a FlexGrid or GridPanel.';
            }
            pt = mouseToPage(pt);
            this._pt = pt.clone();

            // get the variables we need
            var rc = g.controlRect,
                sz = g._szClient,
                tlp = g.topLeftCells,
                etl = g._eTL,
                hdrVis = g.headersVisibility,
                hv = HeadersVisibility,
                tlWid = (hdrVis & hv.Row) ? tlp.columns.getTotalSize() : 0,
                tlHei = (hdrVis & hv.Column) ? tlp.rows.getTotalSize() : 0,
                tlHeiSticky = (hdrVis & hv.Column) ? tlHei + etl.offsetTop : 0,
                ebl = g._eBL,
                blHei = ebl.offsetHeight;

            // convert page to control coordinates
            pt.x = Math.max(0, pt.x - rc.left);
            pt.y = Math.max(0, pt.y - rc.top);

            // account for right to left
            if (this._g._rtl) {
                pt.x = rc.width - pt.x;
            }

            // find out which panel was clicked
            if (!this._p &&
                pt.x >= 0 && pt.y >= etl.offsetTop &&
                sz && pt.x <= sz.width + tlWid && pt.y <= sz.height + tlHeiSticky) {
                if (pt.y <= tlHeiSticky) { // topleft/columnheaders
                    this._p = pt.x <= tlWid ? g.topLeftCells : g.columnHeaders;
                } else if (pt.y <= ebl.offsetTop) { // rowheaders/cells
                    this._p = pt.x <= tlWid ? g.rowHeaders : g.cells;
                } else { // bottomleft/columnfooters
                    this._p = pt.x <= tlWid ? g.bottomLeftCells : g.columnFooters;
                }
            }

            // if we have a panel, get the coordinates
            if (this._p != null) {

                // account for frozen rows/cols
                var rows = this._p.rows,
                    cols = this._p.columns,
                    ct = this._p.cellType,
                    ptFrz = this._p._getFrozenPos(),
                    totHei =
                        (ct == CellType.TopLeft || ct == CellType.ColumnHeader) ? tlHei :
                        (ct == CellType.BottomLeft || ct == CellType.ColumnFooter) ? blHei :
                        rows.getTotalSize(),
                    totWid =
                        (ct == CellType.TopLeft || ct == CellType.BottomLeft || ct == CellType.RowHeader) ? tlWid :
                        cols.getTotalSize();

                // adjust y for scrolling/freezing
                if (ct == CellType.RowHeader || ct == CellType.Cell) {
                    pt.y -= tlHei; // discount header height without 'stickiness'
                    if (pt.y > ptFrz.y || ptFrz.y <= 0) {
                        pt.y -= g._ptScrl.y;
                        pt.y += this._p._getOffsetY(); // account for IE's CSS limitations...
                    }
                } else if (ct == CellType.BottomLeft || ct == CellType.ColumnFooter) {
                    pt.y -= ebl.offsetTop;
                }

                // adjust x for scrolling/freezing
                if (ct == CellType.ColumnHeader || ct == CellType.Cell || ct == CellType.ColumnFooter) {
                    pt.x -= tlWid;
                    if (pt.x > ptFrz.x || ptFrz.x <= 0) {
                        pt.x -= g._ptScrl.x;
                    }
                }

                // enable mouse operations while in "sticky" mode
                if (ct == CellType.ColumnHeader || ct == CellType.TopLeft) {
                    pt.y -= (tlHeiSticky - tlHei);
                }

                // get row and column
                this._row = pt.y > totHei ? -1 : rows.getItemAt(pt.y);
                this._col = pt.x > totWid ? -1 : cols.getItemAt(pt.x);
                if (this._row < 0 || this._col < 0) {
                    this._p = null;
                    return;
                }

                // get edges (larger if touching)
                this._edge = 0;
                var szEdge = HitTestInfo._SZEDGE[this._g.isTouching ? 1 : 0];
                if (this._col > -1) {
                    var col = cols[this._col];
                    if (pt.x - col.pos <= szEdge) {
                        this._edge |= 1; // left
                    }
                    if (col.pos + col.renderSize - pt.x <= szEdge) {
                        this._edge |= 4; // right
                    }
                }
                if (this._row > -1) {
                    var row = rows[this._row];
                    if (pt.y - row.pos <= szEdge) {
                        this._edge |= 2; // top
                    }
                    if (row.pos + row.renderSize - pt.y <= szEdge) {
                        this._edge |= 8; // bottom
                    }
                }
            }
        }
        /**
         * Gets the point in control coordinates that this @see:HitTestInfo refers to.
         */
        get point(): Point {
            return this._pt;
        }
        /**
         * Gets the cell type at the specified position.
         */
        get cellType(): CellType {
            return this._p ? this._p.cellType : grid.CellType.None;
        }
        /**
         * Gets the grid panel at the specified position.
         */
        get panel(): GridPanel {
            return this._p;
        }
        /**
         * Gets the row index of the cell at the specified position.
         */
        get row(): number {
            return this._row;
        }
        /**
         * Gets the column index of the cell at the specified position.
         */
        get col(): number {
            return this._col;
        }
        /**
         * Gets the cell range at the specified position.
         */
        get range(): CellRange {
            return new CellRange(this._row, this._col);
        }
        /**
         * Gets a value that indicates whether the mouse is near the left edge of the cell.
         */
        get edgeLeft(): boolean {
            return (this._edge & 1) != 0;
        }
        /**
         * Gets a value that indicates whether the mouse is near the top edge of the cell.
         */
        get edgeTop(): boolean {
            return (this._edge & 2) != 0;
        }
        /**
         * Gets a value that indicates whether the mouse is near the right edge of the cell.
         */
        get edgeRight(): boolean {
            return (this._edge & 4) != 0;
        }
        /**
         * Gets a value that indicates whether the mouse is near the bottom edge of the cell.
         */
        get edgeBottom(): boolean {
            return (this._edge & 8) != 0;
        }
   }
}
module wijmo.grid {
    'use strict';

    /**
     * Specifies constants that define which areas of the grid support cell merging.
     */
    export enum AllowMerging
    {
        /** No merging. */ 
        None = 0,
        /** Merge scrollable cells. */ 
        Cells = 1,
        /** Merge column headers. */
        ColumnHeaders = 2,
        /** Merge row headers. */
        RowHeaders = 4,
        /** Merge column and row headers. */
        AllHeaders = ColumnHeaders | RowHeaders,
        /** Merge all areas. */
        All = Cells | AllHeaders
    }

    /**
     * Defines the @see:FlexGrid's cell merging behavior.
     *
     * An instance of this class is automatically created and assigned to 
     * the grid's @see:FlexGrid.mergeManager property to implement the
     * grid's default merging behavior. 
     *
     * If you want to customize the default merging behavior, create a class 
     * that derives from @see:MergeManager and override the @see:getMergedRange
     * method.
     */
    export class MergeManager {
        _g: FlexGrid;

        /**
         * Initializes a new instance of the @see:MergeManager class.
         *
         * @param g The @see:FlexGrid object that owns this @see:MergeManager.
         */
        constructor(g: FlexGrid) {
            this._g = g;
        }

        /**
         * Gets a @see:CellRange that specifies the merged extent of a cell
         * in a @see:GridPanel.
         *
         * @param p The @see:GridPanel that contains the range.
         * @param r The index of the row that contains the cell.
         * @param c The index of the column that contains the cell.
         * @param clip Whether to clip the merged range to the grid's current view range.
         * @return A @see:CellRange that specifies the merged range, or null if the cell is not merged.
         */
        getMergedRange(p: GridPanel, r: number, c: number, clip = true): CellRange {
            var rng: CellRange,
                vr: CellRange,
                ct = p.cellType,
                cols = p.columns,
                rows = p.rows,
                row = rows[r],
                col = cols[c];

            // no merging in new row template (TFS 82235)
            if (row instanceof _NewRowTemplate) {
                return null;
            }

            // merge cells in group rows
            if (row instanceof GroupRow && row.dataItem instanceof collections.CollectionViewGroup) {
                rng = new CellRange(r, c);

                // expand left and right preserving aggregates
                if (col.aggregate == Aggregate.None) {
                    while (rng.col > 0 &&
                        cols[rng.col - 1].aggregate == Aggregate.None &&
                        rng.col != cols.frozen) {
                        rng.col--;
                    }
                    while (rng.col2 < cols.length - 1 &&
                        cols[rng.col2 + 1].aggregate == Aggregate.None &&
                        rng.col2 + 1 != cols.frozen) {
                        rng.col2++;
                    }
                }

                // don't start range with invisible columns
                while (rng.col < c && !cols[rng.col].visible) {
                    rng.col++;
                }

                // return merged range
                return rng.isSingleCell ? null : rng; 
            }

            // honor grid's allowMerging setting
            var done = false;
            switch (this._g.allowMerging) {
                case AllowMerging.None:
                    done = true;
                    break;
                case AllowMerging.Cells:
                    done = ct != CellType.Cell;
                    break;
                case AllowMerging.ColumnHeaders:
                    done = ct != CellType.ColumnHeader && ct != CellType.TopLeft;
                    break;
                case AllowMerging.RowHeaders:
                    done = ct != CellType.RowHeader && ct != CellType.TopLeft;
                    break;
                case AllowMerging.AllHeaders:
                    done = ct == CellType.Cell;
                    break;
            }
            if (done) {
                return null;
            }

            // merge up and down columns
            if (cols[c].allowMerging) {
                rng = new CellRange(r, c);

                // clip to current viewport
                var rMin = 0,
                    rMax = rows.length - 1;
                if (r >= rows.frozen) {
                    if (clip && (ct == CellType.Cell || ct == CellType.RowHeader)) {
                        vr = p._getViewRange(true);
                        rMin = vr.topRow;
                        rMax = vr.bottomRow;
                    }
                } else {
                    rMax = rows.frozen - 1;
                }

                // expand up and down
                for (var tr = r - 1; tr >= rMin && this._mergeCell(p, tr, c, r, c); tr--) {
                    rng.row = tr;
                }
                for (var br = r + 1; br <= rMax && this._mergeCell(p, r, c, br, c); br++) {
                    rng.row2 = br;
                }

                // don't start range with invisible rows
                while (rng.row < r && !rows[rng.row].visible) {
                    rng.row++;
                }

                // done
                if (!rng.isSingleCell) {
                    return rng;
                }
            }

            // merge left and right along rows
            if (rows[r].allowMerging) {
                rng = new CellRange(r, c);

                // get merging limits
                var cMin = 0,
                    cMax = cols.length - 1;
                if (c >= cols.frozen) {
                    if (clip && (ct == CellType.Cell || ct == CellType.ColumnHeader)) {
                        vr = p._getViewRange(true);
                        cMin = vr.leftCol;
                        cMax = vr.rightCol;
                    }
                } else {
                    cMax = cols.frozen - 1;
                }

                // expand left and right
                for (var cl = c - 1; cl >= cMin && this._mergeCell(p, r, cl, r, c); cl--) {
                    rng.col = cl;
                }
                for (var cr = c + 1; cr <= cMax && this._mergeCell(p, r, c, r, cr); cr++) {
                    rng.col2 = cr;
                }

                // don't start range with invisible columns
                while (rng.col < c && !cols[rng.col].visible) {
                    rng.col++;
                }

                // done
                if (!rng.isSingleCell) {
                    return rng;
                }
            }

            // no merging...
            return null;
        }

        // check whether two cells should be merged
        _mergeCell(p: GridPanel, r1: number, c1: number, r2: number, c2: number) {

            // group rows and new row templates are handled separately
            var row1 = p.rows[r1],
                row2 = p.rows[r2];
            if (row1 instanceof GroupRow || row1 instanceof _NewRowTemplate ||
                row2 instanceof GroupRow || row2 instanceof _NewRowTemplate) {
                return false;
            }

            // no merging across freezing boundaries
            if (r1 != r2 && p.rows.isFrozen(r1) != p.rows.isFrozen(r2)) {
                return false;
            }
            if (c1 != c2 && p.columns.isFrozen(c1) != p.columns.isFrozen(c2)) {
                return false;
            }

            // no vertical merging if the range is already merged horizontally
            if (r1 != r2) {
                if (c1 > 0) {
                    if ((row1.allowMerging && this._mergeCell(p, r1, c1 - 1, r1, c1)) ||
                        (row2.allowMerging && this._mergeCell(p, r2, c1 - 1, r2, c1))) {
                        return false;
                    }
                }
                if (c2 < p.columns.length - 1) {
                    if ((row1.allowMerging && this._mergeCell(p, r1, c2, r1, c2 + 1)) ||
                        (row2.allowMerging && this._mergeCell(p, r2, c2, r2, c2 + 1))) {
                        return false;
                    }
                }
            }

            // no merging if the data is different
            if (p.getCellData(r1, c1, true) != p.getCellData(r2, c2, true)) {
                return false;
            }

            // OK to merge
            return true;
        }
    }
}

module wijmo.grid {
    'use strict';

    /**
     * Represents a data map for use with a column's @see:Column.dataMap property.
     *
     * Data maps provide the grid with automatic look up capabilities. For example, 
     * you may want to display a customer name instead of his ID, or a color name 
     * instead of its RGB value.
     *
     * The code below binds a grid to a collection of products, then assigns a
     * @see:DataMap to the grid's 'CategoryID' column so the grid displays the
     * category names rather than the raw IDs.
     *
     * The grid takes advantage of data maps also for editing. If the <b>wijmo.input</b>
     * module is loaded, then when editing data-mapped columns the grid will show
     * a drop-down list containing the values on the map.
     *
     * <pre>
     * // bind grid to products
     * var flex = new wijmo.grid.FlexGrid();
     * flex.itemsSource = products;
     * // map CategoryID column to show category name instead of ID
     * var col = flex.columns.getColumn('CategoryID');
     * col.dataMap = new wijmo.grid.DataMap(categories, 'CategoryID', 'CategoryName');
     * </pre>
     *
     * In general, data maps apply to whole columns. However, there are situations
     * where you may want to restrict the options available for a cell based on a
     * value on a different column. For example, if you have "Country" and "City"
     * columns, you will probably want to restrict the cities based on the current
     * country.
     *
     * There are two ways you can implement these "dynamic" data maps:
     *
     * <ol>
     *   <li>
     *     If the @see:DataMap is just a list of strings, you can change it before
     *     the grid enters edit mode. In this case, the cells contain the string
     *     being displayed, and changing the map won't affect other cells in the
     *     same column.
     *     This fiddle demonstrates:
     *     <a href="http://jsfiddle.net/Wijmo5/8brL80r8/">show me</a>.
     *   </li>
     *   <li>
     *     If the @see:DataMap is a real map (stores key values in the cells, shows
     *     a corresponding string), then you can apply a filter to restrict the
     *     values shown in the drop-down. The @see:DataMap will still contain the
     *     same keys and values, so other cells in the same column won't be disturbed
     *     by the filter.
     *     This fiddle demonstrates:
     *     <a href="http://jsfiddle.net/Wijmo5/xborLd4t/">show me</a>.
     *   </li>
     * </ol>
     */
    export class DataMap {
        _cv: collections.ICollectionView;
        _keyPath: string;
        _displayPath: string;
        _editable: boolean;
        _sortByKey: boolean;

        /**
         * Initializes a new instance of the @see:DataMap class.
         *
         * @param itemsSource An array or @see:ICollectionView that contains the items to map.
         * @param selectedValuePath The name of the property that contains the keys (data values).
         * @param displayMemberPath The name of the property to use as the visual representation of the items.
         */
        constructor(itemsSource: any, selectedValuePath?: string, displayMemberPath?: string) {

            // turn arrays into real maps
            if (isArray(itemsSource) && !selectedValuePath && !displayMemberPath) {
                var arr = [];
                for (var i = 0; i < itemsSource.length; i++) {
                    arr.push({ value: itemsSource[i] });
                }
                itemsSource = arr;
                selectedValuePath = displayMemberPath = 'value';
            }

            // initialize map
            this._cv = asCollectionView(itemsSource);
            this._keyPath = asString(selectedValuePath, false);
            this._displayPath = asString(displayMemberPath, false);

            // notify listeners when the map changes
            this._cv.collectionChanged.addHandler(this.onMapChanged, this);
        }
        /**
         * Gets or sets a value that determines whether to use mapped (display)
         * or raw values when sorting the data.
         */
        get sortByDisplayValues(): boolean {
            return this._sortByKey != true;
        }
        set sortByDisplayValues(value: boolean) {
            this._sortByKey = !asBoolean(value);
        }
        /**
         * Gets the @see:ICollectionView object that contains the map data.
         */
        get collectionView(): collections.ICollectionView {
            return this._cv;
        }
        /**
         * Gets the name of the property to use as a key for the item (data value).
         */
        get selectedValuePath(): string {
            return this._keyPath;
        }
        /**
         * Gets the name of the property to use as the visual representation of the item.
         */
        get displayMemberPath(): string {
            return this._displayPath;
        }
        /**
         * Gets the key that corresponds to a given display value.
         *
         * @param displayValue The display value of the item to retrieve.
         */
        getKeyValue(displayValue: string): any {
            var index = this._indexOf(displayValue, this._displayPath, false);
            return index > -1 ? this._cv.sourceCollection[index][this._keyPath] : null;//displayValue;
        }
        /**
         * Gets the display value that corresponds to a given key.
         *
         * @param key The key of the item to retrieve.
         */
        getDisplayValue(key: any): any {
            var index = this._indexOf(key, this._keyPath, true);
            return index > -1 ? this._cv.sourceCollection[index][this._displayPath]: key;
        }
        /**
         * Gets an array with all of the display values on the map.
         *
         * @param dataItem Data item for which to get the display items.
         * This parameter is optional. If not provided, all possible display
         * values should be returned.
         */
        getDisplayValues(dataItem?: any): string[] {
            var values = [];
            if (this._cv && this._displayPath) {
                var items = this._cv.items; // << list filtered/sorted values
                for (var i = 0; i < items.length; i++) {
                    values.push(items[i][this._displayPath]);
                }
            }
            return values;
        }
        /**
         * Gets an array with all of the keys on the map.
         */
        getKeyValues(): string[] {
            var values = [];
            if (this._cv && this._keyPath) {
                var items = this._cv.items; // << list filtered/sorted values
                for (var i = 0; i < items.length; i++) {
                    values.push(items[i][this._keyPath]);
                }
            }
            return values;
        }
        /**
         * Gets or sets a value that indicates whether users should be allowed to enter
         * values that are not present on the @see:DataMap.
         *
         * In order for a @see:DataMap to be editable, the @see:selectedValuePath and
         * @see:displayMemberPath must be set to the same value.
         */
        get isEditable(): boolean {
            return this._editable;
        }
        set isEditable(value: boolean) {
            this._editable = asBoolean(value);
        }
        /**
         * Occurs when the map data changes.
         */
        mapChanged = new Event();
        /**
         * Raises the @see:mapChanged event.
         */
        onMapChanged(e?: EventArgs) {
            this.mapChanged.raise(this, e);
        }

        // implementation

        // gets the index of a value in the sourceCollection (not the view)
        private _indexOf(value: any, path: string, caseSensitive: boolean) {
            if (this._cv && path) {
                var sval = value != null ? value.toString() : '',
                    lcval = caseSensitive ? sval : sval.toLowerCase();

                // look for items
                var items = this._cv.sourceCollection;
                for (var i = 0; i < items.length; i++) {
                    var item = items[i][path];

                    // straight comparison
                    if (item == value) {
                        return i;
                    }

                    // case-insensitive comparison
                    if (!caseSensitive && item.length == lcval.length && item.toLowerCase() == lcval) {
                        return i;
                    }

                    // string-based comparison (like JS objects) 140577
                    if (item != null && item.toString() == sval) {
                        return i;
                    }
                }
            }
            return -1;
        }
    }

}
module wijmo.grid {
    'use strict';

    /**
     * Specifies constants that define the selection behavior.
     */
    export enum SelectionMode {
        /** The user cannot select cells using the mouse or keyboard. */
        None,
        /** The user can select only a single cell at a time. */
        Cell,
        /** The user can select contiguous blocks of cells. */
        CellRange,
        /** The user can select a single row at a time. */
        Row,
        /** The user can select contiguous rows. */
        RowRange,
        /** The user can select non-contiguous rows. */
        ListBox
    }

    /**
     * Specifies constants that represent the selected state of a cell.
     */
    export enum SelectedState {
        /** The cell is not selected. */
        None,
        /** The cell is selected but is not the active cell. */
        Selected,
        /** The cell is selected and is the active cell. */
        Cursor,
    }

    /**
     * Specifies constants that represent a type of movement for the selection.
     */
    export enum SelMove {
        /** Do not change the selection. */
        None,
        /** Select the next visible cell. */
        Next,
        /** Select the previous visible cell. */
        Prev,
        /** Select the first visible cell in the next page. */
        NextPage,
        /** Select the first visible cell in the previous page. */
        PrevPage,
        /** Select the first visible cell. */
        Home,
        /** Select the last visible cell. */
        End,
        /** Select the next visible cell skipping rows if necessary. */
        NextCell,
        /** Select the previous visible cell skipping rows if necessary. */
        PrevCell
    }

    /**
     * Handles the grid's selection.
     */
    export class _SelectionHandler {
        _g: FlexGrid;
        _sel = new CellRange(0, 0);
        _mode = SelectionMode.CellRange;

        /**
         * Initializes a new instance of the @see:_SelectionHandler class.
         *
         * @param g @see:FlexGrid that owns this @see:_SelectionHandler.
         */
        constructor(g: FlexGrid) {
            this._g = g;
        }

        /**
         * Gets or sets the current selection mode.
         */
        get selectionMode(): SelectionMode {
            return this._mode;
        }
        set selectionMode(value: SelectionMode) {
            if (value != this._mode) {

                // update ListBox selection when switching modes
                if (value == SelectionMode.ListBox || this._mode == SelectionMode.ListBox) {
                    var rows = this._g.rows;
                    for (var i = 0; i < rows.length; i++) {
                        var row = rows[i],
                            sel = (value == SelectionMode.ListBox) ? this._sel.containsRow(i) : false;
                        row._setFlag(RowColFlags.Selected, sel, true);
                    }
                }

                // collapse selection when switching to None/Cell/Row modes (TFS 130691)
                switch (value) {
                    case SelectionMode.None:
                        this._sel.setRange(-1, -1);
                        break;
                    case SelectionMode.Cell:
                        this._sel.row2 = this._sel.row;
                        this._sel.col2 = this._sel.col;
                        break;
                    case SelectionMode.Row:
                        this._sel.row2 = this._sel.row;
                        break;
                }

                // apply new mode
                this._mode = value;
                this._g.invalidate();
            }
        }
        /**
         * Gets or sets the current selection.
         */
        get selection(): CellRange {
            return this._sel;
        }
        set selection(value: CellRange) {
            this.select(value);
        }
        /**
         * Selects a cell range and optionally scrolls it into view.
         *
         * @param rng Range to select.
         * @param show Whether to scroll the new selection into view.
         */
        select(rng: any, show: any = true) {

            // allow passing in row and column indices
            if (isNumber(rng) && isNumber(show)) {
                rng = new CellRange(<number>rng, <number>show);
                show = true;
            }
            rng = asType(rng, CellRange);

            // get old and new selections
            var g = this._g,
                oldSel = this._sel,
                newSel = rng,
                lbMode = false;

            // adjust for selection mode
            switch (g.selectionMode) {

                // Cell mode: collapse range into single cell
                case SelectionMode.Cell:
                    rng.row2 = rng.row;
                    rng.col2 = rng.col;
                    break;

                // Row mode: collapse range into single row
                case SelectionMode.Row:
                    rng.row2 = rng.row;
                    break;

                // ListBox mode: remember because handling is quite different
                case SelectionMode.ListBox:
                    lbMode = true;
                    break;
            }

            // check if the selection really is changing
            // (special handling for ListBox mode when re-selecting items)
            var noChange = newSel.equals(oldSel);
            if (lbMode && newSel.row > -1 && !g.rows[newSel.row].isSelected) {
                noChange = false;
            }

            // no change? done
            if (noChange) {
                if (show) {
                    g.scrollIntoView(newSel.row, newSel.col);
                }
                return;
            }

            // raise selectionChanging event
            var e = new CellRangeEventArgs(g.cells, newSel);
            if (!g.onSelectionChanging(e)) {
                return;
            }

            // ListBox mode: update Selected flag and refresh to show changes
            // (after firing the selectionChanging cancelable event)
            if (lbMode) {
                for (var i = 0; i < g.rows.length; i++) {
                    (<Row>g.rows[i])._setFlag(RowColFlags.Selected, newSel.containsRow(i), true);
                }
                g.refreshCells(false, true, true);
            }

            // validate selection after the change
            newSel.row = Math.min(newSel.row, g.rows.length - 1);
            newSel.row2 = Math.min(newSel.row2, g.rows.length - 1);

            // update selection
            this._sel = newSel;

            // show the new selection
            g.refreshCells(false, true, true);
            if (show) {
                g.scrollIntoView(newSel.row, newSel.col);
            }

            // update collectionView cursor
            if (g.collectionView) {
                var index = g._getCvIndex(newSel.row);
                g.collectionView.moveCurrentToPosition(index);
            }

            // raise selectionChanged event
            g.onSelectionChanged(e);
        }
        /**
         * Moves the selection by a specified amount in the vertical and horizontal directions.
         * @param rowMove How to move the row selection.
         * @param colMove How to move the column selection.
         * @param extend Whether to extend the current selection or start a new one.
         */
        moveSelection(rowMove: SelMove, colMove: SelMove, extend: boolean) {
            var row, col,
                g = this._g,
                rows = g.rows,
                cols = g.columns,
                rng = this._getReferenceCell(rowMove, colMove, extend),
                pageSize = Math.max(0, g._szClient.height - g.columnHeaders.height);

            // handle next cell with wrapping
            if (colMove == SelMove.NextCell) {
                col = cols.getNextCell(rng.col, SelMove.Next, pageSize);
                row = rng.row;
                if (col == rng.col) {
                    row = rows.getNextCell(row, SelMove.Next, pageSize);
                    if (row > rng.row) {
                        col = cols.getNextCell(0, SelMove.Next, pageSize);
                        col = cols.getNextCell(col, SelMove.Prev, pageSize);
                    }
                }
                g.select(row, col);

            } else if (colMove == SelMove.PrevCell) {

                col = cols.getNextCell(rng.col, SelMove.Prev, pageSize);
                row = rng.row;
                if (col == rng.col) { // reached first column, wrap to previous row
                    row = rows.getNextCell(row, SelMove.Prev, pageSize);
                    if (row < rng.row) {
                        col = cols.getNextCell(cols.length - 1, SelMove.Prev, pageSize);
                        col = cols.getNextCell(col, SelMove.Next, pageSize);
                    }
                }
                g.select(row, col);

            } else {

                // get target row, column
                row = rows.getNextCell(rng.row, rowMove, pageSize);
                col = cols.getNextCell(rng.col, colMove, pageSize);

                // extend or select
                if (extend) {
                    var sel = g._selHdl._sel;
                    g.select(new CellRange(row, col, sel.row2, sel.col2));
                } else {
                    g.select(row, col);
                }
            }
        }

        // get reference cell for selection change, taking merging into account
        private _getReferenceCell(rowMove: SelMove, colMove: SelMove, extend: boolean): CellRange {
            var g = this._g,
                sel = g._selHdl._sel,
                rng = g.getMergedRange(g.cells, sel.row, sel.col);

            // not merging? use selection as a reference
            if (!rng || rng.isSingleCell) {
                return sel;
            }

            // clone range and set reference cell within the range
            rng = rng.clone();
            switch (rowMove) {
                case SelMove.Next:
                case SelMove.NextCell:
                    rng.row = rng.bottomRow;
                    break;
                case SelMove.None:
                    rng.row = sel.row;
                    break;
            }
            switch (colMove) {
                case SelMove.Next:
                case SelMove.NextCell:
                    rng.col = rng.rightCol;
                    break;
                case SelMove.None:
                    rng.col = sel.col;
                    break;
            }

            // done
            return rng;
        }

        // adjusts a selection to reflect the current selection mode
        /*private*/ _adjustSelection(rng: CellRange): CellRange {
            switch (this._mode) {
                case SelectionMode.Cell:
                    return new CellRange(rng.row, rng.col, rng.row, rng.col);
                case SelectionMode.Row:
                    return new CellRange(rng.row, 0, rng.row, this._g.columns.length - 1);
                case SelectionMode.RowRange:
                case SelectionMode.ListBox:
                    return new CellRange(rng.row, 0, rng.row2, this._g.columns.length - 1);
            }
            return rng;
        }
    }
}

module wijmo.grid {
    'use strict';

    /**
     * Handles the grid's keyboard commands.
     */
    export class _KeyboardHandler {
        _g: FlexGrid;

        /**
         * Initializes a new instance of the @see:_KeyboardHandler class.
         *
         * @param g @see:FlexGrid that owns this @see:_KeyboardHandler.
         */
        constructor(g: FlexGrid) {
            this._g = g;
            g.addEventListener(g.hostElement, 'keypress', this._keypress.bind(this));
            g.addEventListener(g.hostElement, 'keydown', this._keydown.bind(this));
        }

        // handles the key down event (selection)
        /*private*/ _keydown(e: KeyboardEvent) {
            var g = this._g,
                sel = g.selection,
                ctrl = e.ctrlKey || e.metaKey,
                shift = e.shiftKey,
                target = <HTMLElement>e.target,
                handled = true;

            if (g.isRangeValid(sel) && !e.defaultPrevented) {

                // allow input elements that don't belong to us to handle keys (TFS 131138, 191989)
                if (!g.activeEditor && g._isInputElement(target) && !target.getAttribute('wj-part')) {
                    return;
                }

                // pre-process keys while editor is active
                if (g.activeEditor && g._edtHdl._keydown(e)) {
                    return;
                }

                // get the variables we need
                var gr = <GroupRow>tryCast(g.rows[sel.row], GroupRow),
                    ecv = <collections.IEditableCollectionView>tryCast(g.collectionView, 'IEditableCollectionView'),
                    keyCode = e.keyCode;

                // handle clipboard
                if (g.autoClipboard) {

                    // copy: ctrl+c or ctrl+Insert
                    if (ctrl && (keyCode == 67 || keyCode == 45)) {
                        var args = new CellRangeEventArgs(g.cells, sel);
                        if (g.onCopying(args)) {
                            var text = g.getClipString();
                            Clipboard.copy(text);
                            g.onCopied(args);
                        }
                        e.stopPropagation();
                        return;
                    }

                    // paste: ctrl+v or shift+Insert
                    if ((ctrl && keyCode == 86) || (shift && keyCode == 45)) {
                        if (!g.isReadOnly) {
                            var args = new CellRangeEventArgs(g.cells, sel);
                            if (g.onPasting(args)) {
                                Clipboard.paste((text) => {
                                    g.setClipString(text);
                                    g.onPasted(args);
                                });
                            }
                        }
                        e.stopPropagation();
                        return;
                    }
                }

                // reverse left/right keys when rendering in right-to-left
                if (g._rtl) { 
                    switch (keyCode) {
                        case Key.Left:
                            keyCode = Key.Right;
                            break;
                        case Key.Right:
                            keyCode = Key.Left;
                            break;
                    }
                }

                // default key handling
                switch (keyCode) {
                    case 65: // ctrl+A: select all
                        if (ctrl) {
                            g.select(new CellRange(0, 0, g.rows.length - 1, g.columns.length - 1));
                        } else {
                            handled = false;
                        }
                        break;
                    case Key.Left:
                        if (sel.isValid && sel.col == 0 && gr != null && !gr.isCollapsed && gr.hasChildren) {
                            gr.isCollapsed = true;
                        } else {
                            this._moveSel(SelMove.None, ctrl ? SelMove.Home : SelMove.Prev, shift);
                        }
                        break;
                    case Key.Right:
                        if (sel.isValid && sel.col == 0 && gr != null && gr.isCollapsed) {
                            gr.isCollapsed = false;
                        } else {
                            this._moveSel(SelMove.None, ctrl ? SelMove.End : SelMove.Next, shift);
                        }
                        break;
                    case Key.Up: // alt-up toggles the ListBox
                        if (e.altKey && g._edtHdl._toggleListBox(this._g.selection)) {
                            break;
                        }
                        this._moveSel(ctrl ? SelMove.Home : SelMove.Prev, SelMove.None, shift);
                        break;
                    case Key.Down: // alt-down opens the ListBox
                        if (e.altKey && g._edtHdl._toggleListBox(this._g.selection)) {
                            break;
                        }
                        this._moveSel(ctrl ? SelMove.End : SelMove.Next, SelMove.None, shift);
                        break;
                    case Key.PageUp:
                        this._moveSel(SelMove.PrevPage, SelMove.None, shift);
                        break;
                    case Key.PageDown:
                        this._moveSel(SelMove.NextPage, SelMove.None, shift);
                        break;
                    case Key.Home:
                        this._moveSel(ctrl ? SelMove.Home : SelMove.None, SelMove.Home, shift);
                        break;
                    case Key.End:
                        this._moveSel(ctrl ? SelMove.End : SelMove.None, SelMove.End, shift);
                        break;
                    case Key.Tab:
                        this._moveSel(SelMove.None, shift ? SelMove.PrevCell : SelMove.NextCell, false);
                        break;
                    case Key.Enter:
                        this._moveSel(shift ? SelMove.Prev : SelMove.Next, SelMove.None, false);
                        if (!shift && ecv && ecv.currentEditItem != null) {
                            g._edtHdl._commitRowEdits(); // in case we're at the last row (TFS 105989)
                        }
                        break;
                    case Key.Escape:
                        if (ecv) {
                            if (ecv.currentEditItem != null) {
                                ecv.cancelEdit();
                            }
                            if (ecv.currentAddItem != null) {
                                ecv.cancelNew();
                            }
                        }
                        g._mouseHdl.resetMouseState();
                        break;
                    case Key.Delete:
                        handled = this._deleteSel();
                        break;
                    case Key.F2:
                        handled = g.startEditing(true);
                        break;
                    case Key.F4:
                        handled = g._edtHdl._toggleListBox(this._g.selection);
                        break;
                    case Key.Space:
                        handled = g.startEditing(true);
                        if (handled) {
                            setTimeout(() => {
                                var edt = g.activeEditor;
                                if (edt) {
                                    if (edt.type == 'checkbox') {
                                        edt.checked = !edt.checked;
                                        g.finishEditing();
                                    } else {
                                        setSelectionRange(edt, edt.value.length);
                                    }
                                }
                            });
                        }
                        break;
                    default:
                        handled = false;
                        break;
                }
                if (handled) {
                    g.focus(); // http://wijmo.com/topic/angular-2-focus-issue-with-wj-input-number-as-rendering-cell-of-flexgrid/
                    e.preventDefault();
                    e.stopPropagation();
                }
            }
        }

        // handles the key press event (start editing or try auto-complete)
        private _keypress(e: KeyboardEvent) {

            // forward key to editor (auto-complete) or handle ourselves
            var g = this._g;
            if (g.activeEditor) {
                g._edtHdl._keypress(e);
            } else if (e.charCode > Key.Space) {
                if (g.startEditing(false) && g.activeEditor) {
                    setTimeout(() => {
                        var edt = g.activeEditor;
                        if (edt && edt.type != 'checkbox') {
                            edt.value = String.fromCharCode(e.charCode); // FireFox needs this...
                            setSelectionRange(edt, 1);
                            edt.dispatchEvent(g._edtHdl._evtInput); // to apply mask (TFS 131232)
                            g._edtHdl._keypress(e); // to start auto-complete
                        }
                    });
                }
            }
            e.stopPropagation();
        }

        // move the selection
        private _moveSel(rowMove: SelMove, colMove: SelMove, extend: boolean) {
            if (this._g.selectionMode != SelectionMode.None) {
                this._g._selHdl.moveSelection(rowMove, colMove, extend);
            }
        }

        // delete the selected rows
        private _deleteSel(): boolean {
            var g = this._g,
                ecv = <collections.IEditableCollectionView>tryCast(g.collectionView, 'IEditableCollectionView'),
                sel = g.selection,
                rows = g.rows,
                selRows = [];

            // if g.allowDelete and ecv.canRemove, and not editing/adding, (TFS 87718)
            // and the grid allows deleting items, then delete selected rows
            if (g.allowDelete && !g.isReadOnly &&
                (ecv == null || (ecv.canRemove && !ecv.isAddingNew && !ecv.isEditingItem))) {

                // get selected rows
                switch (g.selectionMode) {
                    case SelectionMode.CellRange:
                        if (sel.leftCol == 0 && sel.rightCol == g.columns.length - 1) {
                            for (var i = sel.topRow; i > -1 && i <= sel.bottomRow; i++) {
                                selRows.push(rows[i]);
                            }
                        }
                        break;
                    case SelectionMode.ListBox:
                        for (var i = 0; i < rows.length; i++) {
                            if (rows[i].isSelected) {
                                selRows.push(rows[i]);
                            }
                        }
                        break;
                    case SelectionMode.Row:
                        if (sel.topRow > -1) {
                            selRows.push(rows[sel.topRow]);
                        }
                        break;
                    case SelectionMode.RowRange:
                        for (var i = sel.topRow; i > -1 && i <= sel.bottomRow; i++) {
                            selRows.push(rows[i]);
                        }
                        break;
                }
            }

            // finish with row deletion
            if (selRows.length > 0) {

                // begin updates
                if (ecv) ecv.beginUpdate();
                g.beginUpdate();

                // delete selected rows
                var rng = new CellRange(),
                    e = new CellRangeEventArgs(g.cells, rng);
                for (var i = selRows.length - 1; i >= 0; i--) {
                    var r = selRows[i];
                    rng.row = rng.row2 = r.index;
                    if (g.onDeletingRow(e)) {
                        if (ecv && r.dataItem) {
                            ecv.remove(r.dataItem);
                        } else {
                            g.rows.removeAt(r.index);
                        }
                        g.onDeletedRow(e);
                    }
                }

                // finish updates
                g.endUpdate();
                if (ecv) ecv.endUpdate();

                // make sure one row is selected in ListBox mode (TFS 82683)
                if (g.selectionMode == SelectionMode.ListBox) {
                    var index = g.selection.row;
                    if (index > -1 && index < g.rows.length) {
                        g.rows[index].isSelected = true;
                    }
                }

                // handle childItemsPath (TFS 87577)
                if (g.childItemsPath && g.collectionView) {
                    g.collectionView.refresh();
                }

                // all done
                return true;
            }

            // delete cell content (if there is any) (TFS 94178)
            if (!g.isReadOnly && selRows.length == 0 && sel.isSingleCell) {
                var bcol = g._getBindingColumn(g.cells, sel.row, g.columns[sel.col]);
                if (bcol.isRequired == false || (bcol.isRequired == null && bcol.dataType == DataType.String)) { // TFS 94192
                    if (g.getCellData(sel.row, sel.col, true)) {
                        if (g.startEditing(false, sel.row, sel.col)) {
                            g.setCellData(sel.row, sel.col, '', true); // TFS 118470
                            g.finishEditing(true);
                            g.invalidate();
                            return true;
                        }
                    }
                }
            }

            // no deletion
            return false;
        }
    }
}

module wijmo.grid {
    'use strict';

    // allow resizing by dragging regular cells as well as headers
    var _AR_ALLCELLS = 4;

    /**
     * Specifies constants that define the row/column sizing behavior.
     */
    export enum AllowResizing {
        /** The user may not resize rows or columns. */
        None = 0,
        /** The user may resize columns by dragging the edge of the column headers. */
        Columns = 1,
        /** The user may resize rows by dragging the edge of the row headers. */
        Rows = 2,
        /** The user may resize rows and columns by dragging the edge of the headers. */
        Both = Rows | Columns, // 3
        /** The user may resize columns by dragging the edge of any cell. */
        ColumnsAllCells = Columns | _AR_ALLCELLS, // 5
        /** The user may resize rows by dragging the edge of any cell. */
        RowsAllCells = Rows | _AR_ALLCELLS, // 6
        /** The user may resize rows and columns by dragging the edge of any cell. */
        BothAllCells = Both | _AR_ALLCELLS // 7
    }

    /**
     * Specifies constants that define the row/column auto-sizing behavior.
     */
    export enum AutoSizeMode {
        /** Autosizing is disabled. */
        None = 0,
        /** Autosizing accounts for header cells. */
        Headers = 1,
        /** Autosizing accounts for data cells. */
        Cells = 2,
        /** Autosizing accounts for header and data cells. */
        Both = Headers | Cells
    }

    /**
     * Specifies constants that define the row/column dragging behavior.
     */
    export enum AllowDragging {
        /** The user may not drag rows or columns. */
        None = 0,
        /** The user may drag columns. */
        Columns = 1,
        /** The user may drag rows. */
        Rows = 2,
        /** The user may drag rows and columns. */
        Both = Rows | Columns
    }

    /**
     * Handles the grid's mouse commands.
     */
    export class _MouseHandler {
        _g: FlexGrid;
        _htDown: HitTestInfo;
        _eMouse: MouseEvent;
        _lbSelState: boolean;
        _lbSelRows: Object;
        _szRowCol: RowCol;
        _szStart: number;
        _szArgs: CellRangeEventArgs;
        _dragSource: any;
        _dvMarker: HTMLElement;
        _rngTarget: CellRange;

        static _SZ_MIN = 0; // minimum size allowed when resizing rows/cols

        /**
         * Initializes a new instance of the @see:_MouseHandler class.
         *
         * @param g @see:FlexGrid that owns this @see:_MouseHandler.
         */
        constructor(g: FlexGrid) {
            var host = g.hostElement;
            this._g = g;

            // mouse events: 
            // when the user presses the mouse on the control, hook up handlers to 
            // mouse move/up on the *document*, and unhook on mouse up.
            // this simulates a mouse capture (nice idea from ngGrid).
            // note: use 'document' and not 'window'; that doesn't work on Android.
            g.addEventListener(host, 'mousedown', (e: MouseEvent) => {

                // start actions on left button only: TFS 114623
                if (!e.defaultPrevented && e.button == 0) { 

                    // make sure the click does not belong to another nested grid: TFS 200695
                    if (closest(e.target, '.wj-flexgrid') == this._g.hostElement) {
                        g.addEventListener(document, 'mousemove', mouseMove);
                        g.addEventListener(document, 'mouseup', mouseUp);
                        this._mousedown(e);
                    }
                }
            });
            var mouseMove = (e: MouseEvent) => {
                this._mousemove(e);
            };
            var mouseUp = (e: MouseEvent) => {
                g.removeEventListener(document, 'mousemove');
                g.removeEventListener(document, 'mouseup');
                this._mouseup(e);
            };

            // offer to resize on mousemove (pressing the button not required)
            g.addEventListener(host, 'mousemove', this._hover.bind(this));

            // double-click to auto-size rows/columns and to enter edit mode
            g.addEventListener(host, 'dblclick', this._dblclick.bind(this));

            // prevent user from selecting grid content (as text)
            g.addEventListener(host, 'selectstart', (e) => {
                if (e.target.tagName != 'INPUT') {
                    e.preventDefault();
                }
            });

            // prevent wheel from propagating to parent elements
            g.addEventListener(host, 'wheel', (e: WheelEvent) => {
                var root = g.cells.hostElement.parentElement;
                if (root.scrollHeight > root.offsetHeight) {
                    if ((e.deltaY < 0 && root.scrollTop == 0) ||
                        (e.deltaY > 0 && root.scrollTop + root.offsetHeight >= root.scrollHeight)) {
                        if (wijmo.closest(e.target, '.wj-flexgrid') == g.hostElement) { // to handle nested grids
                            e.preventDefault();
                            e.stopPropagation();
                        }
                    }
                }
            });

            // row and column dragging
            g.addEventListener(host, 'dragstart', this._dragstart.bind(this));
            g.addEventListener(host, 'dragover', this._dragover.bind(this));
            g.addEventListener(host, 'dragleave', this._dragover.bind(this));
            g.addEventListener(host, 'drop', this._drop.bind(this));
            g.addEventListener(host, 'dragend', this._dragend.bind(this));

            // create target indicator element
            this._dvMarker = createElement('<div class="wj-marker">&nbsp;</div>');
        }

        /**
         * Resets the mouse state.
         */
        resetMouseState() {

            // because dragEnd fires too late in FireFox...
            if (this._dragSource) {
                this._dragSource.style.opacity = 1;
            }
            this._showDragMarker(null);

            // reset cursor state (if the grid hasn't been disposed)
            var host = this._g.hostElement;
            if (host) {
                host.style.cursor = 'default';
            }

            // remove event listeners just in case
            var g = this._g;
            g.removeEventListener(document, 'mousemove');
            g.removeEventListener(document, 'mouseup');

            // reset everything else
            this._htDown = null;
            this._lbSelRows = null;
            this._szRowCol = null;
            this._szArgs = null;
            this._dragSource = null;
        }

        // handles the mouse down event
        private _mousedown(e: MouseEvent) {
            var g = this._g;

            // make sure control bounds are refreshed before hit-testing 
            // (in case a container has scrolled)
            g._rcBounds = null;
            var ht = g.hitTest(e),
                ct = ht.cellType;

            // ignore clicks on unknown areas
            if (ct == CellType.None) {
                g.finishEditing();
                return;
            }

            // if the user clicked an active editor, let the editor handle things
            if (ct == CellType.Cell && g.editRange && g.editRange.contains(ht.range)) {
                return;
            }

            // ignore clicks on focused input elements (TFS 135271)
            var ae = getActiveElement();
            if (e.target == ae && g._isInputElement(e.target)) {
                return;
            }

            // unless the target has the focus, give it to the grid (TFS 81949, 102177, 120430)
            if (e.target != ae) {
                g.focus();
            }

            // check where the mouse is
            this._htDown = ht;
            this._eMouse = e;

            // handle resizing
            if (this._szRowCol != null) {
                this._handleResizing(e);
                return;
            }

            // starting cell selection? special handling for ListBox mode
            switch (ct) {
                case CellType.Cell:
                    if (e.ctrlKey && g.selectionMode == SelectionMode.ListBox) {
                        this._startListBoxSelection(ht.row);
                    }
                    this._mouseSelect(e, e.shiftKey);
                    break;
                case CellType.RowHeader:
                    if ((this._g.allowDragging & AllowDragging.Rows) == 0) {
                        if (e.ctrlKey && g.selectionMode == SelectionMode.ListBox) {
                            this._startListBoxSelection(ht.row);
                        }
                        this._mouseSelect(e, e.shiftKey);
                    }
                    break;
            }

            // handle collapse/expand (after selecting the cell)
            if (ct == CellType.Cell && g.rows.maxGroupLevel > -1) {
                var gr = <GroupRow>tryCast(g.rows[ht.row], GroupRow),
                    icon = closest(e.target, '.' + CellFactory._WJC_COLLAPSE);
                if (gr && icon) { // TFS 150397
                    if (e.ctrlKey) {

                        // ctrl+click: collapse/expand entire outline to this level
                        g.collapseGroupsToLevel(gr.isCollapsed ? gr.level + 1 : gr.level);

                    } else {

                        // simple click: toggle this group
                        gr.isCollapsed = !gr.isCollapsed;
                    }

                    // done with the mouse
                    this.resetMouseState();
                    e.preventDefault();
                    return;
                }
            }
        }

        // handles the mouse move event
        private _mousemove(e: MouseEvent) {
            if (this._htDown != null) {

                // in case we lost the focus or the button (TFS 145149)
                // note that e.which doesn't work correctly in FireFox: https://bugzilla.mozilla.org/show_bug.cgi?id=1048294
                setTimeout(() => {
                    if (!e.which || !this._g.containsFocus()) {
                        this.resetMouseState();
                    }
                });

                // handle the event as usual
                this._eMouse = e;
                if (this._szRowCol) {
                    this._handleResizing(e);
                } else {
                    switch (this._htDown.cellType) {
                        case CellType.Cell:
                            this._mouseSelect(e, true);
                            break;
                        case CellType.RowHeader:
                            if ((this._g.allowDragging & AllowDragging.Rows) == 0) {
                                this._mouseSelect(e, true);
                            }
                            break;
                    }
                }
            }
        }

        // handles the mouse up event
        private _mouseup(e: MouseEvent) {

            // IE raises mouseup while touch-dragging...???
            if (this._dragSource && this._g.isTouching) {
                return;
            }

            // select all cells, finish resizing, sorting
            var htd = this._htDown;
            if (htd && htd.cellType == CellType.TopLeft && !this._szArgs && !e.defaultPrevented) {
                var g = this._g,
                    ht = g.hitTest(e);
                if (ht.panel == htd.panel && ht.row == htd.row && ht.col == htd.col) {
                    var rng = g.getMergedRange(htd.panel, htd.row, htd.col) || ht.range;
                    if (rng.row == 0 && rng.col == 0) {
                        g.select(new CellRange(0, 0, g.rows.length - 1, g.columns.length - 1));
                    }
                }
            } else if (this._szArgs) {
                this._finishResizing(e);
            } else {
                this._handleSort(e);
            }

            // done with the mouse
            this.resetMouseState();
        }

        // handles double-clicks
        private _dblclick(e: MouseEvent) {
            var g = this._g,
                ht = g.hitTest(e),
                ct = ht.cellType,
                sel = g.selection,
                rng = ht.range,
                args: CellRangeEventArgs;

            // ignore if already handled
            if (e.defaultPrevented) {
                return;
            }

            // auto-size columns
            if (ht.edgeRight && (g.allowResizing & AllowResizing.Columns)) {
                if (ct == CellType.ColumnHeader || (ct == CellType.Cell && (g.allowResizing & _AR_ALLCELLS))) {
                    e.preventDefault();
                    if (e.ctrlKey && sel.containsColumn(ht.col)) {
                        rng = sel;
                    }
                    for (var c = rng.leftCol; c <= rng.rightCol; c++) {
                        if (g.columns[c].allowResizing) {
                            args = new CellRangeEventArgs(g.cells, new CellRange(-1, c));
                            if (g.onAutoSizingColumn(args) && g.onResizingColumn(args)) {
                                g.autoSizeColumn(c);
                                g.onResizedColumn(args);
                                g.onAutoSizedColumn(args);
                            }
                        }
                    }
                } else if (ct == CellType.TopLeft) {
                    if (g.topLeftCells.columns[ht.col].allowResizing) {
                        e.preventDefault();
                        args = new CellRangeEventArgs(g.topLeftCells, new CellRange(-1, ht.col));
                        if (g.onAutoSizingColumn(args) && g.onResizingColumn(args)) {
                            g.autoSizeColumn(ht.col, true);
                            g.onAutoSizedColumn(args);
                            g.onResizedColumn(args);
                        }
                    }
                }
                this.resetMouseState();
                return;
            }

            // auto-size rows
            if (ht.edgeBottom && (g.allowResizing & AllowResizing.Rows)) {
                if (ct == CellType.RowHeader || (ct == CellType.Cell && (g.allowResizing & _AR_ALLCELLS))) {
                    if (e.ctrlKey && sel.containsRow(ht.row)) {
                        rng = sel;
                    }
                    for (var r = rng.topRow; r <= rng.bottomRow; r++) {
                        if (g.rows[r].allowResizing) {
                            args = new CellRangeEventArgs(g.cells, new CellRange(r, -1));
                            if (g.onAutoSizingRow(args) && g.onResizingRow(args)) {
                                g.autoSizeRow(r);
                                g.onResizedRow(args);
                                g.onAutoSizedRow(args);
                            }
                        }
                    }
                } else if (ct == CellType.TopLeft) {
                    if (g.topLeftCells.rows[ht.row].allowResizing) {
                        args = new CellRangeEventArgs(g.topLeftCells, new CellRange(ht.row, -1));
                        if (g.onAutoSizingRow(args) && g.onResizingRow(args)) {
                            g.autoSizeRow(ht.row, true);
                            g.onResizedRow(args);
                            g.onAutoSizedRow(args);
                        }
                    }
                }
                this.resetMouseState();
            }
        }

        // offer to resize rows/columns
        private _hover(e: MouseEvent) {

            // make sure we're hovering
            if (this._htDown == null) {
                var g = this._g,
                    ht = g.hitTest(e),
                    p = ht.panel,
                    ct = ht.cellType,
                    cursor = 'default';

                // find which row/column is being resized
                this._szRowCol = null;
                if (ct == CellType.ColumnHeader || ct == CellType.TopLeft || (ct == CellType.Cell && (g.allowResizing & _AR_ALLCELLS))) {
                    if (ht.edgeRight && (g.allowResizing & AllowResizing.Columns) != 0) {
                        this._szRowCol = this._getResizeCol(ht);
                    }
                }
                if (ct == CellType.RowHeader || ct == CellType.TopLeft || (ct == CellType.Cell && (g.allowResizing & _AR_ALLCELLS))) {
                    if (ht.edgeBottom && (g.allowResizing & AllowResizing.Rows) != 0) {
                        this._szRowCol = this._getResizeRow(ht);
                    }
                }

                // keep track of element to resize and original size
                if (this._szRowCol instanceof Column) {
                    cursor = 'col-resize';
                } else if (this._szRowCol instanceof Row) {
                    cursor = 'row-resize';
                }
                this._szStart = this._szRowCol ? this._szRowCol.renderSize : 0;

                // update the cursor to provide user feedback
                g.hostElement.style.cursor = cursor;
            }
        }
        private _getResizeCol(ht: HitTestInfo): Column {

            // start with the column under the mouse
            var cols = ht.panel.columns,
                col = cols[ht.col];

            // if the next column in the panel is visible but collapsed, switch
            for (var c = ht.col + 1; c < cols.length; c++) {
                var newCol = cols[c];
                if (newCol.visible) {
                    if (newCol.size < 1) {
                        col = newCol;
                    }
                    break;
                }
            }

            // if this is the last column on a fixed panel, and the first
            // column on the cells panel is visible but collapsed, switch
            if (ht.col == cols.length - 1) {
                if (ht.cellType == CellType.TopLeft || ht.cellType == CellType.RowHeader) {
                    cols = this._g.columns;
                    for (var c = 0; c < cols.length; c++) {
                        var newCol = cols[c];
                        if (newCol.visible) {
                            if (newCol.size < 1) {
                                col = newCol;
                            }
                            break;
                        }
                    }
                }
            }

            // return the column we got
            return col.allowResizing ? col : null;
        }
        private _getResizeRow(ht: HitTestInfo): Row {

            // start with the row under the mouse
            var rows = ht.panel.rows,
                row = rows[ht.row];

            // if the next row in the panel is visible but collapsed, switch
            for (var r = ht.row + 1; r < rows.length; r++) {
                var newRow = rows[r];
                if (newRow.visible) {
                    if (newRow.size < 1) {
                        row = newRow;
                    }
                    break;
                }
            }

            // if this is the last row on a fixed panel, and the first
            // row on the cells panel is visible but collapsed, switch
            if (ht.row == rows.length - 1) {
                if (ht.cellType == CellType.TopLeft || ht.cellType == CellType.ColumnHeader) {
                    rows = this._g.rows;
                    for (var r = 0; r < rows.length; r++) {
                        var newRow = rows[r];
                        if (newRow.visible) {
                            if (newRow.size < 1) {
                                row = newRow;
                            }
                            break;
                        }
                    }
                }
            }

            // return the column we got
            return row.allowResizing ? row : null;
        }

        // handles mouse moves while the button is pressed on the cell area
        private _mouseSelect(e: MouseEvent, extend: boolean) {
            if (this._htDown && this._htDown.panel && this._g.selectionMode != SelectionMode.None) {

                // handle the selection
                var ht = new HitTestInfo(this._htDown.panel, e);
                this._handleSelection(ht, extend);

                // done, prevent selection of content outside the grid
                // (unless the click was on an element that wants/needs it!!! TFS 191989)
                if (!this._g._isInputElement(e.target)) {
                    e.preventDefault();
                }

                // keep calling this if the user keeps the mouse outside the control without moving it
                // but don't do this in IE9, it can keep scrolling forever... TFS 110374
                // NOTE: doesn't seem to be an issue anymore, but keep the check to avoid potential regressions.
                if (!isIE9() && e.button >= 0) {
                    ht = new HitTestInfo(this._g, e);
                    if (ht.cellType != CellType.Cell && ht.cellType != CellType.RowHeader) {
                        setTimeout(() => {
                            this._mouseSelect(this._eMouse, extend);
                        }, 100);
                    }
                }
            }
        }

        // handle row and column resizing
        private _handleResizing(e: MouseEvent) {

            // prevent browser from selecting cell content
            e.preventDefault();

            // resizing column
            if (this._szRowCol instanceof Column) {
                var g = this._g,
                    pageX = e.clientX + pageXOffset, // e.pageXY doesn't work well in Chrome/zoom/touch
                    sz = Math.round(Math.max(_MouseHandler._SZ_MIN, this._szStart + (pageX - this._htDown.point.x) * (g._rtl ? -1 : 1)));
                if (this._szRowCol.renderSize != sz) {
                    if (this._szArgs == null) {
                        var panel = g.rowHeaders.columns.indexOf(this._szRowCol) > -1 ? g.rowHeaders : g.cells;
                        this._szArgs = new CellRangeEventArgs(panel, new CellRange(-1, this._szRowCol.index));
                    }
                    this._szArgs.cancel = false;
                    if (g.onResizingColumn(this._szArgs)) { // TFS 144204
                        if (g.deferResizing || g.isTouching) {
                            this._showResizeMarker(sz);
                        } else {
                            (<Column>this._szRowCol).width = sz;
                        }
                    }
                }
            }

            // resizing row
            if (this._szRowCol instanceof Row) {
                var g = this._g,
                    pageY = e.clientY + pageYOffset, // e.pageXY doesn't work well in Chrome/zoom/touch
                    sz = Math.round(Math.max(_MouseHandler._SZ_MIN, this._szStart + (pageY - this._htDown.point.y)));
                if (this._szRowCol.renderSize != sz) {
                    if (this._szArgs == null) {
                        var panel = g.columnHeaders.rows.indexOf(this._szRowCol) > -1 ? g.columnHeaders : g.cells;
                        this._szArgs = new CellRangeEventArgs(panel, new CellRange(this._szRowCol.index, -1));
                    }
                    this._szArgs.cancel = false;
                    if (g.onResizingRow(this._szArgs)) { // TFS 144204
                        if (g.deferResizing || g.isTouching) {
                            this._showResizeMarker(sz);
                        } else {
                            (<Row>this._szRowCol).height = sz;
                        }
                    }
                }
            }
        }

        // drag-drop handling (dragging rows/columns)
        private _dragstart(e: DragEvent) {
            var g = this._g,
                ht = this._htDown;

            // make sure this is event is ours
            if (!ht) {
                return;
            }

            // get drag source element (if we're not resizing)
            this._dragSource = null;
            if (!this._szRowCol) {
                var args = new CellRangeEventArgs(g.cells, ht.range);
                if (ht.cellType == CellType.ColumnHeader && (g.allowDragging & AllowDragging.Columns) &&
                    ht.col > -1 && g.columns[ht.col].allowDragging) {
                    if (g.onDraggingColumn(args)) {
                        this._dragSource = e.target;
                    }
                } else if (ht.cellType == CellType.RowHeader && (g.allowDragging & AllowDragging.Rows) &&
                    ht.row > -1 && g.rows[ht.row].allowDragging) {
                    var row = g.rows[ht.row];
                    if (!(row instanceof GroupRow) && !(row instanceof _NewRowTemplate)) {
                        if (g.onDraggingRow(args)) {
                            this._dragSource = e.target;
                        }
                    }
                }
            }

            // if we have a valid source, set opacity; ow prevent dragging
            if (this._dragSource && e.dataTransfer) {
                e.dataTransfer.effectAllowed = 'move';
                e.dataTransfer.setData('text', ''); // required in FireFox (note: 'text/html' throws in IE!)
                this._dragSource.style.opacity = .5;
                e.stopPropagation(); // prevent parent grids from cancelling the event (TFS 120810)
                g.beginUpdate(); // suspend updates while dragging
            } else {
                e.preventDefault();
            }
        }
        private _dragend(e: DragEvent) {
            this._g.endUpdate(); // restore updates after dragging
            this.resetMouseState();
        }
        private _dragover(e: DragEvent) {
            var g = this._g,
                ht = g.hitTest(e),
                valid = false;

            // check whether the move is valid
            if (this._htDown && ht.cellType == this._htDown.cellType) {
                if (ht.cellType == CellType.ColumnHeader) {
                    valid = g.columns.canMoveElement(this._htDown.col, ht.col);
                } else if (ht.cellType == CellType.RowHeader) {
                    valid = g.rows.canMoveElement(this._htDown.row, ht.row);
                }
            }

            // if valid, prevent default to allow drop
            if (valid) {
                e.dataTransfer.dropEffect = 'move';
                e.preventDefault();
                this._showDragMarker(ht);
            } else {
                this._showDragMarker(null);
            }
        }
        private _drop(e: DragEvent) {
            var g = this._g,
                ht = g.hitTest(e),
                args = new CellRangeEventArgs(g.cells, ht.range);

            // move the row/col to a new position
            if (this._htDown && ht.cellType == this._htDown.cellType) {
                var sel = g.selection;
                if (ht.cellType == CellType.ColumnHeader) {
                    g.columns.moveElement(this._htDown.col, ht.col);
                    g.select(sel.row, ht.col);
                    g.onDraggedColumn(args);
                } else if (ht.cellType == CellType.RowHeader) {
                    g.rows.moveElement(this._htDown.row, ht.row);
                    g.select(ht.row, sel.col);
                    g.onDraggedRow(args);
                }
            }
            this.resetMouseState();
        }

        // updates the marker to show the new size of the row/col being resized
        private _showResizeMarker(sz: number) {
            var g = this._g;

            // add marker element to panel
            var t = this._dvMarker;
            if (!t.parentElement) {
                g.cells.hostElement.appendChild(t);
            }

            // update marker position
            var css: any,
                ct = this._szArgs.panel.cellType;
            if (this._szRowCol instanceof Column) {
                css = {
                    display: '',
                    left: this._szRowCol.pos + sz - 1,
                    top: 0,
                    right: '',
                    bottom: 0,
                    width: 3,
                    height: '',
                    zIndex: 1000
                }
                if (g._rtl) {
                    css.left = t.parentElement.clientWidth - css.left - css.width;
                }
                if (ct == CellType.TopLeft || ct == CellType.RowHeader) {
                    css.left -= g.topLeftCells.hostElement.offsetWidth;
                }
            } else { 
                css = {
                    left: 0,
                    top: this._szRowCol.pos + sz - 1,
                    right: 0,
                    bottom: '',
                    width: '',
                    height: 3,
                    zIndex: 1000
                }
                if (ct == CellType.TopLeft || ct == CellType.ColumnHeader) {
                    css.top -= g.topLeftCells.hostElement.offsetHeight;
                }
            }

            // apply new position
            setCss(t, css);
        }

        // updates the marker to show the position where the row/col will be inserted
        private _showDragMarker(ht: HitTestInfo) {
            var g = this._g;

            // remove target indicator if no HitTestInfo
            var t = this._dvMarker;
            if (!ht) {
                if (t.parentElement) {
                    t.parentElement.removeChild(t);
                }
                this._rngTarget = null;
                return;
            }

            // avoid work/flicker
            if (ht.range.equals(this._rngTarget)) {
                return;
            }
            this._rngTarget = ht.range;

            // add marker element to panel
            if (!t.parentElement) {
                ht.panel.hostElement.appendChild(t);
            }

            // update marker position
            var css: any = {
                display: '',
                left: 0,
                top: 0,
                width: 6,
                height: 6
            };
            switch (ht.cellType) {
                case CellType.ColumnHeader:
                    css.height = ht.panel.height;
                    var col = g.columns[ht.col];
                    css.left = col.pos - css.width / 2;
                    if (ht.col > this._htDown.col) {
                        css.left += col.renderWidth;
                    }
                    if (g._rtl) {
                        css.left = t.parentElement.clientWidth - css.left - css.width;
                    }
                    break;
                case CellType.RowHeader:
                    css.width = ht.panel.width;
                    var row = g.rows[ht.row];
                    css.top = row.pos - css.height / 2;
                    if (ht.row > this._htDown.row) {
                        css.top += row.renderHeight;
                    }
                    break;
            }

            // update marker
            setCss(t, css);
        }

        // raises the ResizedRow/Column events and 
        // applies the new size to the selection if the control key is pressed
        private _finishResizing(e: MouseEvent) {
            var g = this._g,
                sel = g.selection,
                ctrl = this._eMouse.ctrlKey,
                args = this._szArgs,
                pageX = e.clientX + pageXOffset, // e.pageXY doesn't work well in Chrome/zoom/touch
                pageY = e.clientY + pageYOffset,
                rc: number,
                sz;

            // finish column sizing
            if (args && !args.cancel && args.col > -1) { // TFS 144204

                // apply new size, fire event
                rc = args.col;
                sz = Math.round(Math.max(_MouseHandler._SZ_MIN, this._szStart + (pageX - this._htDown.point.x) * (this._g._rtl ? -1 : 1)));
                args.panel.columns[rc].width = Math.round(sz);
                g.onResizedColumn(args);

                // apply new size to selection if the control key is pressed
                if (ctrl && this._htDown.cellType == CellType.ColumnHeader && sel.containsColumn(rc)) {
                    for (var c = sel.leftCol; c <= sel.rightCol; c++) {
                        if (g.columns[c].allowResizing && c != rc) {
                            args = new CellRangeEventArgs(g.cells, new CellRange(-1, c));
                            if (g.onResizingColumn(args)) {
                                g.columns[c].size = g.columns[rc].size;
                                g.onResizedColumn(args);
                            }
                        }
                    }
                }
            }

            // finish row sizing
            if (args && !args.cancel && args.row > -1) { // TFS 144204

                // apply new size, fire event
                rc = args.row;
                sz = Math.round(Math.max(_MouseHandler._SZ_MIN, this._szStart + (pageY - this._htDown.point.y)));
                args.panel.rows[rc].height = Math.round(sz);
                g.onResizedRow(args);

                // apply new size to selection if the control key is pressed
                if (ctrl && this._htDown.cellType == CellType.RowHeader && sel.containsRow(rc)) {
                    for (var r = sel.topRow; r <= sel.bottomRow; r++) {
                        if (g.rows[r].allowResizing && r != rc) {
                            args = new CellRangeEventArgs(g.cells, new CellRange(r, -1));
                            if (g.onResizingRow(args)) { // TFS 144204
                                g.rows[r].size = g.rows[rc].size;
                                g.onResizedRow(args);
                            }
                        }
                    }
                }
            }
        }

        // start ListBox selection by keeping track of which rows were selected 
        // when the action started
        private _startListBoxSelection(row: number) {
            var rows = this._g.rows;
            this._lbSelState = !rows[row].isSelected;
            this._lbSelRows = {};
            for (var r = 0; r < rows.length; r++) {
                if (rows[r].isSelected) {
                    this._lbSelRows[r] = true;
                }
            }
        }

        // handle mouse selection
        private _handleSelection(ht: HitTestInfo, extend: boolean) {
            var g = this._g,
                rows = g.rows,
                sel = g._selHdl._sel,
                rng = new CellRange(ht.row, ht.col);

            // check that the selection is valid
            if (ht.row > -1 && ht.col > -1) {
                if (this._lbSelRows != null) {

                    // special handling for ListBox mode
                    var changed = false;
                    rng = new CellRange(ht.row, ht.col, this._htDown.row, this._htDown.col);
                    for (var r = 0; r < rows.length; r++) {
                        var selected = rng.containsRow(r) ? this._lbSelState : this._lbSelRows[r] != null;
                        if (selected != rows[r].isSelected) {
                            var e = new CellRangeEventArgs(g.cells, new CellRange(r, sel.col, r, sel.col2));
                            if (g.onSelectionChanging(e)) {
                                rows[r]._setFlag(RowColFlags.Selected, selected, true);
                                changed = true;
                                //rows[r].isSelected = selected; // this invalidates
                                g.onSelectionChanged(e);
                            }
                        }
                    }

                    // if the selection changed, refresh cells to show the change
                    if (changed) {
                        g.refreshCells(false, true, true);
                    }

                    // and scroll the selection into view
                    g.scrollIntoView(ht.row, ht.col);

                } else {

                    // row headers, select the whole row
                    if (ht.cellType == CellType.RowHeader) {
                        rng.col = 0;
                        rng.col2 = g.columns.length - 1;
                    }
                    
                    // extend range if that was asked
                    if (extend) {
                        rng.row2 = sel.row2;
                        rng.col2 = sel.col2;
                    }

                    // select
                    g.select(rng);
                }
            }
        }

        // handle mouse sort
        private _handleSort(e: MouseEvent) {
            var g = this._g,
                cv = g.collectionView,
                ht = g.hitTest(e);

            if (this._htDown && ht.cellType == this._htDown.cellType && ht.col == this._htDown.col && 
                ht.cellType == CellType.ColumnHeader && !ht.edgeRight && ht.col > -1 &&
                cv && cv.canSort && g.allowSorting) {

                // get row that was clicked accounting for merging
                var rng = g.getMergedRange(ht.panel, ht.row, ht.col),
                    row = rng ? rng.row2 : ht.row;
                
                // get column and binding column
                var col = g.columns[ht.col],
                    bcol = g._getBindingColumn(ht.panel, ht.row, col);

                // if the click was on the sort row, sort
                if (row == g._getSortRowIndex() || col != bcol) {
                    var currSort = bcol.currentSort,
                        asc = currSort != '+';
                    if (bcol.allowSorting && bcol.binding) {

                        // can't remove sort from unsorted column
                        if (!currSort && e.ctrlKey) return;

                        // raise sorting column
                        var args = new CellRangeEventArgs(g.columnHeaders, new CellRange(-1, ht.col));
                        if (g.onSortingColumn(args)) {

                            // update sort
                            var sds = cv.sortDescriptions;
                            if (e.ctrlKey) {
                                sds.clear();
                            } else {
                                sds.splice(0, sds.length, new collections.SortDescription(bcol._getBindingSort(), asc));
                            }

                            // raise sorted column
                            g.onSortedColumn(args);
                        }
                    }
                }
            }
        }
    }
}
module wijmo.grid {
    'use strict';

    /**
     * Handles the grid's editing.
     */
    export class _EditHandler {
        _g: FlexGrid;
        _rng: CellRange;
        _edt: HTMLInputElement;
        _lbx: input.ListBox;
        _htDown: HitTestInfo;
        _cancelClick: boolean;
        _fullEdit = false;
        _list = null;
        _evtInput: any;

        /**
         * Initializes a new instance of the @see:_EditHandler class.
         *
         * @param g @see:FlexGrid that owns this @see:_EditHandler.
         */
        constructor(g: FlexGrid) {
            this._g = g;

            // raise input event when selecting from ListBox
            this._evtInput = document.createEvent('HTMLEvents');
            this._evtInput.initEvent('input', true, false);

            // finish editing when selection changes (commit row edits if row changed)
            g.selectionChanging.addHandler((s, e: CellRangeEventArgs) => {

                // exit edit mode
                this.finishEditing();

                // commit any pending edits
                var oldrow = g._selHdl._sel.row;
                if (oldrow != e.row) {
                    var len = g.rows.length,
                        olditem = oldrow > -1 && oldrow < len ? g.rows[oldrow].dataItem : null,
                        newitem = e.row > -1 && e.row < len ? g.rows[e.row].dataItem : null;
                    if (olditem != newitem) {
                        this._commitRowEdits();
                    }
                }
            });

            // commit row edits when losing focus
            g.lostFocus.addHandler(() => {
                if (!this._g.containsFocus()) {
                    var ae = getActiveElement(); // TFS 121877, 122033 Bootstrap modal issue
                    if (!ae || getComputedStyle(ae).position != 'fixed') {
                        this._commitRowEdits();
                    }
                }
            });

            // commit edits when clicking non-cells (e.g. sort, drag, resize),
            // start editing when clicking on checkboxes
            g.addEventListener(g.hostElement, 'mousedown', (e) => {

                // start actions on left button only: TFS 114623
                if (e.defaultPrevented || e.button != 0) { 
                    return;
                }

                // not while resizing...
                if (g._mouseHdl._szRowCol) {
                    return;
                }

                // handle the event as usual
                var sel = g.selection,
                    ht = g.hitTest(e);
                this._htDown = null;
                this._cancelClick = false;
                if (ht.cellType != CellType.Cell && ht.cellType != CellType.None) {

                    // mouse down on non-cell area: commit any pending edits
                    // **REVIEW: this is a fix for TFS 98332
                    if (!this._lbx || !contains(this._lbx.hostElement, <HTMLElement>e.target)) {
                        this._commitRowEdits();
                    }

                } else if (ht.cellType != CellType.None) {

                    // start editing when clicking on checkboxes that are not the active editor
                    var edt = <HTMLInputElement>tryCast(e.target, HTMLInputElement);
                    if (edt && edt.type == 'checkbox' && hasClass(edt.parentElement, 'wj-cell')) {
                        if (edt != this.activeEditor) {

                            // start editing the item that was clicked
                            this.startEditing(false, ht.row, ht.col);

                            // toggle check after editing started
                            setTimeout(() => {
                                edt = this.activeEditor;
                                if (edt && edt.type == 'checkbox') {
                                    edt.checked = !edt.checked;
                                    edt.focus(); // TFS 135943
                                    this.finishEditing();
                                } else {
                                    this._cancelClick = true;
                                }
                            });
                        } else {
                            this.finishEditing();
                        }
                    }

                    // handle drop-down items (even on editors)
                    var icon = document.elementFromPoint(e.clientX, e.clientY);
                    if (closest(icon, '.' + CellFactory._WJC_DROPDOWN)) {
                        this._toggleListBox(ht.range);
                        this._htDown = null;
                        e.preventDefault();
                        return;
                    }

                    // if the click was on the cursor cell, save the hit test info
                    // to start editing when we get the click event later 
                    if (edt == null && ht.row == sel.row && ht.col == sel.col) {
                        this._htDown = ht;
                    }
                }
            }, true);

            // start editing when the user clicks the selected cell
            g.addEventListener(g.hostElement, 'click', (e) => {

                // prevent clicking on checkboxes when startEditing failed
                if (this._cancelClick) {
                    e.preventDefault();
                    e.stopPropagation();
                    return;
                }

                // start editing when clicking a cell without an active editor
                if (this._htDown && !this.activeEditor) {
                    var ht = g.hitTest(e);
                    if (ht.range.equals(this._htDown.range)) {
                        this.startEditing(true, ht.row, ht.col);
                    }
                }
            }, true);
        }
        /**
         * Starts editing a given cell.
         *
         * @param fullEdit Whether to stay in edit mode when the user presses the cursor keys. Defaults to false.
         * @param r Index of the row to be edited. Defaults to the currently selected row.
         * @param c Index of the column to be edited. Defaults to the currently selected column.
         * @param focus Whether to give the editor the focus. Defaults to true.
         * @return True if the edit operation started successfully.
         */
        startEditing(fullEdit = true, r?: number, c?: number, focus?: boolean): boolean {

            // default row/col to current selection
            var g = this._g;
            r = asNumber(r, true, true);
            c = asNumber(c, true, true);
            if (r == null) {
                r = g.selection.row;
            }
            if (c == null) {
                c = g.selection.col;
            }

            // default focus to true
            if (focus == null) {
                focus = true;
            }

            // check that the cell is editable
            if (!this._allowEditing(r, c)) {
                return false;
            }

            // get edit range
            var rng = g.getMergedRange(g.cells, r, c);
            if (!rng) {
                rng = new CellRange(r, c);
            }

            // get item to be edited
            var item = g.rows[r].dataItem;

            // make sure cell is selected
            g.select(rng, true);

            // check that we still have the same item after moving the selection (TFS 110143)
            if (!g.rows[r] || item != g.rows[r].dataItem) {
                return false;
            }

            // no work if we are already editing this cell
            if (rng.equals(this._rng)) {
                return true;
            }

            // start editing cell
            var e = new CellRangeEventArgs(g.cells, rng);
            if (!g.onBeginningEdit(e)) {
                return false;
            }

            // start editing item
            var ecv = tryCast(g.collectionView, 'IEditableCollectionView');
            if (ecv) {
                item = g.rows[r].dataItem;
                ecv.editItem(item);
            }

            // save editing parameters
            this._fullEdit = fullEdit;
            this._rng = rng;
            this._list = null;
            var map = <DataMap>g.columns[c].dataMap;
            if (map) {
                this._list = map.getDisplayValues(item);
            }

            // refresh to create and activate editor
            g.refresh(false);
            var edt = this._edt;
            if (edt) {
                if (edt.type == 'checkbox') {
                    this._fullEdit = false; // no full edit on checkboxes...
                } else if (focus) {
                    setSelectionRange(edt, 0, edt.value.length);
                }
                g.onPrepareCellForEdit(e);

                // give the editor the focus in case it doesn't have it
                // NOTE: this happens on Android, it's strange...
                edt = this._edt;
                if (edt && focus) {
                    edt.focus();
                }
            }

            // done
            return true;
        }
        /**
         * Commits any pending edits and exits edit mode.
         *
         * @param cancel Whether pending edits should be canceled or committed.
         * @return True if the edit operation finished successfully.
         */
        finishEditing(cancel = false): boolean {

            // make sure we're editing
            var edt = this._edt;
            if (!edt) {
                this._removeListBox();
                return true;
            }

            // get parameters
            var g = this._g,
                rng = this._rng,
                e = new CellRangeEventArgs(g.cells, rng),
                focus = this._g.containsFocus();

            // remove focus from editor
            // (to commit edits, important when using cell template editors)
            if (g.activeEditor && focus) {
                var ctl = wijmo.Control.getControl(closest(getActiveElement(), '.wj-control'));
                if (ctl && ctl != this._g) { // TFS 203106
                    ctl.onLostFocus(e);
                }
            }

            // edit ending
            e.cancel = cancel;
            g.onCellEditEnding(e);

            // apply edits
            if (!e.cancel) {
                var value: any = edt.type == 'checkbox' ? edt.checked : edt.value;
                for (var r = rng.topRow; r <= rng.bottomRow && r < g.rows.length; r++) {
                    for (var c = rng.leftCol; c <= rng.rightCol && c < g.columns.length; c++) {
                        g.cells.setCellData(r, c, value, true);
                    }
                }
            }

            // dispose of editor
            this._edt = null;
            this._rng = null;
            this._list = null;
            this._removeListBox();

            // refresh to replace the editor with regular content
            g.refresh(false);

            // restore focus
            if (focus) {
                g.focus();
            }

            // edit ended
            g.onCellEditEnded(e);

            // done
            return true;
        }
        /**
         * Gets the <b>HTMLInputElement</b> that represents the cell editor currently active.
         */
        get activeEditor(): HTMLInputElement {
            return this._edt;
        }
        /**
         * Gets a @see:CellRange that identifies the cell currently being edited.
         */
        get editRange(): CellRange {
            return this._rng;
        }

        // ** implementation

        // checks whether a cell can be edited
        _allowEditing(r: number, c: number) {
            var g = this._g;
            if (g.isReadOnly || g.selectionMode == SelectionMode.None) return false;
            if (r < 0 || r >= g.rows.length || g.rows[r].isReadOnly || !g.rows[r].isVisible) return false;
            if (c < 0 || c >= g.columns.length || g.columns[c].isReadOnly || !g.columns[c].isVisible) return false;
            if (g._getBindingColumn(g.cells, r, g.columns[c]).isReadOnly) return false;
            return true;
        }

        // finish editing the current item
        _commitRowEdits() {
            this.finishEditing();
            var g = this._g,
                ecv = tryCast(g.collectionView, 'IEditableCollectionView');
            if (ecv) {
                if (ecv.currentEditItem) { // || ecv.currentAddItem) { // TFS: 206038
                    var e = new CellRangeEventArgs(g.cells, g.selection);
                    g.onRowEditEnding(e);
                    ecv.commitEdit();
                    g.onRowEditEnded(e);
                }
            }
        }

        // handles keyDown events while editing
        // returns true if the key was handled, false if the grid should handle it
        _keydown(e: KeyboardEvent): boolean {
            switch (e.keyCode) {

                // F2 toggles edit mode
                case Key.F2:
                    this._fullEdit = !this._fullEdit;
                    e.preventDefault();
                    return true;

                // F4 toggles ListBox
                case Key.F4:
                    this._toggleListBox(this._g.selection);
                    e.preventDefault();
                    return true;

                // space toggles checkboxes
                case Key.Space:
                    var edt = this._edt;
                    if (edt && edt.type == 'checkbox') {
                        edt.checked = !edt.checked;
                        this.finishEditing();
                        e.preventDefault();
                    }
                    return true;

                // enter, tab, escape finish editing
                case Key.Enter:
                case Key.Tab:
                    this.finishEditing();
                    return false;
                case Key.Escape:
                    this.finishEditing(true);
                    return true;

                // cursor keys: ListBox selection/finish editing if not in full edit mode
                case Key.Up:
                case Key.Down:
                case Key.Left:
                case Key.Right:
                case Key.PageUp:
                case Key.PageDown:
                case Key.Home:
                case Key.End:

                    // if the ListBox is active, let it handle the key
                    if (this._lbx) {
                        return this._keydownListBox(e);
                    }

                    // open ListBox on alt up/down
                    if (e.altKey) {
                        switch (e.keyCode) {
                            case Key.Up:
                            case Key.Down:
                                this._toggleListBox(this._g.selection);
                                e.preventDefault();
                                return true;
                        }
                    }

                    // finish editing if not in full-edit mode
                    if (!this._fullEdit) {
                        this.finishEditing();
                        return false;
                    }
            }

            // return true to let editor handle the key (not the grid)
            return true;
        }

        // handles keydown events when ListBox is visible
        _keydownListBox(e: KeyboardEvent) {
            var handled = true;
            if (this._lbx) {
                switch (e.keyCode) {
                    case Key.Up:
                        if (e.altKey) {
                            this._toggleListBox(this._g.selection);
                        } else if (this._lbx.selectedIndex > 0) {
                            this._lbx.selectedIndex--;
                        }
                        break;
                    case Key.Down:
                        if (e.altKey) {
                            this._toggleListBox(this._g.selection);
                        } else {
                            this._lbx.selectedIndex++;
                        }
                        break;
                    case Key.Home:
                    case Key.PageUp:
                        this._lbx.selectedIndex = 0;
                        break;
                    case Key.End:
                    case Key.PageDown:
                        this._lbx.selectedIndex = this._lbx.collectionView.items.length - 1;
                        break;
                    default:
                        handled = false;
                        break;
                }
            }

            // if handled, we're done
            if (handled) {
                e.preventDefault();
                return true;
            }

            // return false to let the grid handle the key
            return false;
        }

        // handles keyPress events while editing
        _keypress(e: KeyboardEvent) {

            // auto-complete based on dataMap
            var edt = this._edt;
            if (edt && edt.type != 'checkbox' && e.target == edt &&
                this._list && this._list.length > 0 && e.charCode >= 32) {

                // get text up to selection start
                var start = edt.selectionStart,
                    text = edt.value.substr(0, start);

                // add the new char if the source element is the editor
                // (but not if the source element is the grid!)
                if (e.target == edt) {
                    start++;
                    text += String.fromCharCode(e.charCode);
                }

                // convert to lower-case for matching
                text = text.toLowerCase();

                // look for a match
                for (var i = 0; i < this._list.length; i++) {
                    if (this._list[i].toLowerCase().indexOf(text) == 0) {

                        // found the match, update text and selection
                        edt.value = this._list[i];
                        setSelectionRange(edt, start, this._list[i].length);
                        edt.dispatchEvent(this._evtInput);

                        // eat the key and be done
                        e.preventDefault();
                        break;
                    }
                }
            }
        }

        // shows the drop-down element for a cell (if it is not already visible)
        _toggleListBox(rng: CellRange): boolean {
            var g = this._g;

            // close select element if any; if this is the same cell, we're done
            if (this._lbx) {
                this._removeListBox();
                if (g.selection.contains(rng)) {
                    if (g.activeEditor) {
                        g.activeEditor.focus();
                    } else if (!g.containsFocus()) {
                        g.focus();
                    }
                    return true;
                }
            }

            // if this was a touch, give focus to ListBox to hide soft keyboard
            var lbxFocus = g.isTouching;

            // check that we have a drop-down
            var bcol = g._getBindingColumn(g.cells, rng.row, g.columns[rng.col]);
            if (!wijmo.input || !bcol.dataMap || bcol.showDropDown === false) {
                return false;
            }

            // start editing so we can position the select element
            if (!wijmo.input || !this.startEditing(true, rng.row, rng.col, !lbxFocus)) {
                return false;
            }

            // create and initialize the ListBox
            this._lbx = this._createListBox();
            this._lbx.showSelection();
            if (lbxFocus) {
                this._lbx.focus();
            }
            return true;
        }

        // create the ListBox and add it to the document
        private _createListBox(): input.ListBox {
            var g = this._g,
                rng = this._rng,
                row = g.rows[rng.row],
                col = g._getBindingColumn(g.cells, rng.row, g.columns[rng.col]),
                div = document.createElement('div'),
                lbx = new input.ListBox(div);

            // configure ListBox
            addClass(div, 'wj-dropdown-panel');
            lbx.maxHeight = row.renderHeight * 4;
            lbx.itemsSource = col.dataMap.getDisplayValues(row.dataItem);
            lbx.selectedValue = g.activeEditor
                ? g.activeEditor.value
                : g.getCellData(rng.row, rng.col, true);
            addClass(div, col.dropDownCssClass);

            // close ListBox on clicks
            lbx.addEventListener(lbx.hostElement, 'click', () => {
                this._removeListBox();
                this.finishEditing();
            });

            // close ListBox when losing focus
            lbx.lostFocus.addHandler(() => {
                this._removeListBox();
            });

            // update editor when the selected index changes
            lbx.selectedIndexChanged.addHandler(() => {
                var edt = g.activeEditor;
                if (edt) {
                    edt.value = this._lbx.selectedValue;
                    edt.dispatchEvent(this._evtInput);
                    setSelectionRange(edt, 0, edt.value.length);
                }
            });

            // show the popup
            showPopup(div, g.getCellBoundingRect(rng.row, rng.col));

            // done
            return lbx;
        }

        // remove the ListBox element from the DOM and disconnect its event handlers
        private _removeListBox() {
            if (this._lbx) {
                hidePopup(this._lbx.hostElement, true);
                this._lbx.dispose();
                this._lbx = null;
            }
        }
    }
}
module wijmo.grid {
    'use strict';

    /**
     * Manages the new row template used to add rows to the grid.
     */
    export class _AddNewHandler {
        protected _g: FlexGrid;
        protected _nrt = new _NewRowTemplate();
        protected _keydownBnd: any;
        protected _top: boolean;

        /**
         * Initializes a new instance of the @see:_AddNewHandler class.
         *
         * @param g @see:FlexGrid that owns this @see:_AddNewHandler.
         */
        constructor(g: FlexGrid) {
            this._g = g;
            this._keydownBnd = this._keydown.bind(this);
            this._attach();
        }

        /**
         * Gets or sets a value that indicates whether the new row template should be located
         * at the top of the grid or at the bottom.
         */
        get newRowAtTop(): boolean {
            return this._top;
        }
        set newRowAtTop(value: boolean) {
            if (value != this.newRowAtTop) {
                this._top = asBoolean(value);
                this.updateNewRowTemplate();
            }
        }
        /**
         * Updates the new row template to ensure it's visible only if the grid is
         * bound to a data source that supports adding new items, and that it is 
         * in the right position.
         */
        updateNewRowTemplate() {

            // get variables
            var ecv = <collections.IEditableCollectionView>tryCast(this._g.collectionView, 'IEditableCollectionView'),
                g = this._g,
                rows = g.rows;

            // see if we need a new row template
            var needTemplate = ecv && ecv.canAddNew && g.allowAddNew && !g.isReadOnly;

            // get current template index
            var index = rows.indexOf(this._nrt),
                newRowPos = this._top ? 0 : rows.length - 1,
                insert = false;

            // update template position
            if (!needTemplate && index > -1) { // not needed but present, remove it
                var sel = g.selection; // move selection away from the row being deleted
                if (sel.row == index) {
                    g.select(sel.row - 1, sel.col);
                }
                rows.removeAt(index);
            } else if (needTemplate) {
                if (index < 0) { // needed but not present, add it now
                    insert = true;
                } else if (index != newRowPos) { // position template
                    rows.removeAt(index);
                    insert = true;
                }

                // add the new row template at the proper position
                if (insert) {
                    if (this._top) {
                        rows.insert(0, this._nrt);
                    } else {
                        rows.push(this._nrt);
                    }
                }

                // make sure the new row template is not collapsed
                if (this._nrt) {
                    this._nrt._setFlag(RowColFlags.ParentCollapsed, false);
                }
            }
        }

        // ** implementation

        // add/remove handlers to manage the new row template
        /*protected*/ _attach() {
            var g = this._g;
            if (g) {
                g.beginningEdit.addHandler(this._beginningEdit, this);
                g.pastingCell.addHandler(this._beginningEdit, this);
                g.rowEditEnded.addHandler(this._rowEditEnded, this);
                g.loadedRows.addHandler(this.updateNewRowTemplate, this);
                g.hostElement.addEventListener('keydown', this._keydownBnd, true);
            }
        }
        /*protected*/ _detach() {
            var g = this._g;
            if (g) {
                g.beginningEdit.removeHandler(this._beginningEdit);
                g.pastingCell.removeHandler(this._beginningEdit);
                g.rowEditEnded.removeHandler(this._rowEditEnded);
                g.loadedRows.removeHandler(this.updateNewRowTemplate);
                g.hostElement.removeEventListener('keydown', this._keydownBnd, true);
            }
        }

        // cancel new row at top addition on Escape (same as new row at bottom)
        /*protected*/ _keydown(e: KeyboardEvent) {
            if (!e.defaultPrevented && e.keyCode == Key.Escape) {
                if (this._g.activeEditor == null && this._top && this._nrt.dataItem) {
                    this._nrt.dataItem = null;
                    this._g.invalidate();
                }
            }
        }

        // beginning edit, add new item if necessary
        /*protected*/ _beginningEdit(sender, e: CellRangeEventArgs) {
            if (!e.cancel) {
                var row = this._g.rows[e.row];
                if (tryCast(row, _NewRowTemplate)) {
                    var ecv = <collections.IEditableCollectionView>tryCast(this._g.collectionView, 'IEditableCollectionView');
                    if (ecv && ecv.canAddNew) {

                        // add new row at the top
                        if (this._top) {
                            if (this._nrt.dataItem == null) {

                                // create new item
                                var newItem = null,
                                    src = ecv.sourceCollection,
                                    creator = ecv['newItemCreator'];
                                if (wijmo.isFunction(creator)) {
                                    newItem = creator();
                                } else if (src && src.length) {
                                    newItem = new src[0].constructor();
                                } else {
                                    newItem = {};
                                }

                                // assign new item to new row template
                                this._nrt.dataItem = newItem;
                            }

                        // add new row at the bottom (TFS 145498)
                        } else {
                            var newItem = (ecv.currentAddItem && ecv.currentAddItem == row.dataItem) ? ecv.currentAddItem : ecv.addNew();
                            ecv.moveCurrentTo(newItem);
                            this.updateNewRowTemplate();

                            // update now to ensure the editor will get a fresh layout (TFS 96705)
                            this._g.refresh(true);

                            // fire row added event (user can customize the new row or cancel)
                            this._g.onRowAdded(e);
                            if (e.cancel) {
                                ecv.cancelNew();
                            }
                        }
                    }
                }
            }
        }

        // row has been edited, commit if this is the new row
        /*protected*/ _rowEditEnded(sender, e: CellRangeEventArgs) {
            var ecv = <collections.IEditableCollectionView>tryCast(this._g.collectionView, 'IEditableCollectionView'),
                item = this._nrt.dataItem;
            if (ecv) {

                // adding at the bottom
                if (ecv.isAddingNew) {
                    ecv.commitNew();

                // adding at the top
                } else if (item && !e.cancel) {

                    // clear row template data
                    this._nrt.dataItem = null;

                    // add new item to collection view
                    var newItem = ecv.addNew();
                    for (var k in item) {
                        newItem[k] = item[k];
                    }
                    ecv.commitNew();

                    // move selection back to new row template
                    setTimeout(() => {
                        this._g.select(0, this._g.columns.firstVisibleIndex);
                    }, 20);
                }
            }
        }
    }

    /**
     * Represents a row template used to add items to the source collection.
     */
    export class _NewRowTemplate extends Row {
    }
}
module wijmo.grid {
    'use strict';

    /**
     * Implements a hidden input element so users can choose IME modes when 
     * the FlexGrid has focus, and start composing before the grid enters
     * edit mode.
     */
    export class _ImeHandler {
        _g: FlexGrid;
        _tbx: HTMLInputElement;
        _cssHidden: any;
        _mouseDown: boolean;

        //--------------------------------------------------------------------------
        //#region ** ctor

        /**
         * Initializes a new instance of the @see:_ImeHandler class and attaches it to a @see:FlexGrid.
         * 
         * @param g @see:FlexGrid that this @see:_ImeHandler will be attached to.
         */
        constructor(g: FlexGrid) {

            // create hidden input focus element
            this._tbx = <HTMLInputElement>createElement('<input class="wj-grid-editor wj-form-control" wj-part="ime-target"/>');
            this._cssHidden = {
                opacity: '0',
                pointerEvents: 'none',
                position: 'absolute',
                left: -10,
                top: -10,
                width: 0
            };
            setCss(this._tbx, this._cssHidden);

            // add IME input to the grid, update the focus
            this._g = g,
                this._g.cells.hostElement.parentElement.appendChild(this._tbx);
            this._updateImeFocus();

            // attach event handlers
            var g = this._g,
                host = g.hostElement;
            g.addEventListener(this._tbx, 'compositionstart', this._compositionstart.bind(this));
            g.addEventListener(this._tbx, 'keydown', this._keydown.bind(this));
            g.addEventListener(host, 'blur', this._updateImeFocus.bind(this), true);
            g.addEventListener(host, 'focus', this._updateImeFocus.bind(this), true);
            g.addEventListener(host, 'mousedown', this._mousedown.bind(this), true);
            g.addEventListener(host, 'mouseup', this._mouseup.bind(this), true);
            g.cellEditEnded.addHandler(this._cellEditEnded, this);
            g.selectionChanged.addHandler(this._updateImeFocus, this);
        }

        //#endregion
        //--------------------------------------------------------------------------
        //#region ** object model

        /**
         * Disposes of this @see:_ImeHandler.
         */
        dispose() {
            var g = this._g,
                host = g.hostElement;

            // remove event listeners
            g.removeEventListener(this._tbx, 'compositionstart');
            g.removeEventListener(host, 'blur');
            g.removeEventListener(host, 'focus');
            g.removeEventListener(host, 'mousedown');
            g.removeEventListener(host, 'mouseup');
            g.cellEditEnded.removeHandler(this._cellEditEnded);
            g.selectionChanged.removeHandler(this._updateImeFocus);

            // remove IME input from grid
            if (this._tbx.parentElement) {
                this._tbx.parentElement.removeChild(this._tbx);
            }
        }

        //#endregion
        //--------------------------------------------------------------------------
        //#region ** implementation
        
        // hide IME input after editing
        _cellEditEnded() {
            setCss(this._tbx, this._cssHidden);
            this._tbx.value = '';
        }

        // show IME input as current editor when composition starts
        _compositionstart() {
            var g = this._g;
            if (g.activeEditor == null) {
                var sel = g._selHdl.selection;
                if (g.startEditing(false, sel.row, sel.col, false)) {
                    var rc = g.getCellBoundingRect(sel.row, sel.col),
                        host = g.cells.hostElement;
                    setCss(this._tbx, {
                        opacity: '',
                        pointerEvents: '',
                        left: g.columns[sel.col].pos + host.offsetLeft,
                        top: g.rows[sel.row].pos + host.offsetTop,
                        width: rc.width - 1,
                        height: rc.height - 1,
                    });
                    g._edtHdl._edt = this._tbx;
                }
            }
        }

        // forward up/down keys to grid in case it's in a form element in IE (ugh! TFS 202913)
        _keydown(e: KeyboardEvent) {
            switch (e.keyCode) {
                case Key.Up:
                case Key.Down:
                case Key.PageUp:
                case Key.PageDown:
                    this._g._keyHdl._keydown(e);
            }
        }

        // enable/disable IME on mousedown/up (TFS 194411)
        _mousedown(e) {
            this._mouseDown = true;
            this._updateImeFocus();
        }
        _mouseup(e) {
            this._mouseDown = false;
            this._updateImeFocus();
        }

        // transfer focus from grid to IME input
        _updateImeFocus() {
            var g = this._g;
            if (g.containsFocus() && !g.activeEditor && !g.isTouching && !this._mouseDown) {
                var tbx = this._tbx;
                if (this._enableIme()) {
                    tbx.disabled = false;
                    tbx.select();
                } else if (!tbx.disabled) {
                    tbx.disabled = true;
                    var focused = getActiveElement();
                    if (focused instanceof HTMLElement) {
                        focused.blur();
                    }
                    g.focus();
                }
            }
        }

        // checks whether IME should be enabled for the current selection
        _enableIme(): boolean {
            var g = this._g,
                sel = g.selection;

            // can't edit? can't use IME
            if (sel.row < 0 || sel.col < 0 || !g._edtHdl._allowEditing(sel.row, sel.col)) {
                return false;
            }

            // disable IME for boolean cells (with checkboxes)
            if (g.columns[sel.col].dataType == DataType.Boolean) {
                return false;
            }

            // seems OK to use IME
            return true;
        }

        //#endregion
    }
}
