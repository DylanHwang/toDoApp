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
/**
 * Defines the @see:Sunburst chart control and its associated classes.
 */
module wijmo.chart.hierarchical {
    'use strict';

    /**
     * Sunburst chart control.
     */
    export class Sunburst extends FlexPie {

        //conflicts with _bindingName in FlexPie if use _bindingName, so use _bindName instead;
        private _bindName: any;
        private _processedData: any[] = [];
        private _legendLabels: string[] = [];
        private _level: number = 1;
        private _sliceIndex: number = 0;
        private _childItemsPath: any;

        constructor(element: any, options?) {
            super(element, options);

            //add classes to host element
            this._selectionIndex = 0;
            this.applyTemplate('wj-sunburst', null, null);
        }

 /**
         * Gets or sets the name of the property containing name of the data item;
         * it should be an array or a string.
         */
        get bindingName(): any {
            return this._bindName;
        }
        set bindingName(value: any) {
            if (value != this._bindName) {
                assert(value == null || isArray(value) || isString(value), 'bindingName should be an array or a string.');
                this._bindName = value;
                this._bindChart();
            }
        }

        /**
         * Gets or sets the name of the property (or properties) used to generate 
         * child items in hierarchical data.
         *
         * Set this property to a string to specify the name of the property that
         * contains an item's child items (e.g. <code>'items'</code>). 
         * 
         * Set this property to an array containing the names of the properties
         * that contain child items at each level, when the items are child items
         * at different levels with different names 
         * (e.g. <code>[ 'accounts', 'checks', 'earnings' ]</code>).
         */
        get childItemsPath(): any {
            return this._childItemsPath;
        }
        set childItemsPath(value: any) {
            if (value != this._childItemsPath) {
                assert(value == null || isArray(value) || isString(value), 'childItemsPath should be an array or a string.');
                this._childItemsPath = value;
                this._bindChart();
            }
        }

        _initData() {
            super._initData();
            this._processedData = [];
            this._level = 1;
            this._legendLabels = [];
        }

        _performBind() {
            var items, processedData;

            this._initData();

            if (this._cv) {
                //this._selectionIndex = this._cv.currentPosition;
                items = this._cv.items;
                if (items) {
                    this._processedData = HierarchicalUtil.parseDataToHierarchical(items, this.binding, this.bindingName, this.childItemsPath);
                    this._sum = this._calculateValueAndLevel(this._processedData, 1);
                    this._processedData.forEach(v => {
                        this._legendLabels.push(v.name);
                    });
                }
            }
        }

        private _calculateValueAndLevel(arr, level) {
            var sum = 0,
                values = this._values,
                labels = this._labels;

            if (this._level < level) {
                this._level = level;
            }
            arr.forEach(v => {
                var val;
                if (v.items) {
                    val = this._calculateValueAndLevel(v.items, level + 1);
                    v.value = val;
                    values.push(val);
                    labels.push(v.name);
                } else {
                    val = this._getBindData(v, values, labels, 'value', 'name');
                    v.value = val;
                }
                sum += val;
            });
            return sum;
        }

        _renderPie(engine, radius, innerRadius, startAngle, offset) {
            var center = this._getCenter();

            this._sliceIndex = 0;
            this._renderHierarchicalSlices(engine, center.x, center.y, this._processedData, this._sum, radius, innerRadius, startAngle, 2 * Math.PI, offset, 1);
        }

        _renderHierarchicalSlices(engine, cx, cy, values, sum, radius, innerRadius, startAngle, totalSweep, offset, level) {
            var len = values.length,
                angle = startAngle,
                reversed = this.reversed == true,
                r, ir, segment, sweep, value, val, pel, x, y, currentAngle;

            segment = (radius - innerRadius) / this._level;
            r = radius - (this._level - level) * segment;
            ir = innerRadius + (level - 1) * segment;
            for (var i = 0; i < len; i++) {
                x = cx;
                y = cy;
                pel = engine.startGroup('slice-level' + level);
                if (level === 1) {
                    engine.fill = this._getColorLight(i);
                    engine.stroke = this._getColor(i);
                }

                value = values[i];
                val = Math.abs(value.value);
                sweep = Math.abs(val - sum) < 1E-10 ? totalSweep : totalSweep * val / sum;
                currentAngle = reversed ? angle - 0.5 * sweep : angle + 0.5 * sweep;
                if (offset > 0 && sweep < totalSweep) {
                    x += offset * Math.cos(currentAngle);
                    y += offset * Math.sin(currentAngle);
                }

                if (value.items) {
                    this._renderHierarchicalSlices(engine, x, y, value.items, val, radius, innerRadius, angle, sweep, 0, level + 1);
                }
                this._renderSlice(engine, x, y, currentAngle, this._sliceIndex, r, ir, angle, sweep, totalSweep);
                this._sliceIndex++;

                if (reversed) {
                    angle -= sweep;
                } else {
                    angle += sweep;
                }

                engine.endGroup();
                this._pels.push(pel);
            }
        }

        _getLabelsForLegend() {
            return this._legendLabels || [];
        }

        _highlightCurrent() {
            if (this.selectionMode != SelectionMode.None) {
                this._highlight(true, this._selectionIndex);
            }
        }

    }
}
//
// Contains utilities used by hierarchical chart.
//
module wijmo.chart.hierarchical {
    'use strict';

    export class HierarchicalUtil {
        static parseDataToHierarchical(data, binding, bindingName, childItemsPath): any[] {
            var arr = [],
                items;
            
            if (data.length > 0) {
                if (wijmo.isString(bindingName) && bindingName.indexOf(',') > -1) {
                    bindingName = bindingName.split(',');
                }
                if (childItemsPath) {
                    arr = HierarchicalUtil.parseItems(data, binding, bindingName, childItemsPath);
                } else {
                    //flat data
                    items = HierarchicalUtil.ConvertFlatData(data, binding, bindingName);
                    arr = HierarchicalUtil.parseItems(items, 'value', bindingName, 'items');
                }
            }
            return arr;
        }

        private static parseItems(items, binding, bindingName, childItemsPath): any[] {
            var arr = [], i,
                len = items.length;

            for (i = 0; i < len; i++) {
                arr.push(HierarchicalUtil.parseItem(items[i], binding, bindingName, childItemsPath));
            }
            return arr;
        }

        private static isFlatItem(item, binding) {
            if (wijmo.isArray(item[binding])) {
                return false;
            }
            return true;
        }

        private static ConvertFlatData(items, binding, bindingName): any[] {
            var arr = [],
                data: any = {},
                i, item,
                len = items.length;

            for (i = 0; i < len; i++) {
                item = items[i];
                HierarchicalUtil.ConvertFlatItem(data, item, binding, bindingName);
            }
            HierarchicalUtil.ConvertFlatToHierarchical(arr, data);

            return arr;
        }

        private static ConvertFlatToHierarchical(arr, data) {
            var order = data['flatDataOrder'];

            if (order) {
                order.forEach(v => {
                    var d: any = {},
                        val = data[v],
                        items;

                    d[data['field']] = v;
                    if (val['flatDataOrder']) {
                        items = [];
                        HierarchicalUtil.ConvertFlatToHierarchical(items, val);
                        d.items = items;
                    } else {
                        d.value = val;
                    }
                    arr.push(d);
                });
            }

        }

        private static ConvertFlatItem(data, item, binding, bindingName): boolean {
            var newBindingName, name, len, itemName, newData, converted;

            newBindingName = bindingName.slice();
            name = newBindingName.shift().trim();
            itemName = item[name];

            if (itemName == null) {
                return false;
            }
            if (newBindingName.length === 0) {
                data[itemName] = item[binding];
                if (data['flatDataOrder']) {
                    data['flatDataOrder'].push(itemName);
                } else {
                    data['flatDataOrder'] = [itemName];
                }
                data['field'] = name;
            } else {
                if (data[itemName] == null) {
                    data[itemName] = {};
                    if (data['flatDataOrder']) {
                        data['flatDataOrder'].push(itemName);
                    } else {
                        data['flatDataOrder'] = [itemName];
                    }
                    data['field'] = name;
                }
                newData = data[itemName];
                converted = HierarchicalUtil.ConvertFlatItem(newData, item, binding, newBindingName);
                if (!converted) {
                    data[itemName] = item[binding];
                }
            }
            return true;
        }

        private static parseItem(item, binding, bindingName, childItemsPath) {
            var data: any = {},
                newBindingName, name, value, len, childItem, newChildItemsPath;

            if (wijmo.isArray(childItemsPath)) {
                newChildItemsPath = childItemsPath.slice();
                childItem = newChildItemsPath.length ? newChildItemsPath.shift().trim() : '';
            } else {
                newChildItemsPath = childItemsPath;
                childItem = childItemsPath;
            }
            if (wijmo.isArray(bindingName)) {
                newBindingName = bindingName.slice();
                name = newBindingName.shift().trim();

                data.nameField = name;
                data.name = item[name];
                value = item[childItem];
                if (newBindingName.length === 0) {
                    data.value = item[binding];
                } else {
                    if (value && wijmo.isArray(value)) {
                        data.items = HierarchicalUtil.parseItems(value, binding, newBindingName, newChildItemsPath);
                    } else {
                        data.value = item[binding];
                    }
                }
            } else {
                data.nameField = bindingName;
                data.name = item[bindingName];
                value = item[childItem];
                if (value != null && wijmo.isArray(value)) {
                    data.items = HierarchicalUtil.parseItems(value, binding, bindingName, newChildItemsPath);
                } else {
                    data.value = item[binding];
                }
            }

            return data;
        }

        static parseFlatItem(data, item, binding, bindingName) {
            if (!data.items) {
                data.items = [];
            }
        }
    }
}
