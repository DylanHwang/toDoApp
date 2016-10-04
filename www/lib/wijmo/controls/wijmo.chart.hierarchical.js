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
* Defines the @see:Sunburst chart control and its associated classes.
*/
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart) {
        var hierarchical;
        (function (hierarchical) {
            'use strict';
            /**
             * Sunburst chart control.
             */
            var Sunburst = (function (_super) {
                __extends(Sunburst, _super);
                function Sunburst(element, options) {
                    _super.call(this, element, options);
                    this._processedData = [];
                    this._legendLabels = [];
                    this._level = 1;
                    this._sliceIndex = 0;
                    //add classes to host element
                    this._selectionIndex = 0;
                    this.applyTemplate('wj-sunburst', null, null);
                }
                Object.defineProperty(Sunburst.prototype, "bindingName", {
                    /**
                            * Gets or sets the name of the property containing name of the data item;
                            * it should be an array or a string.
                            */
                    get: function () {
                        return this._bindName;
                    },
                    set: function (value) {
                        if (value != this._bindName) {
                            wijmo.assert(value == null || wijmo.isArray(value) || wijmo.isString(value), 'bindingName should be an array or a string.');
                            this._bindName = value;
                            this._bindChart();
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(Sunburst.prototype, "childItemsPath", {
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
                    get: function () {
                        return this._childItemsPath;
                    },
                    set: function (value) {
                        if (value != this._childItemsPath) {
                            wijmo.assert(value == null || wijmo.isArray(value) || wijmo.isString(value), 'childItemsPath should be an array or a string.');
                            this._childItemsPath = value;
                            this._bindChart();
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                Sunburst.prototype._initData = function () {
                    _super.prototype._initData.call(this);
                    this._processedData = [];
                    this._level = 1;
                    this._legendLabels = [];
                };
                Sunburst.prototype._performBind = function () {
                    var _this = this;
                    var items, processedData;
                    this._initData();
                    if (this._cv) {
                        //this._selectionIndex = this._cv.currentPosition;
                        items = this._cv.items;
                        if (items) {
                            this._processedData = hierarchical.HierarchicalUtil.parseDataToHierarchical(items, this.binding, this.bindingName, this.childItemsPath);
                            this._sum = this._calculateValueAndLevel(this._processedData, 1);
                            this._processedData.forEach(function (v) {
                                _this._legendLabels.push(v.name);
                            });
                        }
                    }
                };
                Sunburst.prototype._calculateValueAndLevel = function (arr, level) {
                    var _this = this;
                    var sum = 0, values = this._values, labels = this._labels;
                    if (this._level < level) {
                        this._level = level;
                    }
                    arr.forEach(function (v) {
                        var val;
                        if (v.items) {
                            val = _this._calculateValueAndLevel(v.items, level + 1);
                            v.value = val;
                            values.push(val);
                            labels.push(v.name);
                        }
                        else {
                            val = _this._getBindData(v, values, labels, 'value', 'name');
                            v.value = val;
                        }
                        sum += val;
                    });
                    return sum;
                };
                Sunburst.prototype._renderPie = function (engine, radius, innerRadius, startAngle, offset) {
                    var center = this._getCenter();
                    this._sliceIndex = 0;
                    this._renderHierarchicalSlices(engine, center.x, center.y, this._processedData, this._sum, radius, innerRadius, startAngle, 2 * Math.PI, offset, 1);
                };
                Sunburst.prototype._renderHierarchicalSlices = function (engine, cx, cy, values, sum, radius, innerRadius, startAngle, totalSweep, offset, level) {
                    var len = values.length, angle = startAngle, reversed = this.reversed == true, r, ir, segment, sweep, value, val, pel, x, y, currentAngle;
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
                        }
                        else {
                            angle += sweep;
                        }
                        engine.endGroup();
                        this._pels.push(pel);
                    }
                };
                Sunburst.prototype._getLabelsForLegend = function () {
                    return this._legendLabels || [];
                };
                Sunburst.prototype._highlightCurrent = function () {
                    if (this.selectionMode != chart.SelectionMode.None) {
                        this._highlight(true, this._selectionIndex);
                    }
                };
                return Sunburst;
            }(chart.FlexPie));
            hierarchical.Sunburst = Sunburst;
        })(hierarchical = chart.hierarchical || (chart.hierarchical = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Sunburst.js.map
//
// Contains utilities used by hierarchical chart.
//
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart) {
        var hierarchical;
        (function (hierarchical) {
            'use strict';
            var HierarchicalUtil = (function () {
                function HierarchicalUtil() {
                }
                HierarchicalUtil.parseDataToHierarchical = function (data, binding, bindingName, childItemsPath) {
                    var arr = [], items;
                    if (data.length > 0) {
                        if (wijmo.isString(bindingName) && bindingName.indexOf(',') > -1) {
                            bindingName = bindingName.split(',');
                        }
                        if (childItemsPath) {
                            arr = HierarchicalUtil.parseItems(data, binding, bindingName, childItemsPath);
                        }
                        else {
                            //flat data
                            items = HierarchicalUtil.ConvertFlatData(data, binding, bindingName);
                            arr = HierarchicalUtil.parseItems(items, 'value', bindingName, 'items');
                        }
                    }
                    return arr;
                };
                HierarchicalUtil.parseItems = function (items, binding, bindingName, childItemsPath) {
                    var arr = [], i, len = items.length;
                    for (i = 0; i < len; i++) {
                        arr.push(HierarchicalUtil.parseItem(items[i], binding, bindingName, childItemsPath));
                    }
                    return arr;
                };
                HierarchicalUtil.isFlatItem = function (item, binding) {
                    if (wijmo.isArray(item[binding])) {
                        return false;
                    }
                    return true;
                };
                HierarchicalUtil.ConvertFlatData = function (items, binding, bindingName) {
                    var arr = [], data = {}, i, item, len = items.length;
                    for (i = 0; i < len; i++) {
                        item = items[i];
                        HierarchicalUtil.ConvertFlatItem(data, item, binding, bindingName);
                    }
                    HierarchicalUtil.ConvertFlatToHierarchical(arr, data);
                    return arr;
                };
                HierarchicalUtil.ConvertFlatToHierarchical = function (arr, data) {
                    var order = data['flatDataOrder'];
                    if (order) {
                        order.forEach(function (v) {
                            var d = {}, val = data[v], items;
                            d[data['field']] = v;
                            if (val['flatDataOrder']) {
                                items = [];
                                HierarchicalUtil.ConvertFlatToHierarchical(items, val);
                                d.items = items;
                            }
                            else {
                                d.value = val;
                            }
                            arr.push(d);
                        });
                    }
                };
                HierarchicalUtil.ConvertFlatItem = function (data, item, binding, bindingName) {
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
                        }
                        else {
                            data['flatDataOrder'] = [itemName];
                        }
                        data['field'] = name;
                    }
                    else {
                        if (data[itemName] == null) {
                            data[itemName] = {};
                            if (data['flatDataOrder']) {
                                data['flatDataOrder'].push(itemName);
                            }
                            else {
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
                };
                HierarchicalUtil.parseItem = function (item, binding, bindingName, childItemsPath) {
                    var data = {}, newBindingName, name, value, len, childItem, newChildItemsPath;
                    if (wijmo.isArray(childItemsPath)) {
                        newChildItemsPath = childItemsPath.slice();
                        childItem = newChildItemsPath.length ? newChildItemsPath.shift().trim() : '';
                    }
                    else {
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
                        }
                        else {
                            if (value && wijmo.isArray(value)) {
                                data.items = HierarchicalUtil.parseItems(value, binding, newBindingName, newChildItemsPath);
                            }
                            else {
                                data.value = item[binding];
                            }
                        }
                    }
                    else {
                        data.nameField = bindingName;
                        data.name = item[bindingName];
                        value = item[childItem];
                        if (value != null && wijmo.isArray(value)) {
                            data.items = HierarchicalUtil.parseItems(value, binding, bindingName, newChildItemsPath);
                        }
                        else {
                            data.value = item[binding];
                        }
                    }
                    return data;
                };
                HierarchicalUtil.parseFlatItem = function (data, item, binding, bindingName) {
                    if (!data.items) {
                        data.items = [];
                    }
                };
                return HierarchicalUtil;
            }());
            hierarchical.HierarchicalUtil = HierarchicalUtil;
        })(hierarchical = chart.hierarchical || (chart.hierarchical = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=HierarchicalUtil.js.map
