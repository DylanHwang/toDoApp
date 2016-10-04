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
* Analytics extensions for @see:FinancialChart.
*/
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart) {
        var finance;
        (function (finance) {
            var analytics;
            (function (analytics) {
                'use strict';
                // internal helper function to validate that a number is truly a number (and not Infinity, NaN, etc.)
                function isValid(value) {
                    return isFinite(value) && !isNaN(value) && wijmo.isNumber(value);
                }
                /**
                 * Represents a Fibonacci Retracements tool for the @see:FinancialChart.
            
                 * The tool enables the calculation and plotting of various alert levels that are
                 * useful in financial charts.
                 *
                 * To add Fibonacci tool to a @see:FinancialChart control, create an instance
                 * of the @see:Fibonacci and add it to the <b>series</b> collection of the chart.
                 * For example:
                 *
                 * <pre>
                 * // create chart
                 * var chart = new wijmo.chart.finance.FinancialChart('#chartElement');
                 * // create Fibonacci tool
                 * var ftool = new wijmo.chart.finance.analytics.Fibonacci();
                 * chart.series.push(ftool);
                 * </pre>
                  */
                var Fibonacci = (function (_super) {
                    __extends(Fibonacci, _super);
                    /**
                     * Initializes a new instance of the @see:Fibonacci class.
                     *
                     * @param options A JavaScript object containing initialization data.
                     */
                    function Fibonacci(options) {
                        _super.call(this);
                        this._levels = [0, 23.6, 38.2, 50, 61.8, 100];
                        this._uptrend = true;
                        this._labelPosition = chart.LabelPosition.Left;
                        if (options) {
                            wijmo.copy(this, options);
                        }
                        this.rendering.addHandler(this._render);
                    }
                    Object.defineProperty(Fibonacci.prototype, "low", {
                        /**
                         * Gets or sets the low value of @see:Fibonacci tool.
                         *
                         * If not specified, the low value is calculated based on data values provided by <b>itemsSource</b>.
                         */
                        get: function () {
                            return this._low;
                        },
                        set: function (value) {
                            if (value != this._low) {
                                this._low = wijmo.asNumber(value, true);
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(Fibonacci.prototype, "high", {
                        /**
                         * Gets or sets the high value of @see:Fibonacci tool.
                         *
                         * If not specified, the high value is caclulated based on
                         * data values provided by the <b>itemsSource</b>.
                         */
                        get: function () {
                            return this._high;
                        },
                        set: function (value) {
                            if (value != this._high) {
                                this._high = wijmo.asNumber(value, true);
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(Fibonacci.prototype, "labelPosition", {
                        /**
                         * Gets or sets the label position for levels in @see:Fibonacci tool.
                         */
                        get: function () {
                            return this._labelPosition;
                        },
                        set: function (value) {
                            if (value != this._labelPosition) {
                                this._labelPosition = wijmo.asEnum(value, chart.LabelPosition, true);
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(Fibonacci.prototype, "uptrend", {
                        /**
                         * Gets or sets a value indicating whether to create uptrending @see:Fibonacci tool.
                         *
                         * Default value is true(uptrend). If the value is false, the downtrending levels are plotted.
                         */
                        get: function () {
                            return this._uptrend;
                        },
                        set: function (value) {
                            if (value != this._uptrend) {
                                this._uptrend = wijmo.asBoolean(value, true);
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(Fibonacci.prototype, "levels", {
                        /**
                         * Gets or sets the array of levels for plotting.
                         *
                         * Default value is [0, 23.6, 38.2, 50, 61.8, 100].
                         */
                        get: function () {
                            return this._levels;
                        },
                        set: function (value) {
                            if (value != this._levels) {
                                this._levels = wijmo.asArray(value, true);
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(Fibonacci.prototype, "minX", {
                        /**
                         * Gets or sets the x minimal value of the @see:Fibonacci tool.
                         *
                         * If not specified, current minimum of x-axis is used.
                         * The value can be specified as a number or Date object.
                         */
                        get: function () {
                            return this._minX;
                        },
                        set: function (value) {
                            if (value != this._minX) {
                                this._minX = value;
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(Fibonacci.prototype, "maxX", {
                        /**
                         * Gets or sets the x maximum value of the @see:Fibonacci tool.
                         *
                         * If not specified, current maximum of x-axis is used.
                         * The value can be specified as a number or Date object.
                         */
                        get: function () {
                            return this._maxX;
                        },
                        set: function (value) {
                            if (value != this._maxX) {
                                this._maxX = value;
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Fibonacci.prototype._getMinX = function () {
                        if (wijmo.isNumber(this._minX)) {
                            return this._minX;
                        }
                        else if (wijmo.isDate(this._minX)) {
                            return wijmo.asDate(this._minX).valueOf();
                        }
                        else {
                            return this._getAxisX().actualMin;
                        }
                    };
                    Fibonacci.prototype._getMaxX = function () {
                        if (wijmo.isNumber(this._maxX)) {
                            return this._maxX;
                        }
                        else if (wijmo.isDate(this._maxX)) {
                            return wijmo.asDate(this._maxX).valueOf();
                        }
                        else {
                            return this._getAxisX().actualMax;
                        }
                    };
                    Fibonacci.prototype._updateLevels = function () {
                        var min = undefined, max = undefined;
                        if (this._low === undefined || this._high === undefined) {
                            var vals = _super.prototype.getValues.call(this, 0);
                            var xvals = _super.prototype.getValues.call(this, 1);
                            if (vals) {
                                var len = vals.length;
                                var xmin = this._getMinX(), xmax = this._getMaxX();
                                for (var i = 0; i < len; i++) {
                                    var val = vals[i];
                                    var xval = xvals ? xvals[i] : i;
                                    if (xval < xmin || xval > xmax) {
                                        continue;
                                    }
                                    if (!isNaN(val)) {
                                        if (min === undefined || min > val) {
                                            min = val;
                                        }
                                        if (max === undefined || max < val) {
                                            max = val;
                                        }
                                    }
                                }
                            }
                        }
                        if (this._low === undefined && min !== undefined) {
                            this._actualLow = min;
                        }
                        else {
                            this._actualLow = this._low;
                        }
                        if (this._high === undefined && max !== undefined) {
                            this._actualHigh = max;
                        }
                        else {
                            this._actualHigh = this._high;
                        }
                    };
                    Fibonacci.prototype._render = function (sender, args) {
                        var ser = sender;
                        ser._updateLevels();
                        var ax = ser._getAxisX();
                        var ay = ser._getAxisY();
                        var eng = args.engine;
                        var swidth = 2, stroke = ser._getSymbolStroke(ser._chart.series.indexOf(ser));
                        var lstyle = chart._BasePlotter.cloneStyle(ser.style, ['fill']);
                        var tstyle = chart._BasePlotter.cloneStyle(ser.style, ['stroke']);
                        var clipPath = ser.chart._plotrectId;
                        eng.stroke = stroke;
                        eng.strokeWidth = swidth;
                        eng.textFill = stroke;
                        var xmin = ser._getMinX(), xmax = ser._getMaxX();
                        if (xmin < ax.actualMin) {
                            xmin = ax.actualMin;
                        }
                        if (xmax > ax.actualMax) {
                            xmax = ax.actualMax;
                        }
                        // start group clipping
                        eng.startGroup(null, clipPath);
                        var llen = ser._levels ? ser._levels.length : 0;
                        for (var i = 0; i < llen; i++) {
                            var lvl = ser._levels[i];
                            var x1 = ax.convert(xmin), x2 = ax.convert(xmax);
                            var y = ser.uptrend ?
                                ay.convert(ser._actualLow + 0.01 * lvl * (ser._actualHigh - ser._actualLow)) :
                                ay.convert(ser._actualHigh - 0.01 * lvl * (ser._actualHigh - ser._actualLow));
                            if (chart._DataInfo.isValid(x1) && chart._DataInfo.isValid(x2) && chart._DataInfo.isValid(y)) {
                                eng.drawLine(x1, y, x2, y, null, lstyle);
                                if (ser.labelPosition != chart.LabelPosition.None) {
                                    var s = lvl.toFixed(1) + '%';
                                    var va = 0;
                                    if ((ser.uptrend && i == 0) || (!ser.uptrend && i == llen - 1)) {
                                        va = 2;
                                    }
                                    switch (ser.labelPosition) {
                                        case chart.LabelPosition.Left:
                                            chart.FlexChartCore._renderText(eng, s, new wijmo.Point(x1, y), 0, va, null, null, tstyle);
                                            break;
                                        case chart.LabelPosition.Center:
                                            chart.FlexChartCore._renderText(eng, s, new wijmo.Point(0.5 * (x1 + x2), y), 1, va, null, null, tstyle);
                                            break;
                                        case chart.LabelPosition.Right:
                                            chart.FlexChartCore._renderText(eng, s, new wijmo.Point(x2, y), 2, va, null, null, tstyle);
                                            break;
                                    }
                                }
                            }
                        }
                        eng.stroke = null;
                        eng.strokeWidth = null;
                        eng.textFill = null;
                        // end group
                        eng.endGroup();
                    };
                    Fibonacci.prototype._getChartType = function () {
                        return chart.ChartType.Line;
                    };
                    return Fibonacci;
                }(chart.SeriesBase));
                analytics.Fibonacci = Fibonacci;
                /**
                 * Represents a Fibonacci Arcs tool for the @see:FinancialChart.
                 */
                var FibonacciArcs = (function (_super) {
                    __extends(FibonacciArcs, _super);
                    /**
                     * Initializes a new instance of the @see:FibonacciArcs class.
                     *
                     * @param options A JavaScript object containing initialization data.
                     */
                    function FibonacciArcs(options) {
                        _super.call(this);
                        this._levels = [38.2, 50, 61.8];
                        this._labelPosition = chart.LabelPosition.Top;
                        // copy options
                        if (options) {
                            wijmo.copy(this, options);
                        }
                        this.rendering.addHandler(this._render, this);
                    }
                    Object.defineProperty(FibonacciArcs.prototype, "start", {
                        /**
                         * Gets or sets the starting @see:DataPoint for the base line.
                         *
                         * The @see:DataPoint x value can be a number or a Date object
                         * (for time-based data).
                         *
                         * Unlike some of the other Fibonacci tools, the starting
                         * @see:DataPoint is <b>not</b> calculated automatically if
                         * undefined.
                         */
                        get: function () {
                            return this._start;
                        },
                        set: function (value) {
                            if (value !== this.start) {
                                this._start = wijmo.asType(value, chart.DataPoint);
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(FibonacciArcs.prototype, "end", {
                        /**
                         * Gets or sets the ending @see:DataPoint for the base line.
                         *
                         * The @see:DataPoint x value can be a number or a Date object
                         * (for time-based data).
                         *
                         * Unlike some of the other Fibonacci tools, the ending
                         * @see:DataPoint is <b>not</b> calculated automatically if
                         * undefined.
                         */
                        get: function () {
                            return this._end;
                        },
                        set: function (value) {
                            if (value !== this.end) {
                                this._end = wijmo.asType(value, chart.DataPoint);
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(FibonacciArcs.prototype, "levels", {
                        /**
                         * Gets or sets the array of levels for plotting.
                         *
                         * Default value is [38.2, 50, 61.8].
                         */
                        get: function () {
                            return this._levels;
                        },
                        set: function (value) {
                            if (value !== this._levels) {
                                this._levels = wijmo.asArray(value, false);
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(FibonacciArcs.prototype, "labelPosition", {
                        /**
                         * Gets or sets the @see:LabelPosition for levels in @see:FibonacciArcs tool.
                         */
                        get: function () {
                            return this._labelPosition;
                        },
                        set: function (value) {
                            if (value !== this.labelPosition) {
                                this._labelPosition = wijmo.asEnum(value, chart.LabelPosition);
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    FibonacciArcs.prototype._render = function (sender, args) {
                        var startX = this._getX(0), startY = this._getY(0), endX = this._getX(1), endY = this._getY(1);
                        if (_super.prototype._getLength.call(this) <= 1 || !isValid(startX) || !isValid(startY) || !isValid(endX) || !isValid(endY)) {
                            return;
                        }
                        var ax = this._getAxisX(), ay = this._getAxisY(), engine = args.engine, swidth = 2, group, si = this.chart.series.indexOf(this), stroke = this._getSymbolStroke(si), lstyle = chart._BasePlotter.cloneStyle(this.style, ["fill"]), tstyle = chart._BasePlotter.cloneStyle(this.style, ["stroke"]);
                        engine.stroke = stroke;
                        engine.strokeWidth = swidth;
                        engine.textFill = stroke;
                        var clipPath = this.chart._plotrectId, yDiff = endY - startY, cx, cy, acy, baseLen, radius, center, lvl, size, lbl;
                        // start group for clipping
                        group = engine.startGroup(null, clipPath);
                        wijmo.addClass(group, 'fibonacci-arcs');
                        // draw base line
                        if (isValid(startX) && isValid(startY) && isValid(endX) && isValid(endY)) {
                            engine.drawLines([ax.convert(startX), ax.convert(endX)], [ay.convert(startY), ay.convert(endY)], null, lstyle);
                        }
                        // get length of base line
                        baseLen = Math.sqrt(Math.pow(ax.convert(endX) - ax.convert(startX), 2) + Math.pow(ay.convert(endY) - ay.convert(startY), 2));
                        // center point for arcs
                        center = new wijmo.Point(endX, endY);
                        // handle level arcs
                        for (var i = 0; i < this.levels.length; i++) {
                            // get level as decimal
                            lvl = this.levels[i] * 0.01;
                            // get the radius of the arc
                            radius = Math.abs(baseLen * lvl);
                            // draw the arc
                            if (isValid(center.x) && isValid(center.y) && isValid(radius)) {
                                cx = ax.convert(center.x);
                                cy = ay.convert(center.y);
                                // draw arc
                                engine.drawDonutSegment(cx, cy, radius, radius, yDiff > 0 ? 0 : Math.PI, Math.PI, null, lstyle);
                                // draw labels
                                if (this.labelPosition !== chart.LabelPosition.None && lvl !== 0) {
                                    // get label and determine its size
                                    lbl = wijmo.Globalize.format(lvl, "p1");
                                    size = engine.measureString(lbl, null, null, tstyle);
                                    // get label's y position
                                    acy = yDiff <= 0 ? cy - radius : cy + radius;
                                    switch (this.labelPosition) {
                                        case chart.LabelPosition.Center:
                                            acy += (size.height * 0.5);
                                            break;
                                        case chart.LabelPosition.Bottom:
                                            acy += yDiff <= 0 ? size.height : 0;
                                            break;
                                        default:
                                            acy += yDiff <= 0 ? 0 : size.height;
                                            break;
                                    }
                                    engine.drawString(lbl, new wijmo.Point(cx - size.width * .5, acy), null, tstyle);
                                }
                            }
                        }
                        engine.stroke = null;
                        engine.strokeWidth = null;
                        engine.textFill = null;
                        // end group
                        engine.endGroup();
                    };
                    FibonacciArcs.prototype._getX = function (dim) {
                        var retval = null;
                        if (dim === 0 && this.start) {
                            retval = this.start.x;
                        }
                        else if (dim === 1 && this.end) {
                            retval = this.end.x;
                        }
                        if (wijmo.isDate(retval)) {
                            retval = wijmo.asDate(retval).valueOf();
                        }
                        return retval;
                    };
                    FibonacciArcs.prototype._getY = function (dim) {
                        var retval = null;
                        if (dim === 0 && this.start) {
                            retval = this.start.y;
                        }
                        else if (dim === 1 && this.end) {
                            retval = this.end.y;
                        }
                        return retval;
                    };
                    FibonacciArcs.prototype._getChartType = function () {
                        return chart.ChartType.Line;
                    };
                    return FibonacciArcs;
                }(chart.SeriesBase));
                analytics.FibonacciArcs = FibonacciArcs;
                /**
                 * Represents a Fibonacci Fans tool for the @see:FinancialChart.
                 */
                var FibonacciFans = (function (_super) {
                    __extends(FibonacciFans, _super);
                    /**
                     * Initializes a new instance of the @see:FibonacciFans class.
                     *
                     * @param options A JavaScript object containing initialization data.
                     */
                    function FibonacciFans(options) {
                        _super.call(this);
                        this._levels = [0, 23.6, 38.2, 50, 61.8, 100];
                        this._labelPosition = chart.LabelPosition.Top;
                        // copy options
                        if (options) {
                            wijmo.copy(this, options);
                        }
                        this.rendering.addHandler(this._render, this);
                    }
                    Object.defineProperty(FibonacciFans.prototype, "start", {
                        /**
                         * Gets or sets the starting @see:DataPoint for the base line.
                         *
                         * If not set, the starting @see:DataPoint is calculated automatically.
                         * The @see:DataPoint x value can be a number or a Date object (for
                         * time-based data).
                         */
                        get: function () {
                            return this._start;
                        },
                        set: function (value) {
                            if (value !== this.start) {
                                this._start = wijmo.asType(value, chart.DataPoint);
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(FibonacciFans.prototype, "end", {
                        /**
                         * Gets or sets the ending @see:DataPoint for the base line.
                         *
                         * If not set, the starting @see:DataPoint is calculated automatically.
                         * The @see:DataPoint x value can be a number or a Date object (for
                         * time-based data).
                         */
                        get: function () {
                            return this._end;
                        },
                        set: function (value) {
                            if (value !== this.end) {
                                this._end = wijmo.asType(value, chart.DataPoint);
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(FibonacciFans.prototype, "levels", {
                        /**
                         * Gets or sets the array of levels for plotting.
                         *
                         * Default value is [0, 23.6, 38.2, 50, 61.8, 100].
                         */
                        get: function () {
                            return this._levels;
                        },
                        set: function (value) {
                            if (value !== this._levels) {
                                this._levels = wijmo.asArray(value, false);
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(FibonacciFans.prototype, "labelPosition", {
                        /**
                         * Gets or sets the @see:LabelPosition for levels in @see:FibonacciFans tool.
                         */
                        get: function () {
                            return this._labelPosition;
                        },
                        set: function (value) {
                            if (value !== this.labelPosition) {
                                this._labelPosition = wijmo.asEnum(value, chart.LabelPosition);
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    FibonacciFans.prototype._updateLevels = function () {
                        // both must be defined, otherwise we calulate start/end automatically
                        if (!this.start || !this.end) {
                            var plotter = this.chart._getPlotter(this), ax = this._getAxisX(), yvals = _super.prototype.getValues.call(this, 0), xvals = _super.prototype.getValues.call(this, 1) || plotter.dataInfo.getXVals(), xmin, xmax, ymin, ymax;
                            // use yvals only - no axisY.[actualMin|actualMax]
                            if (yvals && yvals.length > 0) {
                                ymin = finance._minimum(yvals);
                                ymax = finance._maximum(yvals);
                            }
                            if (xvals && xvals.length > 0) {
                                xmin = finance._minimum(xvals);
                                xmax = finance._maximum(xvals);
                            }
                            else {
                                xmin = ax.actualMin;
                                xmax = ax.actualMax;
                            }
                            if (isValid(xmin) && isValid(ymin) && isValid(xmax) && isValid(ymax)) {
                                this.start = new chart.DataPoint(xmin, ymin);
                                this.end = new chart.DataPoint(xmax, ymax);
                            }
                        }
                    };
                    FibonacciFans.prototype._render = function (sender, args) {
                        this._updateLevels();
                        var startX = this._getX(0), startY = this._getY(0), endX = this._getX(1), endY = this._getY(1);
                        if (_super.prototype._getLength.call(this) <= 1 || !isValid(startX) || !isValid(startY) || !isValid(endX) || !isValid(endY)) {
                            return;
                        }
                        var ax = this._getAxisX(), ay = this._getAxisY(), si = this.chart.series.indexOf(this), engine = args.engine, swidth = 2, stroke = this._getSymbolStroke(si), lstyle = chart._BasePlotter.cloneStyle(this.style, ["fill"]), tstyle = chart._BasePlotter.cloneStyle(this.style, ["stroke"]);
                        engine.stroke = stroke;
                        engine.strokeWidth = swidth;
                        engine.textFill = stroke;
                        var yDiff = endY - startY, xDiff = endX - startX, clipPath = this.chart._plotrectId, x1, x2, y1, y2, pt1, pt2, cp, m, b, lvl, lbl, size, angle;
                        // init local vars for start/end values
                        x1 = startX;
                        y1 = startY;
                        x2 = endX;
                        y2 = endY;
                        // maintain original x2 & set new x2
                        var x = x2;
                        // start group for clipping
                        engine.startGroup(null, clipPath);
                        // handle level lines
                        for (var i = 0; i < this.levels.length; i++) {
                            x2 = xDiff < 0 ? ax.actualMin : ax.actualMax;
                            // get level as decimal
                            lvl = this.levels[i] * 0.01;
                            // get level y2
                            y2 = y1 + lvl * yDiff;
                            // slope and y-intercept for (endX, new y2)
                            m = (y2 - y1) / (x - x1);
                            b = y2 - (m * x);
                            // update y2 for (ax.[actualMin||actualMax], new y2)
                            y2 = m * x2 + b;
                            // keep end point within plot area's bounds for labels
                            if (yDiff > 0 && y2 > ay.actualMax) {
                                y2 = ay.actualMax;
                                x2 = (y2 - b) / m;
                            }
                            else if (yDiff < 0 && y2 < ay.actualMin) {
                                y2 = ay.actualMin;
                                x2 = (y2 - b) / m;
                            }
                            if (isValid(x1) && isValid(y1) && isValid(x2) && isValid(y2)) {
                                // convert once per fan line & associated label
                                pt1 = new wijmo.Point(ax.convert(x1), ay.convert(y1));
                                pt2 = new wijmo.Point(ax.convert(x2), ay.convert(y2));
                                // draw fan line
                                engine.drawLines([pt1.x, pt2.x], [pt1.y, pt2.y], null, lstyle);
                                // draw fan label
                                if (this.labelPosition != chart.LabelPosition.None) {
                                    // get label and determine its size
                                    lbl = wijmo.Globalize.format(lvl, "p1");
                                    size = engine.measureString(lbl, null, null, tstyle);
                                    // find angle for label
                                    angle = Math.atan((pt2.y - pt1.y) / (pt2.x - pt1.x)) * 180 / Math.PI;
                                    // get center point by cloning the label point
                                    cp = pt2.clone();
                                    // update label point for axis boundx
                                    pt2.x = xDiff > 0 ? pt2.x - size.width : pt2.x;
                                    var a = angle * Math.PI / 180, tl = new wijmo.Point(), bl = new wijmo.Point(), tr = new wijmo.Point(), br = new wijmo.Point(), ymin = ay.convert(ay.actualMin), ymax = ay.convert(ay.actualMax), xmin = ax.convert(ax.actualMin), xmax = ax.convert(ax.actualMax), limit, acp = cp.clone();
                                    // adjust pt2.y based on label position property
                                    switch (this.labelPosition) {
                                        // top is the default by nature
                                        case chart.LabelPosition.Center:
                                            pt2.y += size.height * 0.5;
                                            // todo: this works okay, but corners should be calculated in this case
                                            acp.y += size.height * 0.5;
                                            break;
                                        case chart.LabelPosition.Bottom:
                                            pt2.y += size.height;
                                            break;
                                    }
                                    // http://math.stackexchange.com/questions/170650/how-to-get-upper-left-upper-right-lower-left-and-lower-right-corners-xy-coordi
                                    // attempt to keep labels in bounds
                                    if (xDiff > 0) {
                                        // todo: center is slightly off because the corners aren't correct
                                        // calculate coordinates of label's corners
                                        if (this.labelPosition === chart.LabelPosition.Top || this.labelPosition === chart.LabelPosition.Center) {
                                            br = acp.clone();
                                            tr.x = br.x + size.height * Math.sin(a);
                                            tr.y = br.y - size.height * Math.cos(a);
                                            tl.x = br.x - size.width * Math.cos(a) + size.height * Math.sin(a);
                                            tl.y = br.y - size.width * Math.sin(a) - size.height * Math.cos(a);
                                            bl.x = br.x - size.width * Math.cos(a);
                                            bl.y = br.y - size.width * Math.sin(a);
                                        }
                                        else if (this.labelPosition === chart.LabelPosition.Bottom) {
                                            tr = acp.clone();
                                            tl.x = tr.x - size.width * Math.cos(a);
                                            tl.y = tr.y - size.width * Math.sin(a);
                                            bl.x = tl.x - size.height * Math.sin(a);
                                            bl.y = tl.y + size.height * Math.cos(a);
                                            br.x = tl.x + size.width * Math.cos(a) - size.height * Math.sin(a);
                                            br.y = tl.y + size.width * Math.sin(a) + size.height * Math.cos(a);
                                        }
                                        // shift the label under certain conditions
                                        if (yDiff > 0) {
                                            if (tr.y < ymax) {
                                                m = (ay.convertBack(tr.y) - ay.convertBack(tl.y)) / (ax.convertBack(tr.x) - ax.convertBack(tl.x));
                                                b = ay.convertBack(tr.y) - (m * ax.convertBack(tr.x));
                                                limit = ax.convert((ay.actualMax - b) / m);
                                                pt2.x -= Math.abs(tr.x - limit);
                                            }
                                            if (br.x > xmax) {
                                                pt2.x -= Math.abs(xmax - br.x);
                                            }
                                        }
                                        else if (yDiff < 0) {
                                            if (br.y > ymin) {
                                                m = (ay.convertBack(bl.y) - ay.convertBack(br.y)) / (ax.convertBack(bl.x) - ax.convertBack(br.x));
                                                b = ay.convertBack(br.y) - (m * ax.convertBack(br.x));
                                                limit = ax.convert((ay.actualMin - b) / m);
                                                pt2.x -= Math.max(Math.abs(limit - br.x), Math.abs(ymin - br.y));
                                            }
                                            if (tr.x > xmax) {
                                                pt2.x -= Math.abs(xmax - tr.x);
                                            }
                                        }
                                    }
                                    else if (xDiff < 0) {
                                        // todo: center is slightly off because the corners aren't correct
                                        if (this.labelPosition === chart.LabelPosition.Top || this.labelPosition === chart.LabelPosition.Center) {
                                            bl = acp.clone();
                                            tl.x = bl.x + size.height * Math.sin(a);
                                            tl.y = bl.y - size.height * Math.cos(a);
                                            br.x = bl.x + size.width * Math.cos(a);
                                            br.y = bl.y + size.width * Math.sin(a);
                                            tr.x = tl.x + size.width * Math.cos(a);
                                            tr.y = tl.y + size.width * Math.sin(a);
                                        }
                                        else if (this.labelPosition === chart.LabelPosition.Bottom) {
                                            tl = acp.clone();
                                            tr.x = tl.x + size.width * Math.cos(a);
                                            tr.y = tl.y + size.width * Math.sin(a);
                                            bl.x = tl.x - size.height * Math.sin(a);
                                            bl.y = tl.y + size.height * Math.cos(a);
                                            br.x = tl.x + size.width * Math.cos(a) - size.height * Math.sin(a);
                                            br.y = tl.y + size.width * Math.sin(a) + size.height * Math.cos(a);
                                        }
                                        if (yDiff > 0) {
                                            if (tl.y < ymax) {
                                                m = (ay.convertBack(tl.y) - ay.convertBack(tr.y)) / (ax.convertBack(tl.x) - ax.convertBack(tr.x));
                                                b = ay.convertBack(tl.y) - (m * ax.convertBack(tl.x));
                                                limit = ax.convert((ay.actualMax - b) / m);
                                                pt2.x += Math.abs(tl.x - limit);
                                            }
                                            if (bl.x < xmin) {
                                                pt2.x += Math.abs(xmin - bl.x);
                                            }
                                        }
                                        else if (yDiff < 0) {
                                            if (bl.y > ymin) {
                                                m = (ay.convertBack(br.y) - ay.convertBack(bl.y)) / (ax.convertBack(br.x) - ax.convertBack(bl.x));
                                                b = ay.convertBack(bl.y) - (m * ax.convertBack(bl.x));
                                                limit = ax.convert((ay.actualMin - b) / m);
                                                pt2.x += Math.max(Math.abs(limit - bl.x), Math.abs(ymin - bl.y));
                                            }
                                            if (tl.x < xmin) {
                                                pt2.x += Math.abs(xmin - tl.x);
                                            }
                                        }
                                    }
                                    // draw the label
                                    if (angle === 0) {
                                        engine.drawString(lbl, pt2, null, tstyle);
                                    }
                                    else {
                                        engine.drawStringRotated(lbl, pt2, cp, angle, null, tstyle);
                                    }
                                }
                            }
                        }
                        engine.stroke = null;
                        engine.strokeWidth = null;
                        engine.textFill = null;
                        // end group
                        engine.endGroup();
                    };
                    FibonacciFans.prototype._getX = function (dim) {
                        var retval = null;
                        if (dim === 0 && this.start) {
                            retval = this.start.x;
                        }
                        else if (dim === 1 && this.end) {
                            retval = this.end.x;
                        }
                        if (wijmo.isDate(retval)) {
                            retval = wijmo.asDate(retval).valueOf();
                        }
                        return retval;
                    };
                    FibonacciFans.prototype._getY = function (dim) {
                        var retval = null;
                        if (dim === 0 && this.start) {
                            retval = this.start.y;
                        }
                        else if (dim === 1 && this.end) {
                            retval = this.end.y;
                        }
                        return retval;
                    };
                    FibonacciFans.prototype._getChartType = function () {
                        return chart.ChartType.Line;
                    };
                    return FibonacciFans;
                }(chart.SeriesBase));
                analytics.FibonacciFans = FibonacciFans;
                /**
                 * Represents a Fibonacci Time Zones tool for the @see:FinancialChart.
                 */
                var FibonacciTimeZones = (function (_super) {
                    __extends(FibonacciTimeZones, _super);
                    /**
                     * Initializes a new instance of the @see:FibonacciTimeZones class.
                     *
                     * @param options A JavaScript object containing initialization data.
                     */
                    function FibonacciTimeZones(options) {
                        _super.call(this);
                        this._levels = [0, 1, 2, 3, 5, 8, 13, 21, 34];
                        this._labelPosition = chart.LabelPosition.Right;
                        // copy options
                        if (options) {
                            wijmo.copy(this, options);
                        }
                        this.rendering.addHandler(this._render, this);
                    }
                    Object.defineProperty(FibonacciTimeZones.prototype, "startX", {
                        /**
                         * Gets or sets the starting X data point for the time zones.
                         *
                         * If not set, the starting X data point is calculated automatically. The
                         * value can be a number or a Date object (for time-based data).
                         */
                        get: function () {
                            return this._startX;
                        },
                        set: function (value) {
                            if (value !== this.startX) {
                                if (wijmo.isDate(value)) {
                                    this._startX = wijmo.asDate(value);
                                }
                                else {
                                    this._startX = wijmo.asNumber(value);
                                }
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(FibonacciTimeZones.prototype, "endX", {
                        /**
                         * Gets or sets the ending X data point for the time zones.
                         *
                         * If not set, the ending X data point is calculated automatically. The
                         * value can be a number or a Date object (for time-based data).
                         */
                        get: function () {
                            return this._endX;
                        },
                        set: function (value) {
                            if (value !== this.endX) {
                                if (wijmo.isDate(value)) {
                                    this._endX = wijmo.asDate(value);
                                }
                                else {
                                    this._endX = wijmo.asNumber(value);
                                }
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(FibonacciTimeZones.prototype, "levels", {
                        /**
                         * Gets or sets the array of levels for plotting.
                         *
                         * Default value is [0, 1, 2, 3, 5, 8, 13, 21, 34].
                         */
                        get: function () {
                            return this._levels;
                        },
                        set: function (value) {
                            if (value !== this._levels) {
                                this._levels = wijmo.asArray(value, false);
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(FibonacciTimeZones.prototype, "labelPosition", {
                        /**
                         * Gets or sets the @see:LabelPosition for levels in @see:FibonacciTimeZones tool.
                         */
                        get: function () {
                            return this._labelPosition;
                        },
                        set: function (value) {
                            if (value !== this.labelPosition) {
                                this._labelPosition = wijmo.asEnum(value, chart.LabelPosition);
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    FibonacciTimeZones.prototype._render = function (sender, args) {
                        this._updateLevels();
                        var start = this._getX(0), end = this._getX(1);
                        if (_super.prototype._getLength.call(this) <= 1 || !isValid(start) || !isValid(end)) {
                            return;
                        }
                        var diff = end - start, ax = this._getAxisX(), ay = this._getAxisY(), si = this._chart.series.indexOf(this), engine = args.engine, swidth = 2, stroke = this._getSymbolStroke(si), lstyle = chart._BasePlotter.cloneStyle(this.style, ["fill"]), tstyle = chart._BasePlotter.cloneStyle(this.style, ["stroke"]), ymin = ay.convert(ay.actualMin), ymax = ay.convert(ay.actualMax), lvl, x, size, lbl, clipPath = this.chart._plotrectId;
                        // render engine style settings
                        engine.stroke = stroke;
                        engine.strokeWidth = swidth;
                        engine.textFill = stroke;
                        // start and end cannot be equal
                        if (diff === 0) {
                            return;
                        }
                        // start group for clipping
                        engine.startGroup(null, clipPath);
                        // draw the time zones
                        for (var i = 0; i < this.levels.length; i++) {
                            lvl = this.levels[i];
                            x = diff * lvl + start;
                            if (x < ax.actualMin || ax.actualMax < x || !isValid(x)) {
                                continue;
                            }
                            // convert one time
                            x = ax.convert(x);
                            // draw line
                            engine.drawLine(x, ymin, x, ymax, null, lstyle);
                            // draw labels
                            if (this.labelPosition !== chart.LabelPosition.None) {
                                // get label and determine its size
                                lbl = wijmo.Globalize.format(lvl, "n0");
                                size = engine.measureString(lbl, null, null, tstyle);
                                // get label's x position
                                switch (this.labelPosition) {
                                    case chart.LabelPosition.Left:
                                        x -= size.width + swidth;
                                        break;
                                    case chart.LabelPosition.Center:
                                        x -= size.width / 2;
                                        break;
                                    case chart.LabelPosition.Right:
                                        x += swidth;
                                        break;
                                    default:
                                        x = diff < 0 ? x - size.width - swidth : x + swidth;
                                        break;
                                }
                                engine.drawString(lbl, new wijmo.Point(x, ymin), null, tstyle);
                            }
                        }
                        engine.stroke = null;
                        engine.strokeWidth = null;
                        engine.textFill = null;
                        // end group
                        engine.endGroup();
                    };
                    FibonacciTimeZones.prototype._updateLevels = function () {
                        var plotter = this.chart._getPlotter(this), xvals = _super.prototype.getValues.call(this, 1) || plotter.dataInfo.getXVals();
                        if (_super.prototype._getLength.call(this) <= 1) {
                            return;
                        }
                        // get startX & endX as numbers; both must be define or both are ignored
                        var start = this._getX(0), end = this._getX(1), defined = wijmo.isNumber(start) && wijmo.isNumber(end);
                        // automatically init startX & endX if not defined
                        if (!defined && !xvals) {
                            this._startX = 0;
                            this._endX = 1;
                        }
                        else if (!defined && xvals) {
                            this._startX = xvals[0];
                            this._endX = xvals[1];
                        }
                    };
                    FibonacciTimeZones.prototype._getX = function (dim) {
                        var retval = null;
                        if (dim === 0) {
                            retval = this.startX;
                        }
                        else if (dim === 1) {
                            retval = this.endX;
                        }
                        if (wijmo.isDate(retval)) {
                            retval = wijmo.asDate(retval).valueOf();
                        }
                        return retval;
                    };
                    FibonacciTimeZones.prototype._getChartType = function () {
                        return chart.ChartType.Line;
                    };
                    return FibonacciTimeZones;
                }(chart.SeriesBase));
                analytics.FibonacciTimeZones = FibonacciTimeZones;
            })(analytics = finance.analytics || (finance.analytics = {}));
        })(finance = chart.finance || (chart.finance = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Fibonacci.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart) {
        var finance;
        (function (finance) {
            var analytics;
            (function (analytics) {
                "use strict";
                /**
                 * Base class for overlay and indicator series (abstract).
                 */
                var OverlayIndicatorBase = (function (_super) {
                    __extends(OverlayIndicatorBase, _super);
                    function OverlayIndicatorBase() {
                        _super.call(this);
                        // internal field for series that need multiple legend entries
                        // in that case, set value to number of legend entries in ctor
                        this._seriesCount = 1;
                    }
                    Object.defineProperty(OverlayIndicatorBase.prototype, "_plotter", {
                        // access _IPlotter instance
                        get: function () {
                            if (this.chart && !this.__plotter) {
                                this.__plotter = this.chart._getPlotter(this);
                            }
                            return this.__plotter;
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(OverlayIndicatorBase.prototype, "_hitTester", {
                        // access _HitTester instance
                        get: function () {
                            if (this._plotter && !this.__hitTester) {
                                this.__hitTester = this._plotter.hitTester;
                            }
                            return this.__hitTester;
                        },
                        enumerable: true,
                        configurable: true
                    });
                    // return ChartType
                    OverlayIndicatorBase.prototype._getChartType = function () {
                        return chart.ChartType.Line;
                    };
                    // return original X-Values, if available
                    OverlayIndicatorBase.prototype._getXValues = function () {
                        return (_super.prototype.getValues.call(this, 1) || this._plotter.dataInfo.getXVals());
                    };
                    // helper method to get a _DataPoint object for hit testing
                    OverlayIndicatorBase.prototype._getDataPoint = function (dataX, dataY, seriesIndex, pointIndex, ax, ay) {
                        var dpt = new chart._DataPoint(seriesIndex, pointIndex, dataX, dataY);
                        // set x & y related data
                        dpt["y"] = dataY;
                        dpt["yfmt"] = ay._formatValue(dataY);
                        dpt["x"] = dataX;
                        dpt["xfmt"] = ax._formatValue(dataX);
                        return dpt;
                    };
                    // abstract method that determines whether or not calculations need to be ran
                    OverlayIndicatorBase.prototype._shouldCalculate = function () { return true; };
                    // initialize internal collections
                    OverlayIndicatorBase.prototype._init = function () { };
                    // responsible for calculating values
                    OverlayIndicatorBase.prototype._calculate = function () { };
                    OverlayIndicatorBase.prototype._clearValues = function () {
                        _super.prototype._clearValues.call(this);
                        this.__plotter = null;
                        this.__hitTester = null;
                    };
                    // helper for series with multiple names (csv)
                    // Returns undefined or the name.
                    OverlayIndicatorBase.prototype._getName = function (dim) {
                        var retval = undefined;
                        if (this.name) {
                            if (this.name.indexOf(",")) {
                                var names = this.name.split(",");
                                if (names && names.length - 1 >= dim) {
                                    retval = names[dim].trim();
                                }
                            }
                            else {
                                retval = this.name;
                            }
                        }
                        return retval;
                    };
                    // helper for series with multiple styles
                    // Returns the appropriate style for the given index, if
                    // ones exists; null is returned otherwise.
                    OverlayIndicatorBase.prototype._getStyles = function (dim) {
                        var retval = null;
                        if (dim < 0 || this._styles === null) {
                            return retval;
                        }
                        var i = 0;
                        for (var key in this._styles) {
                            if (i === dim && this._styles.hasOwnProperty(key)) {
                                retval = this._styles[key];
                                break;
                            }
                            i++;
                        }
                        return retval;
                    };
                    /* overrides for multiple legend items */
                    OverlayIndicatorBase.prototype.legendItemLength = function () {
                        return this._seriesCount;
                    };
                    OverlayIndicatorBase.prototype.measureLegendItem = function (engine, index) {
                        var name = this._getName(index), retval = new wijmo.Size(0, 0);
                        if (name) {
                            retval = this._measureLegendItem(engine, this._getName(index));
                        }
                        return retval;
                    };
                    OverlayIndicatorBase.prototype.drawLegendItem = function (engine, rect, index) {
                        var style = this._getStyles(index) || this.style, name = this._getName(index);
                        if (name) {
                            this._drawLegendItem(engine, rect, this._getChartType(), this._getName(index), style, this.symbolStyle);
                        }
                    };
                    return OverlayIndicatorBase;
                }(chart.SeriesBase));
                analytics.OverlayIndicatorBase = OverlayIndicatorBase;
                /**
                 * Base class for overlay and indicator series that render a single series (abstract).
                 */
                var SingleOverlayIndicatorBase = (function (_super) {
                    __extends(SingleOverlayIndicatorBase, _super);
                    function SingleOverlayIndicatorBase() {
                        _super.call(this);
                        this._period = 14;
                    }
                    Object.defineProperty(SingleOverlayIndicatorBase.prototype, "period", {
                        /**
                         * Gets or sets the period for the calculation as an integer value.
                         */
                        get: function () {
                            return this._period;
                        },
                        set: function (value) {
                            if (value !== this._period) {
                                this._period = wijmo.asInt(value, false, true);
                                this._clearValues();
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    // return the derived values
                    SingleOverlayIndicatorBase.prototype.getValues = function (dim) {
                        var retval = null;
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return retval;
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        if (dim === 0) {
                            retval = this._yvals;
                        }
                        else if (dim === 1) {
                            retval = this._xvals;
                        }
                        return retval;
                    };
                    // return limits for the derived values
                    SingleOverlayIndicatorBase.prototype.getDataRect = function () {
                        var rect = null;
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return rect;
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        var xmin = finance._minimum(this._xvals), xmax = finance._maximum(this._xvals), ymin = finance._minimum(this._yvals), ymax = finance._maximum(this._yvals);
                        if (chart._DataInfo.isValid(xmin) && chart._DataInfo.isValid(xmax) && chart._DataInfo.isValid(ymin) && chart._DataInfo.isValid(ymax)) {
                            rect = new wijmo.Rect(xmin, ymin, xmax - xmin, ymax - ymin);
                        }
                        return rect;
                    };
                    // clear the internal collections for the derived values
                    SingleOverlayIndicatorBase.prototype._clearValues = function () {
                        _super.prototype._clearValues.call(this);
                        this._xvals = null;
                        this._yvals = null;
                    };
                    // determine if the derived values need to be calculated
                    SingleOverlayIndicatorBase.prototype._shouldCalculate = function () {
                        return !this._yvals || !this._xvals;
                    };
                    // initialize internal collections for the derived values
                    SingleOverlayIndicatorBase.prototype._init = function () {
                        _super.prototype._init.call(this);
                        this._yvals = [];
                        this._xvals = [];
                    };
                    // override to get correct item for hit testing
                    SingleOverlayIndicatorBase.prototype._getItem = function (pointIndex) {
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return _super.prototype._getItem.call(this, pointIndex);
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        var originalLen = _super.prototype._getLength.call(this), len = finance._minimum(this._yvals.length, this._xvals.length);
                        // data index
                        pointIndex = originalLen - len + pointIndex;
                        return _super.prototype._getItem.call(this, pointIndex);
                    };
                    return SingleOverlayIndicatorBase;
                }(OverlayIndicatorBase));
                analytics.SingleOverlayIndicatorBase = SingleOverlayIndicatorBase;
            })(analytics = finance.analytics || (finance.analytics = {}));
        })(finance = chart.finance || (chart.finance = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=OverlayIndicatorBase.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart) {
        var finance;
        (function (finance) {
            var analytics;
            (function (analytics) {
                "use strict";
                /**
                 * Represents an Average True Range indicator series for the @see:FinancialChart.
                 *
                 * Average true range is used to measure the volatility of an asset. Average true range
                 * does not provide any indication of the price's trend, but rather the degree of price
                 * volatility.
                 */
                var ATR = (function (_super) {
                    __extends(ATR, _super);
                    function ATR() {
                        _super.call(this);
                        this.period = 14;
                    }
                    ATR.prototype._calculate = function () {
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return;
                        }
                        var highs = _super.prototype._getBindingValues.call(this, 0), lows = _super.prototype._getBindingValues.call(this, 1), closes = _super.prototype._getBindingValues.call(this, 3), xs = this._getXValues();
                        this._yvals = finance._avgTrueRng(highs, lows, closes, this.period);
                        this._xvals = xs ? xs.slice(this.period - 1) : finance._range(this.period - 1, highs.length);
                    };
                    return ATR;
                }(analytics.SingleOverlayIndicatorBase));
                analytics.ATR = ATR;
            })(analytics = finance.analytics || (finance.analytics = {}));
        })(finance = chart.finance || (chart.finance = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=ATR.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart) {
        var finance;
        (function (finance) {
            var analytics;
            (function (analytics) {
                "use strict";
                /**
                 * Represents a Commodity Channel Index indicator series for the @see:FinancialChart.
                 *
                 * The commodity channel index is an oscillator that measures an asset's current price
                 * level relative to an average price level over a specified period of time.
                 */
                var CCI = (function (_super) {
                    __extends(CCI, _super);
                    function CCI() {
                        _super.call(this);
                        this._constant = 0.015;
                        this.period = 20;
                    }
                    Object.defineProperty(CCI.prototype, "constant", {
                        /**
                         * Gets or sets the constant value for the CCI calculation.  The default
                         * value is 0.015.
                         */
                        get: function () {
                            return this._constant;
                        },
                        set: function (value) {
                            if (value !== this._constant) {
                                this._constant = wijmo.asNumber(value, false);
                                this._clearValues();
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    CCI.prototype._calculate = function () {
                        var originalLen = _super.prototype._getLength.call(this);
                        if (originalLen <= 0) {
                            return;
                        }
                        var highs = _super.prototype._getBindingValues.call(this, 0), lows = _super.prototype._getBindingValues.call(this, 1), closes = _super.prototype._getBindingValues.call(this, 3), xs = this._getXValues();
                        this._yvals = _cci(highs, lows, closes, this.period, this.constant);
                        this._xvals = xs ? xs.slice(this.period - 1) : finance._range(this.period - 1, originalLen - 1);
                    };
                    return CCI;
                }(analytics.SingleOverlayIndicatorBase));
                analytics.CCI = CCI;
                // calculate Commodity Channel Index for a set of financial data
                function _cci(highs, lows, closes, period, constant) {
                    wijmo.asArray(highs, false);
                    wijmo.asArray(lows, false);
                    wijmo.asArray(closes, false);
                    wijmo.asInt(period, false, true);
                    wijmo.asNumber(constant, false, true);
                    var len = finance._minimum(highs.length, lows.length, closes.length), typicalPrices = [], meanDeviations = [], smas, i, ccis = [];
                    wijmo.assert(len > period && period > 1, "CCI period must be an integer less than the length of the data and greater than one.");
                    // typical prices
                    for (i = 0; i < len; i++) {
                        typicalPrices.push(finance._average(highs[i], lows[i], closes[i]));
                    }
                    // simple moving average of typical prices
                    smas = finance._sma(typicalPrices, period);
                    // mean deviation
                    var temp;
                    for (i = 0; i < smas.length; i++) {
                        temp = typicalPrices.slice(i, period + i)
                            .reduce(function (prev, curr) { return prev + Math.abs(smas[i] - curr); }, 0);
                        meanDeviations.push(temp / period);
                    }
                    // get subset of typical prices
                    typicalPrices.splice(0, period - 1);
                    // cci
                    for (i = 0; i < smas.length; i++) {
                        ccis.push((typicalPrices[i] - smas[i]) / (constant * meanDeviations[i]));
                    }
                    return ccis;
                }
                analytics._cci = _cci;
            })(analytics = finance.analytics || (finance.analytics = {}));
        })(finance = chart.finance || (chart.finance = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=CCI.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart) {
        var finance;
        (function (finance) {
            var analytics;
            (function (analytics) {
                "use strict";
                /**
                 * Represents a Willaims %R indicator series for the @see:FinancialChart.
                 *
                 * Williams %R is a momentum indicator that is the inverse of a fast stochastic
                 * oscillator (@see:Stochastic).  The Williams %R indicator is designed to
                 * tell whether an asset is trading near the high or low of its trading range.
                 */
                var WilliamsR = (function (_super) {
                    __extends(WilliamsR, _super);
                    function WilliamsR() {
                        _super.call(this);
                        this.period = 14;
                    }
                    WilliamsR.prototype._calculate = function () {
                        var originalLen = _super.prototype._getLength.call(this);
                        if (originalLen <= 0) {
                            return;
                        }
                        var highs = _super.prototype._getBindingValues.call(this, 0), lows = _super.prototype._getBindingValues.call(this, 1), closes = _super.prototype._getBindingValues.call(this, 3), xs = this._getXValues();
                        this._yvals = _williamsR(highs, lows, closes, this.period);
                        this._xvals = xs ? xs.slice(this.period - 1) : finance._range(this.period - 1, originalLen - 1);
                    };
                    return WilliamsR;
                }(analytics.SingleOverlayIndicatorBase));
                analytics.WilliamsR = WilliamsR;
                // calculate Williams %R for a set of financial data
                function _williamsR(highs, lows, closes, period) {
                    wijmo.asArray(highs, false);
                    wijmo.asArray(lows, false);
                    wijmo.asArray(closes, false);
                    wijmo.asInt(period, false, true);
                    var len = finance._minimum(highs.length, lows.length, closes.length), extremeHighs = [], extremeLows = [], williamsRs = [], i;
                    wijmo.assert(len > period && period > 1, "Williams %R period must be an integer less than the length of the data and greater than one.");
                    // get extreme high/low for each period
                    for (i = period; i <= highs.length; i++) {
                        extremeHighs.push(finance._maximum(highs.slice(i - period, i)));
                        extremeLows.push(finance._minimum(lows.slice(i - period, i)));
                    }
                    // get subset of closing prices
                    closes.splice(0, period - 1);
                    // williams %r
                    for (i = 0; i < extremeHighs.length; i++) {
                        williamsRs.push((extremeHighs[i] - closes[i]) / (extremeHighs[i] - extremeLows[i]) * -100);
                    }
                    return williamsRs;
                }
                analytics._williamsR = _williamsR;
            })(analytics = finance.analytics || (finance.analytics = {}));
        })(finance = chart.finance || (chart.finance = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=WilliamsR.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart) {
        var finance;
        (function (finance) {
            var analytics;
            (function (analytics) {
                "use strict";
                (function (MovingAverageType) {
                    MovingAverageType[MovingAverageType["Simple"] = 0] = "Simple";
                    MovingAverageType[MovingAverageType["Exponential"] = 1] = "Exponential";
                })(analytics.MovingAverageType || (analytics.MovingAverageType = {}));
                var MovingAverageType = analytics.MovingAverageType;
                /**
                 * Represents a Moving Average Envelopes overlay series for the @see:FinancialChart.
                 *
                 * Moving average envelopes are moving averages set above and below a standard moving
                 * average.  The amount above/below the standard moving average is percentage based and
                 * dictated by the @see:size property.
                 */
                var Envelopes = (function (_super) {
                    __extends(Envelopes, _super);
                    function Envelopes() {
                        _super.call(this);
                        this._period = 20;
                        this._type = MovingAverageType.Simple;
                        this._size = 0.025;
                        this.rendering.addHandler(this._rendering, this);
                    }
                    Object.defineProperty(Envelopes.prototype, "period", {
                        /**
                         * Gets or sets the period for the calculation as an integer value.
                         */
                        get: function () {
                            return this._period;
                        },
                        set: function (value) {
                            if (value !== this._period) {
                                this._period = wijmo.asInt(value, false, true);
                                this._clearValues();
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(Envelopes.prototype, "type", {
                        /**
                         * Gets or sets the moving average type for the
                         * envelopes.  The default value is Simple.
                         */
                        get: function () {
                            return this._type;
                        },
                        set: function (value) {
                            if (value !== this._type) {
                                this._type = wijmo.asEnum(value, MovingAverageType, false);
                                this._clearValues();
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(Envelopes.prototype, "size", {
                        /**
                         * Gets or set the size of the moving average
                         * envelopes.  The default value is 2.5 percent (0.025).
                         */
                        get: function () {
                            return this._size;
                        },
                        set: function (value) {
                            if (value !== this._size) {
                                this._size = wijmo.asNumber(value, false, true);
                                this._clearValues();
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Envelopes.prototype.getDataRect = function () {
                        var rect = null;
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return rect;
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        var ys = this._upperYVals.concat(this._lowerYVals), xmin = finance._minimum(this._xVals), xmax = finance._maximum(this._xVals), ymin = finance._minimum(ys), ymax = finance._maximum(ys);
                        if (chart._DataInfo.isValid(xmin) && chart._DataInfo.isValid(xmax) && chart._DataInfo.isValid(ymin) && chart._DataInfo.isValid(ymax)) {
                            rect = new wijmo.Rect(xmin, ymin, xmax - xmin, ymax - ymin);
                        }
                        return rect;
                    };
                    Envelopes.prototype._clearValues = function () {
                        _super.prototype._clearValues.call(this);
                        this._upperYVals = null;
                        this._lowerYVals = null;
                        this._xVals = null;
                    };
                    Envelopes.prototype._init = function () {
                        _super.prototype._init.call(this);
                        this._upperYVals = [];
                        this._lowerYVals = [];
                        this._xVals = [];
                    };
                    Envelopes.prototype._shouldCalculate = function () {
                        return !this._upperYVals || !this._lowerYVals || !this._xVals;
                    };
                    // creates calculated values
                    Envelopes.prototype._calculate = function () {
                        var _this = this;
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return;
                        }
                        var ys = _super.prototype.getValues.call(this, 0), xs = this._getXValues(), avgs;
                        // moving average calculations
                        switch (this.type) {
                            case MovingAverageType.Exponential:
                                avgs = finance._ema(ys, this.period);
                                break;
                            case MovingAverageType.Simple:
                            default:
                                avgs = finance._sma(ys, this.period);
                                break;
                        }
                        this._xVals = xs ? xs.slice(this.period - 1) : finance._range(this.period - 1, _super.prototype._getLength.call(this) - 1);
                        this._upperYVals = avgs.map(function (value) { return value + (value * _this.size); });
                        this._lowerYVals = avgs.map(function (value) { return value - (value * _this.size); });
                    };
                    // custom rendering in order to draw multiple lines for a single SeriesBase object
                    Envelopes.prototype._rendering = function (sender, args) {
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return;
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        var si = this.chart.series.indexOf(this), engine = args.engine, ax = this._getAxisX(), ay = this._getAxisY(), len = finance._minimum(this._upperYVals.length, this._lowerYVals.length, this._xVals.length), style = chart._BasePlotter.cloneStyle(this.style, ["fill"]), stroke = this._getSymbolStroke(si), clipPath = this.chart._plotrectId, swidth = 2;
                        if (!len || len <= 0) {
                            return;
                        }
                        engine.stroke = stroke;
                        engine.strokeWidth = swidth;
                        var xs = [], uys = [], lys = [], originalLen = this._getLength(), dpt, area, di;
                        for (var i = 0; i < len; i++) {
                            // data index
                            di = originalLen - len + i;
                            // x values
                            xs.push(ax.convert(this._xVals[i]));
                            // upper
                            uys.push(ay.convert(this._upperYVals[i]));
                            dpt = this._getDataPoint(this._xVals[i], this._upperYVals[i], si, di, ax, ay);
                            area = new chart._CircleArea(new wijmo.Point(xs[i], uys[i]), 0.5 * engine.strokeWidth);
                            area.tag = dpt;
                            this._hitTester.add(area, si);
                            // lower
                            lys.push(ay.convert(this._lowerYVals[i]));
                            dpt = this._getDataPoint(this._xVals[i], this._lowerYVals[i], si, di, ax, ay);
                            area = new chart._CircleArea(new wijmo.Point(xs[i], lys[i]), 0.5 * engine.strokeWidth);
                            area.tag = dpt;
                            this._hitTester.add(area, si);
                        }
                        this._hitTester.add(new chart._LinesArea(xs, uys), si);
                        this._hitTester.add(new chart._LinesArea(xs, lys), si);
                        engine.drawLines(xs, uys, null, style, clipPath);
                        engine.drawLines(xs, lys, null, style, clipPath);
                    };
                    Envelopes.prototype.getCalculatedValues = function (key) {
                        key = wijmo.asString(key, false);
                        var retval = [], i = 0;
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return retval;
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        switch (key) {
                            case "upperEnvelope":
                                for (; i < this._upperYVals.length; i++) {
                                    retval.push({
                                        x: this._xVals[i],
                                        y: this._upperYVals[i]
                                    });
                                }
                                break;
                            case "lowerEnvelope":
                                for (; i < this._lowerYVals.length; i++) {
                                    retval.push({
                                        x: this._xVals[i],
                                        y: this._lowerYVals[i]
                                    });
                                }
                                break;
                        }
                        return retval;
                    };
                    return Envelopes;
                }(analytics.OverlayIndicatorBase));
                analytics.Envelopes = Envelopes;
            })(analytics = finance.analytics || (finance.analytics = {}));
        })(finance = chart.finance || (chart.finance = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Envelopes.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart) {
        var finance;
        (function (finance) {
            var analytics;
            (function (analytics) {
                "use strict";
                /**
                 * Represents a Bollinger Bands&reg; overlay series for the @see:FinancialChart.
                 *
                 * <i>Bollinger Bands is a registered trademark of John Bollinger.</i>
                 */
                var BollingerBands = (function (_super) {
                    __extends(BollingerBands, _super);
                    function BollingerBands() {
                        _super.call(this);
                        this._period = 20;
                        this._multiplier = 2;
                        this.rendering.addHandler(this._rendering, this);
                    }
                    Object.defineProperty(BollingerBands.prototype, "period", {
                        /**
                         * Gets or sets the period for the calculation as an integer value.
                         */
                        get: function () {
                            return this._period;
                        },
                        set: function (value) {
                            if (value !== this._period) {
                                this._period = wijmo.asInt(value, false, true);
                                this._clearValues();
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(BollingerBands.prototype, "multiplier", {
                        /**
                         * Gets or sets the standard deviation multiplier.
                         */
                        get: function () {
                            return this._multiplier;
                        },
                        set: function (value) {
                            if (value !== this._multiplier) {
                                this._multiplier = wijmo.asNumber(value, false, true);
                                this._clearValues();
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    BollingerBands.prototype.getDataRect = function () {
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return null;
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        var ys = this._upperYVals.concat(this._lowerYVals), xmin = finance._minimum(this._xVals), xmax = finance._maximum(this._xVals), ymin = finance._minimum(ys), ymax = finance._maximum(ys);
                        if (chart._DataInfo.isValid(xmin) && chart._DataInfo.isValid(xmax) && chart._DataInfo.isValid(ymin) && chart._DataInfo.isValid(ymax)) {
                            return new wijmo.Rect(xmin, ymin, xmax - xmin, ymax - ymin);
                        }
                        else {
                            return null;
                        }
                    };
                    BollingerBands.prototype._clearValues = function () {
                        _super.prototype._clearValues.call(this);
                        this._upperYVals = null;
                        this._middleYVals = null;
                        this._lowerYVals = null;
                        this._xVals = null;
                    };
                    BollingerBands.prototype._shouldCalculate = function () {
                        return !this._upperYVals || !this._middleYVals || !this._lowerYVals || !this._xVals;
                    };
                    BollingerBands.prototype._init = function () {
                        _super.prototype._init.call(this);
                        this._upperYVals = [];
                        this._middleYVals = [];
                        this._lowerYVals = [];
                        this._xVals = [];
                    };
                    BollingerBands.prototype._calculate = function () {
                        var originalLen = _super.prototype._getLength.call(this);
                        if (originalLen <= 0) {
                            return;
                        }
                        var ys = _super.prototype.getValues.call(this, 0), xs = this._getXValues();
                        var values = _bollingerBands(ys, this.period, this.multiplier);
                        this._upperYVals = values.uppers;
                        this._middleYVals = values.middles;
                        this._lowerYVals = values.lowers;
                        this._xVals = xs ? xs.slice(this.period - 1) : finance._range(this.period - 1, originalLen - 1);
                    };
                    BollingerBands.prototype._rendering = function (sender, args) {
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return;
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        var si = this.chart.series.indexOf(this), engine = args.engine, ax = this._getAxisX(), ay = this._getAxisY(), len = finance._minimum(this._upperYVals.length, this._middleYVals.length, this._lowerYVals.length, this._xVals.length), style = chart._BasePlotter.cloneStyle(this.style, ["fill"]), stroke = this._getSymbolStroke(si), clipPath = this.chart._plotrectId, swidth = 2;
                        if (!len || len <= 0) {
                            return;
                        }
                        engine.stroke = stroke;
                        engine.strokeWidth = swidth;
                        var xs = [], uys = [], mys = [], lys = [], originalLen = this._getLength(), dpt, area, di;
                        for (var i = 0; i < len; i++) {
                            // data index
                            di = originalLen - len + i;
                            // x values
                            xs.push(ax.convert(this._xVals[i]));
                            // upper
                            uys.push(ay.convert(this._upperYVals[i]));
                            dpt = this._getDataPoint(this._xVals[i], this._upperYVals[i], si, di, ax, ay);
                            area = new chart._CircleArea(new wijmo.Point(xs[i], uys[i]), 0.5 * engine.strokeWidth);
                            area.tag = dpt;
                            this._hitTester.add(area, si);
                            // middle
                            mys.push(ay.convert(this._middleYVals[i]));
                            dpt = this._getDataPoint(this._xVals[i], this._middleYVals[i], si, di, ax, ay);
                            area = new chart._CircleArea(new wijmo.Point(xs[i], mys[i]), 0.5 * engine.strokeWidth);
                            area.tag = dpt;
                            this._hitTester.add(area, si);
                            // lower
                            lys.push(ay.convert(this._lowerYVals[i]));
                            dpt = this._getDataPoint(this._xVals[i], this._lowerYVals[i], si, di, ax, ay);
                            area = new chart._CircleArea(new wijmo.Point(xs[i], lys[i]), 0.5 * engine.strokeWidth);
                            area.tag = dpt;
                            this._hitTester.add(area, si);
                        }
                        this._hitTester.add(new chart._LinesArea(xs, uys), si);
                        this._hitTester.add(new chart._LinesArea(xs, mys), si);
                        this._hitTester.add(new chart._LinesArea(xs, lys), si);
                        engine.drawLines(xs, uys, null, style, clipPath);
                        engine.drawLines(xs, mys, null, style, clipPath);
                        engine.drawLines(xs, lys, null, style, clipPath);
                    };
                    BollingerBands.prototype.getCalculatedValues = function (key) {
                        key = wijmo.asString(key, false);
                        var retval = [], i = 0;
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return retval;
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        switch (key) {
                            case "upperBand":
                                for (; i < this._upperYVals.length; i++) {
                                    retval.push({
                                        x: this._xVals[i],
                                        y: this._upperYVals[i]
                                    });
                                }
                                break;
                            case "middleBand":
                                for (; i < this._middleYVals.length; i++) {
                                    retval.push({
                                        x: this._xVals[i],
                                        y: this._middleYVals[i]
                                    });
                                }
                                break;
                            case "lowerBand":
                                for (; i < this._lowerYVals.length; i++) {
                                    retval.push({
                                        x: this._xVals[i],
                                        y: this._lowerYVals[i]
                                    });
                                }
                                break;
                        }
                        return retval;
                    };
                    return BollingerBands;
                }(analytics.OverlayIndicatorBase));
                analytics.BollingerBands = BollingerBands;
                // calculate Bollinger Bands for a set of financial data
                function _bollingerBands(ys, period, multiplier) {
                    wijmo.asArray(ys, false);
                    wijmo.asInt(period, false, true);
                    wijmo.asNumber(multiplier, false, true);
                    wijmo.assert(ys.length > period && period > 1, "Bollinger Bands period must be an integer less than the length of the data and greater than one.");
                    var avgs = finance._sma(ys, period), devs = [], i;
                    // get standard deviations
                    for (i = period; i <= ys.length; i++) {
                        devs.push(finance._stdDeviation(ys.slice(i - period, i)));
                    }
                    var middles = avgs, uppers = avgs.map(function (value, index) { return value + (devs[index] * multiplier); }), lowers = avgs.map(function (value, index) { return value - (devs[index] * multiplier); });
                    return {
                        lowers: lowers,
                        middles: middles,
                        uppers: uppers
                    };
                }
                analytics._bollingerBands = _bollingerBands;
            })(analytics = finance.analytics || (finance.analytics = {}));
        })(finance = chart.finance || (chart.finance = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=BollingerBands.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart) {
        var finance;
        (function (finance) {
            var analytics;
            (function (analytics) {
                "use strict";
                /**
                 * Represents a Relative Strength Index indicator series for the @see:FinancialChart.
                 *
                 * Relative strength index is a momentum osciallator designed to measure the current
                 * and historical strength or weakness of an asset based on the closing prices of a
                 * recent trading period.
                 */
                var RSI = (function (_super) {
                    __extends(RSI, _super);
                    function RSI() {
                        _super.call(this);
                        this.period = 14;
                    }
                    RSI.prototype._calculate = function () {
                        var originalLen = _super.prototype._getLength.call(this);
                        if (originalLen <= 0) {
                            return;
                        }
                        var ys = _super.prototype._getBindingValues.call(this, 0), // getValues(0) is overridden
                        xs = this._getXValues();
                        this._yvals = _rsi(ys, this.period);
                        this._xvals = xs ? xs.slice(this.period) : finance._range(this.period, originalLen);
                    };
                    return RSI;
                }(analytics.SingleOverlayIndicatorBase));
                analytics.RSI = RSI;
                // calculate Relative Strength Index for a set of financial data
                function _rsi(ys, period) {
                    wijmo.asArray(ys, false);
                    wijmo.asInt(period, true, false);
                    wijmo.assert(ys.length > period && period > 1, "RSI period must be an integer less than the length of the data and greater than one.");
                    var changes = [], avgGains = [], avgLosses = [], gains, losses, rsis = [], rs, i;
                    // calculate changes
                    for (i = 1; i < ys.length; i++) {
                        changes.push(ys[i] - ys[i - 1]);
                    }
                    // get gains and losses
                    gains = changes.map(function (value) { return value > 0 ? value : 0; });
                    losses = changes.map(function (value) { return value < 0 ? Math.abs(value) : 0; });
                    // calculate rs and rsi
                    for (i = period; i <= changes.length; i++) {
                        if (i === period) {
                            avgGains.push(finance._sum(gains.slice(i - period, i)) / period);
                            avgLosses.push(finance._sum(losses.slice(i - period, i)) / period);
                        }
                        else {
                            avgGains.push((gains[i - 1] + (avgGains[i - period - 1] * (period - 1))) / period);
                            avgLosses.push((losses[i - 1] + (avgLosses[i - period - 1] * (period - 1))) / period);
                        }
                        rs = avgGains[i - period] / avgLosses[i - period];
                        rs = isFinite(rs) ? rs : 0;
                        rsis.push(100 - (100 / (1 + rs)));
                    }
                    return rsis;
                }
                analytics._rsi = _rsi;
            })(analytics = finance.analytics || (finance.analytics = {}));
        })(finance = chart.finance || (chart.finance = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=RSI.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart) {
        var finance;
        (function (finance) {
            var analytics;
            (function (analytics) {
                "use strict";
                /**
                 * Base class for @see:Macd and @see:MacdHistogram series (abstract).
                 */
                var MacdBase = (function (_super) {
                    __extends(MacdBase, _super);
                    function MacdBase() {
                        _super.call(this);
                        this._fastPeriod = 12;
                        this._slowPeriod = 26;
                        this._smoothingPeriod = 9;
                    }
                    Object.defineProperty(MacdBase.prototype, "fastPeriod", {
                        /**
                         * Gets or sets the fast exponential moving average period
                         * for the MACD line.
                         */
                        get: function () {
                            return this._fastPeriod;
                        },
                        set: function (value) {
                            if (value !== this._fastPeriod) {
                                this._fastPeriod = wijmo.asInt(value, false, true);
                                this._clearValues();
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(MacdBase.prototype, "slowPeriod", {
                        /**
                         * Gets or sets the slow exponential moving average period
                         * for the MACD line.
                         */
                        get: function () {
                            return this._slowPeriod;
                        },
                        set: function (value) {
                            if (value !== this._slowPeriod) {
                                this._slowPeriod = wijmo.asInt(value, false, true);
                                this._clearValues();
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(MacdBase.prototype, "smoothingPeriod", {
                        /**
                         * Gets or sets the exponential moving average period
                         * for the signal line.
                         */
                        get: function () {
                            return this._smoothingPeriod;
                        },
                        set: function (value) {
                            if (value !== this._smoothingPeriod) {
                                this._smoothingPeriod = wijmo.asInt(value, false, true);
                                this._clearValues();
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    MacdBase.prototype._clearValues = function () {
                        _super.prototype._clearValues.call(this);
                        this._macdVals = null;
                        this._macdXVals = null;
                        this._signalVals = null;
                        this._signalXVals = null;
                        this._histogramVals = null;
                        this._histogramXVals = null;
                    };
                    MacdBase.prototype._shouldCalculate = function () {
                        return !this._macdVals || !this._macdXVals ||
                            !this._signalVals || !this._signalXVals ||
                            !this._histogramVals || !this._histogramXVals;
                    };
                    MacdBase.prototype._init = function () {
                        _super.prototype._init.call(this);
                        this._macdVals = [];
                        this._macdXVals = [];
                        this._signalVals = [];
                        this._signalXVals = [];
                        this._histogramVals = [];
                        this._histogramXVals = [];
                    };
                    MacdBase.prototype._calculate = function () {
                        var originalLen = _super.prototype._getLength.call(this);
                        if (originalLen <= 0) {
                            return;
                        }
                        var ys = _super.prototype.getValues.call(this, 0), xs = this._getXValues();
                        var values = _macd(ys, this.fastPeriod, this.slowPeriod, this.smoothingPeriod);
                        this._macdVals = values.macds;
                        this._signalVals = values.signals;
                        this._histogramVals = values.histograms;
                        this._macdXVals = xs ? xs.slice(originalLen - this._macdVals.length, originalLen) : finance._range(originalLen - this._macdVals.length, originalLen - 1);
                        this._signalXVals = xs ? xs.slice(originalLen - this._signalVals.length, originalLen) : finance._range(originalLen - this._signalVals.length, originalLen - 1);
                        this._histogramXVals = xs ? xs.slice(originalLen - this._histogramVals.length, originalLen) : finance._range(originalLen - this._histogramVals.length, originalLen - 1);
                    };
                    return MacdBase;
                }(analytics.OverlayIndicatorBase));
                analytics.MacdBase = MacdBase;
                /**
                 * Represents a Moving Average Convergence/Divergence (MACD) indicator series
                 * for the @see:FinancialChart.
                 *
                 * The MACD indicator is designed to reveal changes in strength, direction, momentum,
                 * and duration of an asset's price trend.
                 */
                var Macd = (function (_super) {
                    __extends(Macd, _super);
                    function Macd() {
                        _super.call(this);
                        this._seriesCount = 2;
                        this.rendering.addHandler(this._rendering, this);
                    }
                    Object.defineProperty(Macd.prototype, "styles", {
                        /**
                         * Gets or sets the styles for the MACD and Signal lines.
                         *
                         * The following options are supported:
                         *
                         * <pre>series.styles = {
                         *   macdLine: {
                         *      stroke: 'red',
                         *      strokeWidth: 1
                         *   },
                         *   signalLine: {
                         *      stroke: 'green',
                         *      strokeWidth: 1
                         *   },
                         * }</pre>
                         */
                        get: function () {
                            return this._styles;
                        },
                        set: function (value) {
                            if (value !== this._styles) {
                                this._styles = value;
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Macd.prototype.getDataRect = function () {
                        var rect = null;
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return rect;
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        var ys = [], xs = [];
                        xs.push.apply(xs, this._macdXVals);
                        xs.push.apply(xs, this._signalXVals);
                        ys.push.apply(ys, this._macdVals);
                        ys.push.apply(ys, this._signalVals);
                        var xmin = finance._minimum(xs), xmax = finance._maximum(xs), ymin = finance._minimum(ys), ymax = finance._maximum(ys);
                        if (chart._DataInfo.isValid(xmin) && chart._DataInfo.isValid(xmax) && chart._DataInfo.isValid(ymin) && chart._DataInfo.isValid(ymax)) {
                            rect = new wijmo.Rect(xmin, ymin, xmax - xmin, ymax - ymin);
                        }
                        return rect;
                    };
                    Macd.prototype._rendering = function (sender, args) {
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return;
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        var si = this.chart.series.indexOf(this), engine = args.engine, ax = this._getAxisX(), ay = this._getAxisY(), style = chart._BasePlotter.cloneStyle(this.style, ["fill"]), stroke = this._getSymbolStroke(si), clipPath = this.chart._plotrectId, swidth = 2, macdStyle = null, macdStroke = stroke, macdStrokeWidth = swidth, signalStyle = null, signalStroke = stroke, signalStrokeWidth = swidth;
                        // handle "styles"
                        if (this.styles && wijmo.isObject(this.styles)) {
                            if (this.styles.macdLine && wijmo.isObject(this.styles.macdLine)) {
                                macdStyle = chart._BasePlotter.cloneStyle(this.styles.macdLine, ["fill"]);
                                macdStroke = macdStyle.stroke ? macdStyle.stroke : stroke;
                                macdStrokeWidth = macdStyle.strokeWidth ? macdStyle.strokeWidth : swidth;
                            }
                            if (this.styles.signalLine && wijmo.isObject(this.styles.signalLine)) {
                                signalStyle = chart._BasePlotter.cloneStyle(this.styles.signalLine, ["fill"]);
                                signalStroke = signalStyle.stroke ? signalStyle.stroke : stroke;
                                signalStrokeWidth = signalStyle.strokeWidth ? signalStyle.strokeWidth : swidth;
                            }
                        }
                        var macdVals = [], macdXVals = [], signalVals = [], signalXVals = [], dpt, area, originalLen = this._getLength(), i, di;
                        // macd line
                        for (i = 0; i < this._macdVals.length; i++) {
                            // data index
                            di = originalLen - this._macdVals.length + i;
                            // x & yvalues
                            macdXVals.push(ax.convert(this._macdXVals[i]));
                            macdVals.push(ay.convert(this._macdVals[i]));
                            // hit testing
                            dpt = this._getDataPoint(this._macdXVals[i], this._macdVals[i], si, di, ax, ay);
                            dpt["name"] = this._getName(0);
                            area = new chart._CircleArea(new wijmo.Point(macdXVals[i], macdVals[i]), 0.5 * engine.strokeWidth);
                            area.tag = dpt;
                            this._hitTester.add(area, si);
                        }
                        this._hitTester.add(new chart._LinesArea(macdXVals, macdVals), si);
                        engine.stroke = macdStroke;
                        engine.strokeWidth = macdStrokeWidth;
                        engine.drawLines(macdXVals, macdVals, null, style, clipPath);
                        // signal line
                        for (i = 0; i < this._signalVals.length; i++) {
                            // data index
                            di = originalLen - this._signalVals.length + i;
                            // x & yvalues
                            signalXVals.push(ax.convert(this._signalXVals[i]));
                            signalVals.push(ay.convert(this._signalVals[i]));
                            // hit testing
                            dpt = this._getDataPoint(this._signalXVals[i], this._signalVals[i], si, di, ax, ay);
                            dpt["name"] = this._getName(1);
                            area = new chart._CircleArea(new wijmo.Point(signalXVals[i], signalVals[i]), 0.5 * engine.strokeWidth);
                            area.tag = dpt;
                            this._hitTester.add(area, si);
                        }
                        this._hitTester.add(new chart._LinesArea(signalXVals, signalVals), si);
                        engine.stroke = signalStroke;
                        engine.strokeWidth = signalStrokeWidth;
                        engine.drawLines(signalXVals, signalVals, null, style, clipPath);
                    };
                    Macd.prototype.getCalculatedValues = function (key) {
                        key = wijmo.asString(key, false);
                        var retval = [], i = 0;
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return retval;
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        switch (key) {
                            case "macdLine":
                                for (; i < this._macdVals.length; i++) {
                                    retval.push({
                                        x: this._macdXVals[i],
                                        y: this._macdVals[i]
                                    });
                                }
                                break;
                            case "signalLine":
                                for (; i < this._signalVals.length; i++) {
                                    retval.push({
                                        x: this._signalXVals[i],
                                        y: this._signalVals[i]
                                    });
                                }
                                break;
                        }
                        return retval;
                    };
                    return Macd;
                }(MacdBase));
                analytics.Macd = Macd;
                /**
                 * Represents a Moving Average Convergence/Divergence (MACD) Histogram indicator series
                 * for the @see:FinancialChart.
                 *
                 * The MACD indicator is designed to reveal changes in strength, direction, momentum,
                 * and duration of an asset's price trend.
                 */
                var MacdHistogram = (function (_super) {
                    __extends(MacdHistogram, _super);
                    function MacdHistogram() {
                        _super.call(this);
                    }
                    MacdHistogram.prototype.getValues = function (dim) {
                        var retval = null;
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return retval;
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        if (dim === 0) {
                            retval = this._histogramVals;
                        }
                        else if (dim === 1) {
                            retval = this._histogramXVals;
                        }
                        return retval;
                    };
                    MacdHistogram.prototype.getDataRect = function () {
                        var rect = null;
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return rect;
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        var xmin = finance._minimum(this._histogramXVals), xmax = finance._maximum(this._histogramXVals), ymin = finance._minimum(this._histogramVals), ymax = finance._maximum(this._histogramVals);
                        if (chart._DataInfo.isValid(xmin) && chart._DataInfo.isValid(xmax) && chart._DataInfo.isValid(ymin) && chart._DataInfo.isValid(ymax)) {
                            rect = new wijmo.Rect(xmin, ymin, xmax - xmin, ymax - ymin);
                        }
                        return rect;
                    };
                    MacdHistogram.prototype._getChartType = function () {
                        return chart.ChartType.Column;
                    };
                    // override to get correct item for hit testing
                    MacdHistogram.prototype._getItem = function (pointIndex) {
                        var originalLen = _super.prototype._getLength.call(this), len = finance._minimum(this._histogramVals.length, this._histogramXVals.length);
                        // data index
                        pointIndex = originalLen - len + pointIndex;
                        return _super.prototype._getItem.call(this, pointIndex);
                    };
                    return MacdHistogram;
                }(MacdBase));
                analytics.MacdHistogram = MacdHistogram;
                // calculate MACD for a set of financial data
                function _macd(ys, fastPeriod, slowPeriod, smoothingPeriod) {
                    wijmo.asArray(ys, false);
                    wijmo.asInt(fastPeriod, false, true);
                    wijmo.asInt(slowPeriod, false, true);
                    wijmo.asInt(smoothingPeriod, false, true);
                    var opposite = fastPeriod > slowPeriod, temp;
                    if (opposite) {
                        temp = slowPeriod;
                        slowPeriod = fastPeriod;
                        fastPeriod = temp;
                    }
                    var fastEmas = finance._ema(ys, fastPeriod), slowEmas = finance._ema(ys, slowPeriod), macds = [], histograms = [], signals, i;
                    // get subset of fast emas for macd line calculation
                    fastEmas.splice(0, slowPeriod - fastPeriod);
                    // macd line
                    for (i = 0; i < fastEmas.length; i++) {
                        temp = fastEmas[i] - slowEmas[i];
                        if (opposite)
                            temp *= -1;
                        macds.push(temp);
                    }
                    // signal line
                    signals = finance._ema(macds, smoothingPeriod);
                    // macd histogram
                    var macdTemp = macds.slice(macds.length - signals.length, macds.length);
                    for (i = 0; i < macdTemp.length; i++) {
                        histograms.push(macdTemp[i] - signals[i]);
                    }
                    return {
                        macds: macds,
                        signals: signals,
                        histograms: histograms
                    };
                }
                analytics._macd = _macd;
            })(analytics = finance.analytics || (finance.analytics = {}));
        })(finance = chart.finance || (chart.finance = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Macd.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart) {
        var finance;
        (function (finance) {
            var analytics;
            (function (analytics) {
                "use strict";
                /**
                 * Represents a Stochastic Oscillator indicator series for the @see:FinancialChart.
                 *
                 * Stochastic oscillators are momentum indicators designed to predict price turning
                 * points by comparing an asset's closing price to its high-low range.
                 *
                 * The @see:Stochastic series can be used for fast (default), slow and full stochastic
                 * oscillators.  To create a slow or full stochastic oscillator, set the @see:smoothingPeriod
                 * to an integer value greater than one; slow stochastic oscillators generally use a fixed
                 * @see:smoothingPeriod of three.  To create or revert to a fast stochastic oscillator, set the
                 * @see:smoothingPeriod to an integer value of one.
                 */
                var Stochastic = (function (_super) {
                    __extends(Stochastic, _super);
                    function Stochastic() {
                        _super.call(this);
                        this._kPeriod = 14;
                        this._dPeriod = 3;
                        this._smoothingPeriod = 1;
                        this._seriesCount = 2;
                        this.rendering.addHandler(this._rendering, this);
                    }
                    Object.defineProperty(Stochastic.prototype, "kPeriod", {
                        /**
                         * Gets or sets the period for the %K calculation.
                         */
                        get: function () {
                            return this._kPeriod;
                        },
                        set: function (value) {
                            if (value !== this._kPeriod) {
                                this._kPeriod = wijmo.asInt(value, false, true);
                                this._clearValues();
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(Stochastic.prototype, "dPeriod", {
                        /**
                         * Gets or sets the period for the %D simple moving average.
                         */
                        get: function () {
                            return this._dPeriod;
                        },
                        set: function (value) {
                            if (value !== this._dPeriod) {
                                this._dPeriod = wijmo.asInt(value, false, true);
                                this._clearValues();
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(Stochastic.prototype, "smoothingPeriod", {
                        /**
                         * Gets or sets the smoothing period for full %K.
                         */
                        get: function () {
                            return this._smoothingPeriod;
                        },
                        set: function (value) {
                            if (value !== this._smoothingPeriod) {
                                this._smoothingPeriod = wijmo.asInt(value, false, true);
                                this._clearValues();
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Object.defineProperty(Stochastic.prototype, "styles", {
                        /**
                         * Gets or sets the styles for the %K and %D lines.
                         *
                         * The following options are supported:
                         *
                         * <pre>series.styles = {
                         *   kLine: {
                         *      stroke: 'red',
                         *      strokeWidth: 1
                         *   },
                         *   dLine: {
                         *      stroke: 'green',
                         *      strokeWidth: 1
                         *   },
                         * }</pre>
                         */
                        get: function () {
                            return this._styles;
                        },
                        set: function (value) {
                            if (value !== this._styles) {
                                this._styles = value;
                                this._invalidate();
                            }
                        },
                        enumerable: true,
                        configurable: true
                    });
                    Stochastic.prototype.getDataRect = function () {
                        var rect = null;
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return rect;
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        var ys = this._kVals.concat(this._dVals), xs = this._kXVals.concat(this._dXVals), xmin = finance._minimum(xs), xmax = finance._maximum(xs), ymin = finance._minimum(ys), ymax = finance._maximum(ys);
                        if (chart._DataInfo.isValid(xmin) && chart._DataInfo.isValid(xmax) && chart._DataInfo.isValid(ymin) && chart._DataInfo.isValid(ymax)) {
                            rect = new wijmo.Rect(xmin, ymin, xmax - xmin, ymax - ymin);
                        }
                        return rect;
                    };
                    Stochastic.prototype._clearValues = function () {
                        _super.prototype._clearValues.call(this);
                        this._kVals = null;
                        this._kXVals = null;
                        this._dVals = null;
                        this._dXVals = null;
                    };
                    Stochastic.prototype._shouldCalculate = function () {
                        return !this._kVals || !this._kXVals ||
                            !this._dVals || !this._dXVals;
                    };
                    Stochastic.prototype._init = function () {
                        _super.prototype._init.call(this);
                        this._kVals = [];
                        this._kXVals = [];
                        this._dVals = [];
                        this._dXVals = [];
                    };
                    Stochastic.prototype._calculate = function () {
                        var originalLen = _super.prototype._getLength.call(this);
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return;
                        }
                        var highs = _super.prototype._getBindingValues.call(this, 0), lows = _super.prototype._getBindingValues.call(this, 1), closes = _super.prototype._getBindingValues.call(this, 3), xs = this._getXValues();
                        var values = _stochastic(highs, lows, closes, this.kPeriod, this.dPeriod, this.smoothingPeriod);
                        this._kVals = values.ks;
                        this._dVals = values.ds;
                        // get %K x-values
                        this._kXVals = xs ? xs.slice(this.kPeriod - 1) : finance._range(this.kPeriod - 1, originalLen - 1);
                        if (this.smoothingPeriod && this.smoothingPeriod > 1) {
                            this._kXVals = this._kXVals.slice(this._kXVals.length - this._kVals.length, this._kXVals.length);
                        }
                        // get %D x-values
                        this._dXVals = this._kXVals.slice(this._kXVals.length - this._dVals.length, this._kXVals.length);
                    };
                    Stochastic.prototype._rendering = function (sender, args) {
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return;
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        var si = this.chart.series.indexOf(this), engine = args.engine, ax = this._getAxisX(), ay = this._getAxisY(), style = chart._BasePlotter.cloneStyle(this.style, ["fill"]), stroke = this._getSymbolStroke(si), clipPath = this.chart._plotrectId, swidth = 2, kStyle = null, kStroke = stroke, kStrokeWidth = swidth, dStyle = null, dStroke = stroke, dStrokeWidth = swidth;
                        // handle "styles"
                        if (this.styles && wijmo.isObject(this.styles)) {
                            if (this.styles.kLine && wijmo.isObject(this.styles.kLine)) {
                                kStyle = chart._BasePlotter.cloneStyle(this.styles.kLine, ["fill"]);
                                kStroke = kStyle.stroke ? kStyle.stroke : stroke;
                                kStrokeWidth = kStyle.strokeWidth ? kStyle.strokeWidth : swidth;
                            }
                            if (this.styles.dLine && wijmo.isObject(this.styles.dLine)) {
                                dStyle = chart._BasePlotter.cloneStyle(this.styles.dLine, ["fill"]);
                                dStroke = dStyle.stroke ? dStyle.stroke : stroke;
                                dStrokeWidth = dStyle.strokeWidth ? dStyle.strokeWidth : swidth;
                            }
                        }
                        var kVals = [], kXVals = [], dVals = [], dXVals = [], originalLen = this._getLength(), dpt, area, i, di;
                        // %K
                        for (i = 0; i < this._kVals.length; i++) {
                            // data index
                            di = originalLen - this._kVals.length + i;
                            // x & yvalues
                            kXVals.push(ax.convert(this._kXVals[i]));
                            kVals.push(ay.convert(this._kVals[i]));
                            // hit testing
                            dpt = this._getDataPoint(this._kXVals[i], this._kVals[i], si, di, ax, ay);
                            dpt["name"] = this._getName(0);
                            area = new chart._CircleArea(new wijmo.Point(kXVals[i], kVals[i]), 0.5 * engine.strokeWidth);
                            area.tag = dpt;
                            this._hitTester.add(area, si);
                        }
                        this._hitTester.add(new chart._LinesArea(kXVals, kVals), si);
                        engine.stroke = kStroke;
                        engine.strokeWidth = kStrokeWidth;
                        engine.drawLines(kXVals, kVals, null, style, clipPath);
                        // %D
                        for (i = 0; i < this._dVals.length; i++) {
                            // data index
                            di = originalLen - this._dVals.length + i;
                            // x & yvalues
                            dXVals.push(ax.convert(this._dXVals[i]));
                            dVals.push(ay.convert(this._dVals[i]));
                            // hit testing
                            dpt = this._getDataPoint(this._dXVals[i], this._dVals[i], si, di, ax, ay);
                            dpt["name"] = this._getName(1);
                            area = new chart._CircleArea(new wijmo.Point(dXVals[i], dVals[i]), 0.5 * engine.strokeWidth);
                            area.tag = dpt;
                            this._hitTester.add(area, si);
                        }
                        this._hitTester.add(new chart._LinesArea(dXVals, dVals), si);
                        engine.stroke = dStroke;
                        engine.strokeWidth = dStrokeWidth;
                        engine.drawLines(dXVals, dVals, null, style, clipPath);
                    };
                    Stochastic.prototype.getCalculatedValues = function (key) {
                        key = wijmo.asString(key, false);
                        var retval = [], i = 0;
                        if (_super.prototype._getLength.call(this) <= 0) {
                            return retval;
                        }
                        else if (this._shouldCalculate()) {
                            this._init();
                            this._calculate();
                        }
                        switch (key) {
                            case "kLine":
                                for (; i < this._kVals.length; i++) {
                                    retval.push({
                                        x: this._kXVals[i],
                                        y: this._kVals[i]
                                    });
                                }
                                break;
                            case "dLine":
                                for (; i < this._dVals.length; i++) {
                                    retval.push({
                                        x: this._dXVals[i],
                                        y: this._dVals[i]
                                    });
                                }
                                break;
                        }
                        return retval;
                    };
                    return Stochastic;
                }(analytics.OverlayIndicatorBase));
                analytics.Stochastic = Stochastic;
                // calculate Stochastics for a set of financial data
                function _stochastic(highs, lows, closes, kPeriod, dPeriod, smoothingPeriod) {
                    wijmo.asArray(highs, false);
                    wijmo.asArray(lows, false);
                    wijmo.asArray(closes, false);
                    wijmo.asInt(kPeriod, false, true);
                    wijmo.asInt(dPeriod, false, true);
                    wijmo.asInt(smoothingPeriod, true, true);
                    var extremeHighs = [], extremeLows = [], kvals = [], dvals, i;
                    // get extreme highs/lows for each period
                    for (i = kPeriod; i <= highs.length; i++) {
                        extremeHighs.push(finance._maximum(highs.slice(i - kPeriod, i)));
                        extremeLows.push(finance._minimum(lows.slice(i - kPeriod, i)));
                    }
                    // get subset of closing prices
                    closes = closes.slice(kPeriod - 1);
                    // %K
                    for (i = 0; i < closes.length; i++) {
                        kvals.push((closes[i] - extremeLows[i]) / (extremeHighs[i] - extremeLows[i]) * 100);
                    }
                    // %K in slow/full
                    if (smoothingPeriod && smoothingPeriod > 1) {
                        kvals = finance._sma(kvals, smoothingPeriod);
                    }
                    // %D
                    dvals = finance._sma(kvals, dPeriod);
                    return {
                        ks: kvals,
                        ds: dvals
                    };
                }
                analytics._stochastic = _stochastic;
            })(analytics = finance.analytics || (finance.analytics = {}));
        })(finance = chart.finance || (chart.finance = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Stochastic.js.map
