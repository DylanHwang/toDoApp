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
* Defines the @see:FlexRadar control and its associated classes.
*
*/
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart) {
        var radar;
        (function (radar) {
            'use strict';
            /**
             * Specifies the type of radar chart.
             */
            (function (RadarChartType) {
                /** Shows vertical bars and allows you to compare values of items across categories. */
                RadarChartType[RadarChartType["Column"] = 0] = "Column";
                /** Shows patterns within the data using X and Y coordinates. */
                RadarChartType[RadarChartType["Scatter"] = 1] = "Scatter";
                /** Shows trends over a period of time or across categories. */
                RadarChartType[RadarChartType["Line"] = 2] = "Line";
                /** Shows line chart with a symbol on each data point. */
                RadarChartType[RadarChartType["LineSymbols"] = 3] = "LineSymbols";
                /** Shows line chart with the area below the line filled with color. */
                RadarChartType[RadarChartType["Area"] = 4] = "Area";
            })(radar.RadarChartType || (radar.RadarChartType = {}));
            var RadarChartType = radar.RadarChartType;
            /**
             * radar chart control.
             */
            var FlexRadar = (function (_super) {
                __extends(FlexRadar, _super);
                /**
                 * Initializes a new instance of the @see:FlexRadar class.
                 *
                 * @param element The DOM element that hosts the control, or a selector for the
                 * host element (e.g. '#theCtrl').
                 * @param options A JavaScript object containing initialization data for the
                 * control.
                 */
                function FlexRadar(element, options) {
                    _super.call(this, element, options);
                    this._chartType = RadarChartType.Line;
                    this._startAngle = 0;
                    this._totalAngle = 360;
                    this._reversed = false;
                    this._areas = [];
                }
                Object.defineProperty(FlexRadar.prototype, "_radarLinePlotter", {
                    get: function () {
                        if (this.__radarLinePlotter == null) {
                            this.__radarLinePlotter = new radar._RadarLinePlotter();
                            this._initPlotter(this.__radarLinePlotter);
                        }
                        return this.__radarLinePlotter;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(FlexRadar.prototype, "_radarColumnPlotter", {
                    get: function () {
                        if (this.__radarColumnPlotter == null) {
                            this.__radarColumnPlotter = new radar._RadarBarPlotter();
                            this._initPlotter(this.__radarColumnPlotter);
                        }
                        return this.__radarColumnPlotter;
                    },
                    enumerable: true,
                    configurable: true
                });
                FlexRadar.prototype._initAxes = function () {
                    _super.prototype._initAxes.call(this);
                    this.axes.pop();
                    this.axes.pop();
                    this.axisX = new radar.FlexRadarAxis(chart.Position.Bottom);
                    this.axisX.majorGrid = true;
                    this.axisY = new radar.FlexRadarAxis(chart.Position.Left);
                    this.axisY.majorTickMarks = chart.TickMark.Outside;
                    this.axes.push(this.axisX);
                    this.axes.push(this.axisY);
                };
                FlexRadar.prototype._layout = function (rect, size, engine) {
                    _super.prototype._layout.call(this, rect, size, engine);
                    var height = this.axisX._height;
                    this._plotRect.top += height / 2;
                    var pr = this._plotRect;
                    this._radius = Math.min(pr.width, pr.height) / 2;
                    this._center = new wijmo.Point(pr.left + pr.width / 2, pr.top + pr.height / 2);
                };
                Object.defineProperty(FlexRadar.prototype, "chartType", {
                    /**
                     * Gets or sets the type of radar chart to create.
                     */
                    get: function () {
                        return this._chartType;
                    },
                    set: function (value) {
                        if (value != this._chartType) {
                            this._chartType = wijmo.asEnum(value, RadarChartType);
                            this.invalidate();
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(FlexRadar.prototype, "startAngle", {
                    /**
                     * Gets or sets the starting angle for the radar, in degrees.
                     *
                     * Angles are measured clockwise, starting at the 12 o'clock position.
                     */
                    get: function () {
                        return this._startAngle;
                    },
                    set: function (value) {
                        if (value != this._startAngle) {
                            this._startAngle = wijmo.asNumber(value, true);
                            this.invalidate();
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(FlexRadar.prototype, "totalAngle", {
                    /**
                     * Gets or sets the total angle for the radar, in degrees.  Its default value is 360.
                     * The value must be greater than 0, or less than or equal to 360.
                     */
                    get: function () {
                        return this._totalAngle;
                    },
                    set: function (value) {
                        if (value != this._totalAngle && value >= 0) {
                            this._totalAngle = wijmo.asNumber(value, true);
                            if (this._totalAngle <= 0) {
                                wijmo.assert(false, "totalAngle must be greater than 0.");
                                this._totalAngle = 0;
                            }
                            if (this._totalAngle > 360) {
                                wijmo.assert(false, "totalAngle must be less than or equal to 360.");
                                this._totalAngle = 360;
                            }
                            this.invalidate();
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(FlexRadar.prototype, "reversed", {
                    /**
                     * Gets or sets a value that determines whether angles are reversed
                     * (counter-clockwise).
                     *
                     * The default value is false, which causes angles to be measured in
                     * the clockwise direction.
                     */
                    get: function () {
                        return this._reversed;
                    },
                    set: function (value) {
                        if (value != this._reversed) {
                            this._reversed = wijmo.asBoolean(value, true);
                            this.invalidate();
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(FlexRadar.prototype, "stacking", {
                    /**
                     * Gets or sets a value that determines whether and how the series objects are stacked.
                     */
                    get: function () {
                        return this._stacking;
                    },
                    set: function (value) {
                        if (value != this._stacking) {
                            this._stacking = wijmo.asEnum(value, chart.Stacking);
                            this.invalidate();
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                FlexRadar.prototype._getChartType = function () {
                    var ct = chart.ChartType.Line;
                    switch (this.chartType) {
                        case RadarChartType.Area:
                            ct = chart.ChartType.Area;
                            break;
                        case RadarChartType.Line:
                            ct = chart.ChartType.Line;
                            break;
                        case RadarChartType.Column:
                            ct = chart.ChartType.Column;
                            break;
                        case RadarChartType.LineSymbols:
                            ct = chart.ChartType.LineSymbols;
                            break;
                        case RadarChartType.Scatter:
                            ct = chart.ChartType.Scatter;
                            break;
                    }
                    return ct;
                };
                FlexRadar.prototype._getPlotter = function (series) {
                    var chartType = this.chartType, plotter = null, isSeries = false;
                    if (series) {
                        var stype = series.chartType;
                        if (stype != null && stype != chartType) {
                            chartType = stype;
                            isSeries = true;
                        }
                    }
                    switch (chartType) {
                        // no plotter found for RadarChartType - try based on ChartType
                        case RadarChartType.Line:
                            this._radarLinePlotter.hasSymbols = false;
                            this._radarLinePlotter.hasLines = true;
                            this._radarLinePlotter.isArea = false;
                            plotter = this._radarLinePlotter;
                            break;
                        case RadarChartType.LineSymbols:
                            this._radarLinePlotter.hasSymbols = true;
                            this._radarLinePlotter.hasLines = true;
                            this._radarLinePlotter.isArea = false;
                            plotter = this._radarLinePlotter;
                            break;
                        case RadarChartType.Area:
                            this._radarLinePlotter.hasSymbols = false;
                            this._radarLinePlotter.hasLines = true;
                            this._radarLinePlotter.isArea = true;
                            plotter = this._radarLinePlotter;
                            break;
                        case RadarChartType.Scatter:
                            this._radarLinePlotter.hasSymbols = true;
                            this._radarLinePlotter.hasLines = false;
                            this._radarLinePlotter.isArea = false;
                            plotter = this._radarLinePlotter;
                            break;
                        case RadarChartType.Column:
                            this._radarColumnPlotter.isVolume = false;
                            this._radarColumnPlotter.width = 0.8;
                            plotter = this._radarColumnPlotter;
                            break;
                        default:
                            plotter = _super.prototype._getPlotter.call(this, series);
                            break;
                    }
                    return plotter;
                };
                FlexRadar.prototype._convertPoint = function (radius, angle) {
                    var pt = new wijmo.Point(), center = this._center;
                    pt.x = center.x + radius * Math.sin(angle);
                    pt.y = center.y - radius * Math.cos(angle);
                    return pt;
                };
                FlexRadar.prototype._createSeries = function () {
                    return new radar.FlexRadarSeries();
                };
                FlexRadar.prototype._clearCachedValues = function () {
                    _super.prototype._clearCachedValues.call(this);
                    this._isPolar = false;
                    this._areas = [];
                };
                FlexRadar.prototype._performBind = function () {
                    _super.prototype._performBind.call(this);
                    if (this._xDataType === wijmo.DataType.Number) {
                        this._isPolar = true;
                    }
                };
                FlexRadar.prototype._render = function (engine, applyElement) {
                    if (applyElement === void 0) { applyElement = true; }
                    this._areas = [];
                    _super.prototype._render.call(this, engine, applyElement);
                };
                return FlexRadar;
            }(chart.FlexChartCore));
            radar.FlexRadar = FlexRadar;
        })(radar = chart.radar || (chart.radar = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=FlexRadar.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart) {
        var radar;
        (function (radar) {
            'use strict';
            /**
             * Represents a series of data points to display in the chart.
             *
             * The @see:FlexRadarSeries class supports all basic chart types. You may define
             * a different chart type on each @see:FlexRadarSeries object that you add to the
             * @see:FlexRadar series collection. This overrides the @see:chartType
             * property set on the chart that is the default for all @see:FlexRadarSeries objects
             * in its collection.
             */
            var FlexRadarSeries = (function (_super) {
                __extends(FlexRadarSeries, _super);
                function FlexRadarSeries() {
                    _super.apply(this, arguments);
                }
                Object.defineProperty(FlexRadarSeries.prototype, "chartType", {
                    /**
                     * Gets or sets the chart type for a specific series, overriding the chart type
                     * set on the overall chart. Please note that ColumnVolume, EquiVolume,
                     * CandleVolume and ArmsCandleVolume chart types are not supported and should be
                     * set on the @see:FinancialChart.
                     */
                    get: function () {
                        return this._finChartType;
                    },
                    set: function (value) {
                        if (value != this._finChartType) {
                            this._finChartType = wijmo.asEnum(value, radar.RadarChartType, true);
                            this._invalidate();
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                FlexRadarSeries.prototype._getChartType = function () {
                    var ct;
                    switch (this.chartType) {
                        case radar.RadarChartType.Area:
                            ct = chart.ChartType.Area;
                            break;
                        case radar.RadarChartType.Line:
                            ct = chart.ChartType.Line;
                            break;
                        case radar.RadarChartType.Column:
                            ct = chart.ChartType.Column;
                            break;
                        case radar.RadarChartType.LineSymbols:
                            ct = chart.ChartType.LineSymbols;
                            break;
                        case radar.RadarChartType.Scatter:
                            ct = chart.ChartType.Scatter;
                            break;
                    }
                    return ct;
                };
                FlexRadarSeries.prototype._getLabelPoint = function (dataPoint) {
                    var ax = this._getAxisX(), ay = this._getAxisY(), angle = ax.convert(dataPoint.dataX), radius = ay.convert(dataPoint.dataY);
                    return this.chart._convertPoint(radius, angle);
                };
                return FlexRadarSeries;
            }(chart.SeriesBase));
            radar.FlexRadarSeries = FlexRadarSeries;
        })(radar = chart.radar || (chart.radar = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=FlexRadarSeries.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart_1) {
        var radar;
        (function (radar) {
            'use strict';
            /**
             * Represents an axis in the radar chart.
             */
            var FlexRadarAxis = (function (_super) {
                __extends(FlexRadarAxis, _super);
                function FlexRadarAxis() {
                    _super.apply(this, arguments);
                    this._points = [];
                    this._axisLabels = [];
                }
                FlexRadarAxis.prototype._render = function (engine) {
                    var _this = this;
                    _super.prototype._render.call(this, engine);
                    var labels = this._axisLabels;
                    if (labels.length) {
                        var renderLabels = function () {
                            var cls = _this.axisType == chart_1.AxisType.X ? 'wj-axis-x-labels ' + chart_1.FlexChart._CSS_AXIS_X : 'wj-axis-y-labels ' + chart_1.FlexChart._CSS_AXIS_Y;
                            engine.startGroup(cls);
                            labels.forEach(function (lbl) {
                                var labelAngle = lbl.labelAngle;
                                if (labelAngle > 0) {
                                    if (labelAngle == 90) {
                                        chart_1.FlexChart._renderRotatedText(engine, lbl.text, lbl.pos, lbl.align, lbl.vAlign, lbl.pos, labelAngle, lbl.class);
                                    }
                                    else {
                                        chart_1.FlexChart._renderRotatedText(engine, lbl.text, lbl.pos, lbl.align, lbl.vAlign, lbl.pos, labelAngle, lbl.class);
                                    }
                                }
                                else if (labelAngle < 0) {
                                    if (labelAngle == -90) {
                                        chart_1.FlexChart._renderRotatedText(engine, lbl.text, lbl.pos, lbl.align, lbl.vAlign, lbl.pos, labelAngle, lbl.class);
                                    }
                                    else {
                                        chart_1.FlexChart._renderRotatedText(engine, lbl.text, lbl.pos, lbl.align, lbl.vAlign, lbl.pos, labelAngle, lbl.class);
                                    }
                                }
                                else {
                                    _this._renderLabel(engine, lbl.val, lbl.text, lbl.pos, lbl.align, lbl.vAlign /*1*/, lbl.class);
                                }
                                //this._renderLabel(engine, lbl.val, lbl.text, lbl.pos, lbl.align, lbl.vAlign, lbl.class);
                            });
                            engine.endGroup();
                            _this._axisLabels = [];
                            _this._chart.rendered.removeHandler(renderLabels);
                        };
                        this._chart.rendered.addHandler(renderLabels, this);
                    }
                };
                FlexRadarAxis.prototype._getHeight = function (engine, maxw) {
                    var height = _super.prototype._getHeight.call(this, engine, maxw);
                    if (this._axisType == chart_1.AxisType.Y) {
                        height -= this.labelPadding * 2;
                        height += 4;
                    }
                    this._height = height * 2;
                    return height * 2;
                };
                FlexRadarAxis.prototype._updateActualLimits = function (dataType, dataMin, dataMax, labels, values) {
                    var _this = this;
                    if (labels === void 0) { labels = null; }
                    if (values === void 0) { values = null; }
                    _super.prototype._updateActualLimits.call(this, dataType, dataMin, dataMax, labels, values);
                    var chart = this._chart, lbls = this._lbls, min = this.actualMin.valueOf ? this.actualMin.valueOf() : this.actualMin, max = this.actualMax.valueOf ? this.actualMax.valueOf() : this.actualMax, len;
                    if (this._lbls && this === chart.axisX) {
                        chart._angles = [];
                        if (this._isTimeAxis && this._lbls.length === 0) {
                            this._values.forEach(function (v) {
                                lbls.push(_this._formatValue(v));
                            });
                        }
                        len = lbls.length;
                        if (chart.totalAngle < 360) {
                            len -= 1;
                        }
                        lbls.forEach(function (v, i) {
                            var val = min + (i / len) * (max - min), angle = chart.startAngle + (i / len) * chart.totalAngle;
                            if (!isNaN(angle) && !isNaN(val)) {
                                chart._angles.push({
                                    value: _this.convert(val),
                                    angle: angle
                                });
                            }
                        });
                    }
                };
                FlexRadarAxis.prototype._updateActualLimitsByChartType = function (labels, min, max) {
                    var chart = this._chart, ctype = chart._getChartType();
                    if (ctype != chart_1.ChartType.Column && chart.totalAngle === 360) {
                        if (this.axisType === chart_1.AxisType.X) {
                            if (this._isTimeAxis) {
                                var len = (chart._xlabels.length || chart._xvals.length) - 1;
                                len = len < 1 ? 1 : len;
                                max += (max - min) / len;
                            }
                            else {
                                max += 1;
                            }
                        }
                    }
                    return { min: min, max: max };
                };
                /**
                 * Converts the specified value from data to pixel coordinates.
                 *
                 * @param val The data value to convert.
                 * @param maxValue The max value of the data, it's optional.
                 * @param minValue The min value of the data, it's optional.
                 */
                FlexRadarAxis.prototype.convert = function (val, maxValue, minValue) {
                    var max = maxValue == null ? this.actualMax : maxValue, min = minValue == null ? this.actualMin : minValue, chart = this._chart;
                    if (!chart) {
                        return NaN;
                    }
                    if (max == min) {
                        return 0;
                    }
                    if (this.axisType === chart_1.AxisType.X) {
                        if (chart.reversed) {
                            return (chart.startAngle - (val - min) / (max - min) * chart.totalAngle) * Math.PI / 180;
                        }
                        else {
                            return (chart.startAngle + (val - min) / (max - min) * chart.totalAngle) * Math.PI / 180;
                        }
                    }
                    else {
                        var base = this.logBase;
                        if (!base) {
                            return (val - min) / (max - min) * chart._radius;
                        }
                        else {
                            if (val <= 0) {
                                return NaN;
                            }
                            var maxl = Math.log(max / min);
                            return Math.log(val / min) / maxl * chart._radius;
                        }
                    }
                };
                FlexRadarAxis.prototype._renderLineAndTitle = function (engine) {
                    var chart = this._chart, lineClass = chart_1.FlexChart._CSS_LINE, 
                    //pie segment draw from 9 o'clock in IRenderEngine
                    startAngle = (chart.startAngle - 90) * Math.PI / 180, totalAngle = chart.totalAngle * Math.PI / 180, radius = chart._radius;
                    if (this.axisType === chart_1.AxisType.X && this.axisLine) {
                        engine.stroke = chart_1.FlexChart._FG;
                        if (chart._isPolar) {
                            startAngle = chart.reversed ? startAngle - totalAngle : startAngle;
                            engine.drawPieSegment(chart._center.x, chart._center.y, radius, startAngle, totalAngle, lineClass);
                        }
                        else {
                            this._renderPolygon(engine, radius, lineClass);
                        }
                    }
                };
                FlexRadarAxis.prototype._renderPolygon = function (engine, r, cls) {
                    var chart = this._chart, cAngles = chart._angles, angleLen = cAngles.length, showXMinor = chart.axisX.minorGrid, gXPoints = [], gYPoints = [];
                    cAngles.forEach(function (a, i) {
                        if (showXMinor && i > 0) {
                            var newP = chart._convertPoint(r, a.value - (a.value - cAngles[i - 1].value) / 2);
                            gXPoints.push(newP.x);
                            gYPoints.push(newP.y);
                        }
                        var p = chart._convertPoint(r, a.value);
                        gXPoints.push(p.x);
                        gYPoints.push(p.y);
                    });
                    if (chart.totalAngle < 360) {
                        gXPoints.push(chart._center.x);
                        gYPoints.push(chart._center.y);
                    }
                    else if (showXMinor && angleLen >= 2) {
                        //add last point
                        var newP = chart._convertPoint(r, cAngles[angleLen - 1].value - (cAngles[angleLen - 2].value - cAngles[angleLen - 1].value) / 2);
                        gXPoints.push(newP.x);
                        gYPoints.push(newP.y);
                    }
                    engine.drawPolygon(gXPoints, gYPoints, cls);
                };
                FlexRadarAxis.prototype._renderMinors = function (engine, ticks, isVert, isNear) {
                    var _this = this;
                    var chart = this._chart, glineClass = chart_1.FlexChart._CSS_GRIDLINE_MINOR, grid = this.minorGrid, cAngles = chart._angles, angleLen = cAngles.length, showXMinor = chart.axisX.minorGrid, 
                    //gXPoints = [], gYPoints = [],
                    gstroke = chart_1.FlexChart._FG, gth = this._GRIDLINE_WIDTH, 
                    //pie segment draw from 9 o'clock in IRenderEngine
                    startAngle = chart.startAngle * Math.PI / 180, totalAngle = chart.totalAngle * Math.PI / 180, tover = this._TICK_OVERLAP, tickMarks = this.minorTickMarks, hasTicks = true, angle;
                    this._vals.minor = ticks;
                    if (tickMarks == chart_1.TickMark.Outside) {
                        tover = 1;
                    }
                    else if (tickMarks == chart_1.TickMark.Inside) {
                        tover = -1;
                    }
                    else if (tickMarks == chart_1.TickMark.Cross) {
                        tover = 0;
                    }
                    else {
                        hasTicks = false;
                    }
                    if (this.axisType == chart_1.AxisType.Y) {
                        engine.stroke = gstroke;
                        engine.strokeWidth = gth;
                        ticks.forEach(function (val) {
                            var y = _this.convert(val), t;
                            if (grid) {
                                _this._renderYGridLine(engine, chart, y, glineClass);
                            }
                            ;
                            if (hasTicks) {
                                cAngles.forEach(function (a, i) {
                                    if (showXMinor && i > 0) {
                                        angle = a.value - (a.value - cAngles[i - 1].value) / 2;
                                        var newP = chart._convertPoint(y, angle);
                                        _this._drawMinorTickLength(engine, tover, angle, newP);
                                    }
                                    angle = a.value;
                                    var p = chart._convertPoint(y, angle);
                                    _this._drawMinorTickLength(engine, tover, angle, p);
                                });
                                if (showXMinor && angleLen >= 2) {
                                    //add last point
                                    angle = cAngles[angleLen - 1].value - (cAngles[angleLen - 2].value - cAngles[angleLen - 1].value) / 2;
                                    var newP = chart._convertPoint(y, angle);
                                    _this._drawMinorTickLength(engine, tover, angle, newP);
                                }
                            }
                        });
                    }
                    else {
                        engine.stroke = gstroke;
                        engine.strokeWidth = gth;
                        ticks.forEach(function (val) {
                            var x = _this.convert(val);
                            if (grid) {
                                _this._renderXGridLine(engine, chart, x, glineClass);
                            }
                            if (hasTicks) {
                            }
                        });
                    }
                };
                FlexRadarAxis.prototype._drawMinorTickLength = function (engine, tover, angle, pt) {
                    var th = this._TICK_HEIGHT, tickClass = chart_1.FlexChart._CSS_TICK_MINOR;
                    var x1 = 0.5 * (tover - 1) * th * Math.cos(angle);
                    var x2 = 0.5 * (1 + tover) * th * Math.cos(angle);
                    var y1 = 0.5 * (tover - 1) * th * Math.sin(angle);
                    var y2 = 0.5 * (1 + tover) * th * Math.sin(angle);
                    engine.drawLine(pt.x + x1, pt.y + y1, pt.x + x2, pt.y + y2, tickClass);
                };
                FlexRadarAxis.prototype._renderLabelsAndTicks = function (engine, index, val, sval, labelAngle, tickMarks, showLabel, t1, t2) {
                    this._points = [];
                    labelAngle = this.labelAngle || 0;
                    var hasLbl = true, chart = this._chart, labelPadding = this.labelPadding || 2, lblClass = chart_1.FlexChart._CSS_LABEL, glineClass = chart_1.FlexChart._CSS_GRIDLINE, tickClass = chart_1.FlexChart._CSS_TICK, tstroke = chart_1.FlexChart._FG, tth = this._TICK_WIDTH, has_gline = this.majorGrid, gstroke = chart_1.FlexChart._FG, gth = this._GRIDLINE_WIDTH, 
                    //pie segment draw from 9 o'clock in IRenderEngine
                    startAngle = chart.startAngle * Math.PI / 180, totalAngle = chart.totalAngle * Math.PI / 180, gXPoints = [], gYPoints = [];
                    if (this.axisType == chart_1.AxisType.Y) {
                        has_gline = val != this.actualMin && has_gline && val != this.actualMax;
                        var y = this.convert(val), point = chart._convertPoint(y, startAngle);
                        if (has_gline) {
                            engine.stroke = gstroke;
                            engine.strokeWidth = gth;
                            this._renderYGridLine(engine, chart, y, glineClass);
                        }
                        engine.stroke = tstroke;
                        engine.strokeWidth = tth;
                        if (showLabel) {
                            var lpt = new wijmo.Point(point.x - labelPadding - Math.abs(t1 - t2), point.y);
                            this._axisLabels.push({
                                val: val,
                                text: sval,
                                pos: lpt,
                                align: 2,
                                vAlign: 1,
                                labelAngle: labelAngle,
                                class: lblClass
                            });
                        }
                        if (tickMarks != chart_1.TickMark.None) {
                            if (hasLbl) {
                                engine.drawLine(point.x - t2, point.y, point.x - t1, point.y, tickClass);
                            }
                        }
                    }
                    else {
                        var x = this.convert(val);
                        //point = chart._convertPoint(chart._radius, x);
                        if (has_gline) {
                            engine.stroke = gstroke;
                            engine.strokeWidth = gth;
                            //engine.drawLine(chart._center.x, chart._center.y, point.x, point.y);
                            this._renderXGridLine(engine, chart, x, glineClass);
                        }
                        engine.stroke = tstroke;
                        engine.strokeWidth = tth;
                        if (showLabel) {
                            var lpt = chart._convertPoint(chart._radius + labelPadding, x), angle, valign, align;
                            if (chart._angles && chart._angles.length) {
                                angle = chart._angles[index].angle;
                            }
                            else {
                                angle = chart.startAngle + (this.actualMin + val) * chart.totalAngle / (this.actualMax - this.actualMin);
                            }
                            angle = angle % 360;
                            angle = angle >= 0 ? angle : angle + 360;
                            valign = this._getXLabelVAlign(angle);
                            align = this._getXLabelAlign(angle);
                            if (chart._isPolar) {
                                sval = this._formatValue(angle);
                            }
                            if (labelAngle > 0) {
                                if (labelAngle == 90) {
                                    chart_1.FlexChart._renderRotatedText(engine, sval, lpt, align, valign, lpt, labelAngle, lblClass);
                                }
                                else {
                                    chart_1.FlexChart._renderRotatedText(engine, sval, lpt, align, valign, lpt, labelAngle, lblClass);
                                }
                            }
                            else if (labelAngle < 0) {
                                if (labelAngle == -90) {
                                    chart_1.FlexChart._renderRotatedText(engine, sval, lpt, align, valign, lpt, labelAngle, lblClass);
                                }
                                else {
                                    chart_1.FlexChart._renderRotatedText(engine, sval, lpt, align, valign, lpt, labelAngle, lblClass);
                                }
                            }
                            else {
                                this._renderLabel(engine, val, sval, lpt, align, valign /*1*/, lblClass);
                            }
                        }
                    }
                    return hasLbl;
                };
                FlexRadarAxis.prototype._renderXGridLine = function (engine, chart, x, cls) {
                    var center = chart._center, point = chart._convertPoint(chart._radius, x);
                    engine.drawLine(center.x, center.y, point.x, point.y, cls);
                };
                FlexRadarAxis.prototype._renderYGridLine = function (engine, chart, y, cls) {
                    var cAngles = chart._angles, center = chart._center, startAngle = chart.startAngle * Math.PI / 180, totalAngle = chart.totalAngle * Math.PI / 180;
                    if (chart._isPolar) {
                        startAngle = (chart.startAngle - 90) * Math.PI / 180;
                        startAngle = chart.reversed ? startAngle - totalAngle : startAngle;
                        engine.drawPieSegment(center.x, center.y, y, startAngle, totalAngle, cls);
                    }
                    else {
                        this._renderPolygon(engine, y, cls);
                    }
                };
                FlexRadarAxis.prototype._getXLabelVAlign = function (angle) {
                    var vAlign = 1, chart = this._chart, startAngle = chart.startAngle, reversed = chart.reversed;
                    if (reversed) {
                        angle = (360 + startAngle + (startAngle % 360 - angle % 360)) % 360;
                    }
                    if (angle === 0) {
                        vAlign = 2;
                    }
                    else if (angle === 180) {
                        vAlign = 0;
                    }
                    return vAlign;
                };
                FlexRadarAxis.prototype._getXLabelAlign = function (angle) {
                    var align = 0, chart = this._chart, startAngle = chart.startAngle, reversed = chart.reversed;
                    if (reversed) {
                        angle = (360 + startAngle + (startAngle % 360 - angle % 360)) % 360;
                    }
                    if (angle > 0 && angle < 180) {
                        align = -1;
                    }
                    else if (angle > 180 && angle < 360) {
                        align = 1;
                    }
                    return align + 1;
                };
                FlexRadarAxis.prototype._createTimeLabels = function (start, len, vals, lbls) {
                    var _this = this;
                    if (this._axisType == chart_1.AxisType.Y) {
                        _super.prototype._createTimeLabels.call(this, start, len, vals, lbls);
                    }
                    else {
                        var values = this._values, fmt = this.format;
                        if (!values || values.length === 0) {
                            return;
                        }
                        values.forEach(function (v) {
                            vals.push(v);
                            lbls.push(_this._formatValue(v));
                        });
                    }
                };
                return FlexRadarAxis;
            }(chart_1.Axis));
            radar.FlexRadarAxis = FlexRadarAxis;
        })(radar = chart_1.radar || (chart_1.radar = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=FlexRadarAxis.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart_1) {
        var radar;
        (function (radar) {
            'use strict';
            /**
             * Line/scatter radar chart plotter.
             */
            var _RadarLinePlotter = (function (_super) {
                __extends(_RadarLinePlotter, _super);
                function _RadarLinePlotter() {
                    _super.apply(this, arguments);
                    this.isArea = false;
                }
                _RadarLinePlotter.prototype.plotSeries = function (engine, ax, ay, series, palette, iser, nser) {
                    var ser = wijmo.asType(series, chart_1.SeriesBase), chart = this.chart, chartType = ser._getChartType() || chart._getChartType(), si = chart.series.indexOf(series);
                    var ys = series.getValues(0);
                    var xs = series.getValues(1);
                    if (!ys) {
                        return;
                    }
                    if (!xs) {
                        xs = this.dataInfo.getXVals();
                    }
                    var style = chart_1._BasePlotter.cloneStyle(series.style, ['fill']);
                    var len = ys.length;
                    var hasXs = true;
                    if (!xs) {
                        hasXs = false;
                        xs = new Array(len);
                    }
                    else {
                        len = Math.min(len, xs.length);
                    }
                    var swidth = this._DEFAULT_WIDTH, fill = ser._getSymbolFill(si), altFill = ser._getAltSymbolFill(si) || fill, stroke = ser._getSymbolStroke(si), altStroke = ser._getAltSymbolStroke(si) || stroke, symSize = ser._getSymbolSize();
                    engine.stroke = stroke;
                    engine.strokeWidth = swidth;
                    engine.fill = fill;
                    var xvals = new Array();
                    var yvals = new Array();
                    var stacked = this.stacking != chart_1.Stacking.None && !ser._isCustomAxisY();
                    var stacked100 = this.stacking == chart_1.Stacking.Stacked100pc && !ser._isCustomAxisY();
                    if (ser._getChartType() !== undefined) {
                        stacked = stacked100 = false;
                    }
                    var interpolateNulls = this.chart.interpolateNulls;
                    var hasNulls = false;
                    for (var i = 0; i < len; i++) {
                        var datax = hasXs ? xs[i] : i;
                        var datay = ys[i];
                        if (chart_1._DataInfo.isValid(datax) && chart_1._DataInfo.isValid(datay)) {
                            if (stacked) {
                                if (stacked100) {
                                    var sumabs = this.dataInfo.getStackedAbsSum(datax);
                                    datay = datay / sumabs;
                                }
                                if (datay >= 0) {
                                    var sum = isNaN(this.stackPos[datax]) ? 0 : this.stackPos[datax];
                                    datay = this.stackPos[datax] = sum + datay;
                                }
                                else {
                                    var sum = isNaN(this.stackNeg[datax]) ? 0 : this.stackNeg[datax];
                                    datay = this.stackNeg[datax] = sum + datay;
                                }
                            }
                            var dpt;
                            dpt = new chart_1._DataPoint(si, i, datax, datay);
                            var angle = ax.convert(datax), radius = ay.convert(datay), point = this.chart._convertPoint(radius, angle);
                            datax = point.x;
                            datay = point.y;
                            if (!isNaN(datax) && !isNaN(datay)) {
                                xvals.push(datax);
                                yvals.push(datay);
                                var area = new chart_1._CircleArea(new wijmo.Point(datax, datay), 0.5 * symSize);
                                area.tag = dpt;
                                this.hitTester.add(area, si);
                            }
                            else {
                                hasNulls = true;
                                if (interpolateNulls !== true) {
                                    xvals.push(undefined);
                                    yvals.push(undefined);
                                }
                            }
                        }
                        else {
                            hasNulls = true;
                            if (interpolateNulls !== true) {
                                xvals.push(undefined);
                                yvals.push(undefined);
                            }
                        }
                    }
                    var itemIndex = 0;
                    if (this.hasLines) {
                        if (this.isArea) {
                            engine.fill = palette._getColorLight(si);
                        }
                        else {
                            engine.fill = 'none';
                        }
                        if (hasNulls && interpolateNulls !== true) {
                            var dx = [];
                            var dy = [];
                            for (var i = 0; i < len; i++) {
                                if (xvals[i] === undefined) {
                                    dx.push(undefined);
                                    dy.push(0);
                                }
                                else {
                                    dx.push(xvals[i]);
                                    dy.push(yvals[i]);
                                }
                            }
                            if (dx.length > 1) {
                                if (chart._isPolar && chartType !== chart_1.ChartType.Area) {
                                    this._drawLines(engine, dx, dy, null, style, this.chart._plotrectId);
                                }
                                else {
                                    if (chart.totalAngle < 360) {
                                        dx.push(chart._center.x);
                                        dy.push(chart._center.y);
                                    }
                                    engine.drawPolygon(dx, dy, null, style, this.chart._plotrectId);
                                }
                                //this._drawLines(engine, dx, dy, null, style, this.chart._plotrectId);
                                this.hitTester.add(new chart_1._LinesArea(dx, dy), si);
                                itemIndex++;
                            }
                        }
                        else {
                            if (chart._isPolar && chartType !== chart_1.ChartType.Area) {
                                this._drawLines(engine, xvals, yvals, null, style, this.chart._plotrectId);
                            }
                            else {
                                if (chart.totalAngle < 360) {
                                    xvals.push(chart._center.x);
                                    yvals.push(chart._center.y);
                                }
                                engine.drawPolygon(xvals, yvals, null, style, this.chart._plotrectId);
                            }
                            //this._drawLines(engine, xvals, yvals, null, style, this.chart._plotrectId);
                            this.hitTester.add(new chart_1._LinesArea(xvals, yvals), si);
                            itemIndex++;
                        }
                    }
                    engine.fill = fill;
                    for (var i = 0; i < len; i++) {
                        var datax = xvals[i];
                        var datay = yvals[i];
                        // scatter fill/stroke
                        if (this.hasLines === false) {
                            engine.fill = ys[i] > 0 ? fill : altFill;
                            engine.stroke = ys[i] > 0 ? stroke : altStroke;
                        }
                        //if (DataInfo.isValid(datax) && DataInfo.isValid(datay)) {
                        if (this.isValid(datax, datay, ax, ay)) {
                            if ((this.hasSymbols || this.chart.itemFormatter) && symSize > 0) {
                                this._drawSymbol(engine, datax, datay, symSize, ser, i);
                            }
                            series._setPointIndex(i, itemIndex);
                            itemIndex++;
                        }
                    }
                };
                return _RadarLinePlotter;
            }(chart_1._LinePlotter));
            radar._RadarLinePlotter = _RadarLinePlotter;
        })(radar = chart_1.radar || (chart_1.radar = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_RadarLinePlotter.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var chart;
    (function (chart_1) {
        var radar;
        (function (radar) {
            'use strict';
            /**
             * Column(Rose) radar chart plotter.
             */
            var _RadarBarPlotter = (function (_super) {
                __extends(_RadarBarPlotter, _super);
                function _RadarBarPlotter() {
                    _super.apply(this, arguments);
                }
                _RadarBarPlotter.prototype.adjustLimits = function (dataInfo, plotRect) {
                    //return super.adjustLimits(dataInfo, plotRect);
                    this.dataInfo = dataInfo;
                    var xmin = dataInfo.getMinX();
                    var ymin = dataInfo.getMinY();
                    var xmax = dataInfo.getMaxX();
                    var ymax = dataInfo.getMaxY();
                    var dx = dataInfo.getDeltaX();
                    if (dx <= 0) {
                        dx = 1;
                    }
                    if (this.chart.totalAngle < 360) {
                        dx = 0;
                    }
                    this.unload();
                    if (!this.chart.axisY.logBase) {
                        if (this.origin > ymax) {
                            ymax = this.origin;
                        }
                        else if (this.origin < ymin) {
                            ymin = this.origin;
                        }
                    }
                    return new wijmo.Rect(xmin, ymin, xmax - xmin + dx, ymax - ymin);
                };
                _RadarBarPlotter.prototype.plotSeries = function (engine, ax, ay, series, palette, iser, nser) {
                    var si = this.chart.series.indexOf(series);
                    var ser = wijmo.asType(series, chart_1.SeriesBase);
                    var options = this.chart.options;
                    var cw = this.width;
                    var wpx = 0;
                    var chart = this.chart;
                    var startAngle = -90 * Math.PI / 180;
                    iser = iser || 0;
                    var axid = ser._getAxisY()._uniqueId;
                    var area;
                    var stackNeg = this.stackNegMap[axid];
                    var stackPos = this.stackPosMap[axid];
                    var stacked = this.stacking != chart_1.Stacking.None;
                    var stacked100 = this.stacking == chart_1.Stacking.Stacked100pc;
                    var yvals = series.getValues(0);
                    var xvals = series.getValues(1);
                    if (!yvals) {
                        return;
                    }
                    if (!xvals) {
                        xvals = this.dataInfo.getXVals();
                    }
                    // find minimal distance between point and use it as column width
                    var delta;
                    if (xvals) {
                        delta = chart.totalAngle / xvals.length;
                    }
                    else {
                        delta = chart.totalAngle / (ax.actualMax - ax.actualMin);
                    }
                    if (delta > 0) {
                        if (stacked) {
                            cw = delta * cw * Math.PI / 180;
                        }
                        else {
                            cw = delta * Math.pow(cw, iser + 1) * Math.PI / 180;
                        }
                    }
                    // set series fill and stroke from style
                    var fill = ser._getSymbolFill(si), altFill = ser._getAltSymbolFill(si) || fill, stroke = ser._getSymbolStroke(si), altStroke = ser._getAltSymbolStroke(si) || stroke;
                    var len = yvals.length;
                    if (xvals != null) {
                        len = Math.min(len, xvals.length);
                    }
                    var origin = this.origin;
                    var itemIndex = 0, currentFill, currentStroke;
                    if (ser._getChartType() !== undefined) {
                        stacked = stacked100 = false;
                    }
                    if (origin < ay.actualMin) {
                        origin = ay.actualMin;
                    }
                    else if (origin > ay.actualMax) {
                        origin = ay.actualMax;
                    }
                    var originScreen = ay.convert(origin), xmin = ax.actualMin, xmax = ax.actualMax;
                    if (ser._isCustomAxisY()) {
                        stacked = stacked100 = false;
                    }
                    if (!chart._areas[si]) {
                        chart._areas[si] = [];
                    }
                    for (var i = 0; i < len; i++) {
                        var datax = xvals ? xvals[i] : i;
                        var datay = yvals[i];
                        if (this._getSymbolOrigin) {
                            originScreen = ay.convert(this._getSymbolOrigin(origin, i, len));
                        }
                        if (this._getSymbolStyles) {
                            var style = this._getSymbolStyles(i, len);
                            fill = style && style.fill ? style.fill : fill;
                            altFill = style && style.fill ? style.fill : altFill;
                            stroke = style && style.stroke ? style.stroke : stroke;
                            altStroke = style && style.stroke ? style.stroke : altStroke;
                        }
                        // apply fill and stroke
                        currentFill = datay > 0 ? fill : altFill;
                        currentStroke = datay > 0 ? stroke : altStroke;
                        engine.fill = currentFill;
                        engine.stroke = currentStroke;
                        if (chart_1._DataInfo.isValid(datax) && chart_1._DataInfo.isValid(datay)) {
                            if (stacked) {
                                var x0 = datax - 0.5 * cw, x1 = datax + 0.5 * cw;
                                if ((x0 < xmin && x1 < xmin) || (x0 > xmax && x1 > xmax)) {
                                    continue;
                                }
                                x0 = ax.convert(x0);
                                x1 = ax.convert(x1);
                                if (!chart_1._DataInfo.isValid(x0) || !chart_1._DataInfo.isValid(x1)) {
                                    continue;
                                }
                                var y0, y1;
                                if (stacked100) {
                                    var sumabs = this.dataInfo.getStackedAbsSum(datax);
                                    datay = datay / sumabs;
                                }
                                var sum = isNaN(stackPos[datax]) ? 0 : stackPos[datax];
                                y0 = sum;
                                y1 = sum + datay;
                                stackPos[datax] = sum + datay;
                                var angle = ax.convert(datax), radius0 = ay.convert(y0), radius1 = ay.convert(y1);
                                angle = angle - cw / 2;
                                engine.drawDonutSegment(chart._center.x, chart._center.y, radius1, radius0, angle + startAngle, cw, null, ser.symbolStyle);
                                area = new chart_1._DonutSegment(new wijmo.Point(chart._center.x, chart._center.y), radius1, radius0, angle + startAngle, cw);
                                area.tag = new chart_1._DataPoint(si, i, datax, sum + datay);
                                this.hitTester.add(area, si);
                            }
                            else {
                                var angle = ax.convert(datax), radius = ay.convert(datay), p = chart._convertPoint(radius, angle);
                                angle = angle - cw / 2;
                                engine.drawPieSegment(chart._center.x, chart._center.y, radius, angle + startAngle, cw, null, ser.symbolStyle);
                                area = new chart_1._PieSegment(new wijmo.Point(chart._center.x, chart._center.y), radius, angle + startAngle, cw);
                                area.tag = new chart_1._DataPoint(si, i, datax, datay);
                                this.hitTester.add(area, si);
                            }
                            chart._areas[si].push(area);
                            series._setPointIndex(i, itemIndex);
                            itemIndex++;
                        }
                    }
                };
                return _RadarBarPlotter;
            }(chart_1._BarPlotter));
            radar._RadarBarPlotter = _RadarBarPlotter;
        })(radar = chart_1.radar || (chart_1.radar = {}));
    })(chart = wijmo.chart || (wijmo.chart = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_RadarBarPlotter.js.map
