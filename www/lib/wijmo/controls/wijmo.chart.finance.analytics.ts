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
 * Analytics extensions for @see:FinancialChart.
 */
module wijmo.chart.finance.analytics {
    'use strict';

    // internal helper function to validate that a number is truly a number (and not Infinity, NaN, etc.)
    function isValid(value: number): boolean {
        return isFinite(value) && !isNaN(value) && isNumber(value);
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
    export class Fibonacci extends SeriesBase {
        private _high: number;
        private _low: number;
        private _minX: any;
        private _maxX: any;
        private _actualHigh: number;
        private _actualLow: number;
        private _levels: number[] = [0, 23.6, 38.2, 50, 61.8, 100];
        private _uptrend = true;
        private _labelPosition: LabelPosition = LabelPosition.Left;

        /**
         * Initializes a new instance of the @see:Fibonacci class.
         *
         * @param options A JavaScript object containing initialization data.
         */
        constructor(options?) {
            super();
            if (options) {
                wijmo.copy(this, options);
            }
            this.rendering.addHandler(this._render);
        }

        /**
         * Gets or sets the low value of @see:Fibonacci tool.
         *
         * If not specified, the low value is calculated based on data values provided by <b>itemsSource</b>.
         */
        get low(): number {
            return this._low;
        }
        set low(value: number) {
            if (value != this._low) {
                this._low = asNumber(value, true);
                this._invalidate();
            }
        }

        /**
         * Gets or sets the high value of @see:Fibonacci tool.
         *
         * If not specified, the high value is caclulated based on
         * data values provided by the <b>itemsSource</b>.
         */
        get high(): number {
            return this._high;
        }
        set high(value: number) {
            if (value != this._high) {
                this._high = asNumber(value, true);
                this._invalidate();
            }
        }

        /**
         * Gets or sets the label position for levels in @see:Fibonacci tool.
         */
        get labelPosition(): LabelPosition {
            return this._labelPosition;
        }
        set labelPosition(value: LabelPosition) {
            if (value != this._labelPosition) {
                this._labelPosition = asEnum(value, LabelPosition, true);
                this._invalidate();
            }
        }

        /**
         * Gets or sets a value indicating whether to create uptrending @see:Fibonacci tool.
         *
         * Default value is true(uptrend). If the value is false, the downtrending levels are plotted.
         */
        get uptrend(): boolean {
            return this._uptrend;
        }
        set uptrend(value: boolean) {
            if (value != this._uptrend) {
                this._uptrend = asBoolean(value, true);
                this._invalidate();
            }
        }

        /**
         * Gets or sets the array of levels for plotting.
         *
         * Default value is [0, 23.6, 38.2, 50, 61.8, 100].
         */
        get levels(): number[] {
            return this._levels;
        }
        set levels(value: number[]) {
            if (value != this._levels) {
                this._levels = asArray(value, true);
                this._invalidate();
            }
        }

        /**
         * Gets or sets the x minimal value of the @see:Fibonacci tool.
         *
         * If not specified, current minimum of x-axis is used.
         * The value can be specified as a number or Date object.
         */
        get minX(): any {
            return this._minX;
        }
        set minX(value: any) {
            if (value != this._minX) {
                this._minX = value;
                this._invalidate();
            }
        }

        /**
         * Gets or sets the x maximum value of the @see:Fibonacci tool.
         *
         * If not specified, current maximum of x-axis is used.
         * The value can be specified as a number or Date object.
         */
        get maxX(): any {
            return this._maxX;
        }
        set maxX(value: any) {
            if (value != this._maxX) {
                this._maxX = value;
                this._invalidate();
            }
        }

        private _getMinX(): number {
            if (isNumber(this._minX)) {
                return this._minX;
            } else if (isDate(this._minX)) {
                return asDate(this._minX).valueOf();
            } else {
                return this._getAxisX().actualMin;
            }
        }

        private _getMaxX(): number {
            if (isNumber(this._maxX)) {
                return this._maxX;
            } else if (isDate(this._maxX)) {
                return asDate(this._maxX).valueOf();
            } else {
                return this._getAxisX().actualMax;
            }
        }

        private _updateLevels(): void {
            var min = undefined,
                max = undefined;
            if (this._low === undefined || this._high === undefined) {
                var vals = super.getValues(0);
                var xvals = super.getValues(1);
                if (vals) {
                    var len = vals.length;
                    var xmin = this._getMinX(),
                        xmax = this._getMaxX();

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
            } else {
                this._actualLow = this._low;
            }

            if (this._high === undefined && max !== undefined) {
                this._actualHigh = max;
            } else {
                this._actualHigh = this._high;
            }
        }

        private _render(sender: SeriesBase, args: RenderEventArgs): void {
            var ser = <Fibonacci>sender;
            ser._updateLevels();

            var ax = ser._getAxisX();
            var ay = ser._getAxisY();
            var eng = args.engine;

            var swidth = 2,
                stroke = ser._getSymbolStroke(ser._chart.series.indexOf(ser));

            var lstyle = _BasePlotter.cloneStyle(ser.style, ['fill']);
            var tstyle = _BasePlotter.cloneStyle(ser.style, ['stroke']);
            var clipPath = ser.chart._plotrectId;

            eng.stroke = stroke;
            eng.strokeWidth = swidth;
            eng.textFill = stroke;

            var xmin = ser._getMinX(),
                xmax = ser._getMaxX();

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
                var x1 = ax.convert(xmin),
                    x2 = ax.convert(xmax);
                var y = ser.uptrend ?
                    ay.convert(ser._actualLow + 0.01 * lvl * (ser._actualHigh - ser._actualLow)) :
                    ay.convert(ser._actualHigh - 0.01 * lvl * (ser._actualHigh - ser._actualLow));

                if (_DataInfo.isValid(x1) && _DataInfo.isValid(x2) && _DataInfo.isValid(y)) {
                    eng.drawLine(x1, y, x2, y, null, lstyle);

                    if (ser.labelPosition != LabelPosition.None) {
                        var s = lvl.toFixed(1) + '%';
                        var va = 0;
                        if ((ser.uptrend && i == 0) || (!ser.uptrend && i == llen - 1)) {
                            va = 2;
                        }

                        switch (ser.labelPosition) {
                            case LabelPosition.Left:
                                FlexChartCore._renderText(eng, s, new Point(x1, y), 0, va, null, null, tstyle);
                                break;
                            case LabelPosition.Center:
                                FlexChartCore._renderText(eng, s, new Point(0.5 * (x1 + x2), y), 1, va, null, null, tstyle);
                                break;
                            case LabelPosition.Right:
                                FlexChartCore._renderText(eng, s, new Point(x2, y), 2, va, null, null, tstyle);
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
        }

        _getChartType(): ChartType {
            return ChartType.Line;
        }
    }

    /**
     * Represents a Fibonacci Arcs tool for the @see:FinancialChart.
     */
    export class FibonacciArcs extends SeriesBase {
        private _start: DataPoint;
        private _end: DataPoint;
        private _levels: number[] = [38.2, 50, 61.8];
        private _labelPosition: LabelPosition = LabelPosition.Top;

        /**
         * Initializes a new instance of the @see:FibonacciArcs class.
         *
         * @param options A JavaScript object containing initialization data.
         */
        constructor(options?: any) {
            super();

            // copy options
            if (options) {
                wijmo.copy(this, options);
            }

            this.rendering.addHandler(this._render, this);
        }

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
        get start(): DataPoint {
            return this._start;
        }
        set start(value: DataPoint) {
            if (value !== this.start) {
                this._start = asType(value, DataPoint);
                this._invalidate();
            }
        }

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
        get end(): DataPoint {
            return this._end;
        }
        set end(value: DataPoint) {
            if (value !== this.end) {
                this._end = asType(value, DataPoint);
                this._invalidate();
            }
        }

        /**
         * Gets or sets the array of levels for plotting.
         *
         * Default value is [38.2, 50, 61.8].
         */
        get levels(): number[] {
            return this._levels;
        }
        set levels(value: number[]) {
            if (value !== this._levels) {
                this._levels = asArray(value, false);
                this._invalidate();
            }
        }

        /**
         * Gets or sets the @see:LabelPosition for levels in @see:FibonacciArcs tool.
         */
        get labelPosition(): LabelPosition {
            return this._labelPosition;
        }
        set labelPosition(value: LabelPosition) {
            if (value !== this.labelPosition) {
                this._labelPosition = asEnum(value, LabelPosition);
                this._invalidate();
            }
        }

        _render(sender: SeriesBase, args: RenderEventArgs): void {
            var startX = this._getX(0),
                startY = this._getY(0),
                endX = this._getX(1),
                endY = this._getY(1);

            if (super._getLength() <= 1 || !isValid(startX) || !isValid(startY) || !isValid(endX) || !isValid(endY)) {
                return;
            }

            var ax = this._getAxisX(),
                ay = this._getAxisY(),
                engine = args.engine,
                swidth = 2,
                group,
                si = this.chart.series.indexOf(this),
                stroke = this._getSymbolStroke(si),
                lstyle = _BasePlotter.cloneStyle(this.style, ["fill"]),
                tstyle = _BasePlotter.cloneStyle(this.style, ["stroke"]);

            engine.stroke = stroke;
            engine.strokeWidth = swidth;
            engine.textFill = stroke;

            var clipPath = this.chart._plotrectId,
                yDiff = endY - startY,
                cx: number, cy: number, acy: number,
                baseLen: number, radius: number,
                center: Point, lvl: number,
                size: Size, lbl: string;

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
            center = new Point(endX, endY);

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
                    if (this.labelPosition !== LabelPosition.None && lvl !== 0) {
                        // get label and determine its size
                        lbl = Globalize.format(lvl, "p1");
                        size = engine.measureString(lbl, null, null, tstyle);

                        // get label's y position
                        acy = yDiff <= 0 ? cy - radius : cy + radius;
                        switch (this.labelPosition) {
                            case LabelPosition.Center:
                                acy += (size.height * 0.5);
                                break;
                            case LabelPosition.Bottom:
                                acy += yDiff <= 0 ? size.height : 0;
                                break;
                            default:
                                acy += yDiff <= 0 ? 0 : size.height;
                                break;
                        }

                        engine.drawString(lbl, new Point(cx - size.width * .5, acy), null, tstyle);
                    }
                }
            }

            engine.stroke = null;
            engine.strokeWidth = null;
            engine.textFill = null;
            // end group
            engine.endGroup();
        }

        private _getX(dim: number): number {
            var retval = null;

            if (dim === 0 && this.start) {
                retval = this.start.x;
            } else if (dim === 1 && this.end) {
                retval = this.end.x;
            }

            if (isDate(retval)) {
                retval = asDate(retval).valueOf();
            }

            return retval;
        }

        private _getY(dim: number): number {
            var retval = null;

            if (dim === 0 && this.start) {
                retval = this.start.y;
            } else if (dim === 1 && this.end) {
                retval = this.end.y;
            }

            return retval;
        }

        _getChartType(): ChartType {
            return ChartType.Line;
        }
    }

    /**
     * Represents a Fibonacci Fans tool for the @see:FinancialChart.
     */
    export class FibonacciFans extends SeriesBase {
        private _start: DataPoint;
        private _end: DataPoint;
        private _levels: number[] = [0, 23.6, 38.2, 50, 61.8, 100];
        private _labelPosition: LabelPosition = LabelPosition.Top;

        /**
         * Initializes a new instance of the @see:FibonacciFans class.
         *
         * @param options A JavaScript object containing initialization data.
         */
        constructor(options?: any) {
            super();

            // copy options
            if (options) {
                wijmo.copy(this, options);
            }

            this.rendering.addHandler(this._render, this);
        }

        /**
         * Gets or sets the starting @see:DataPoint for the base line.
         *
         * If not set, the starting @see:DataPoint is calculated automatically.
         * The @see:DataPoint x value can be a number or a Date object (for
         * time-based data).
         */
        get start(): DataPoint {
            return this._start;
        }
        set start(value: DataPoint) {
            if (value !== this.start) {
                this._start = asType(value, DataPoint);
                this._invalidate();
            }
        }

        /**
         * Gets or sets the ending @see:DataPoint for the base line.
         *
         * If not set, the starting @see:DataPoint is calculated automatically.
         * The @see:DataPoint x value can be a number or a Date object (for
         * time-based data).
         */
        get end(): DataPoint {
            return this._end;
        }
        set end(value: DataPoint) {
            if (value !== this.end) {
                this._end = asType(value, DataPoint);
                this._invalidate();
            }
        }

        /**
         * Gets or sets the array of levels for plotting.
         *
         * Default value is [0, 23.6, 38.2, 50, 61.8, 100].
         */
        get levels(): number[] {
            return this._levels;
        }
        set levels(value: number[]) {
            if (value !== this._levels) {
                this._levels = asArray(value, false);
                this._invalidate();
            }
        }

        /**
         * Gets or sets the @see:LabelPosition for levels in @see:FibonacciFans tool.
         */
        get labelPosition(): LabelPosition {
            return this._labelPosition;
        }
        set labelPosition(value: LabelPosition) {
            if (value !== this.labelPosition) {
                this._labelPosition = asEnum(value, LabelPosition);
                this._invalidate();
            }
        }

        _updateLevels(): void {
            // both must be defined, otherwise we calulate start/end automatically
            if (!this.start || !this.end) {
                var plotter = this.chart._getPlotter(this),
                    ax = this._getAxisX(),
                    yvals = super.getValues(0),
                    xvals = super.getValues(1) || plotter.dataInfo.getXVals(),
                    xmin: number, xmax: number,
                    ymin: number, ymax: number;

                // use yvals only - no axisY.[actualMin|actualMax]
                if (yvals && yvals.length > 0) {
                    ymin = _minimum(yvals);
                    ymax = _maximum(yvals);
                }

                if (xvals && xvals.length > 0) {
                    xmin = _minimum(xvals);
                    xmax = _maximum(xvals);
                } else {
                    xmin = ax.actualMin;
                    xmax = ax.actualMax;
                }

                if (isValid(xmin) && isValid(ymin) && isValid(xmax) && isValid(ymax)) {
                    this.start = new DataPoint(xmin, ymin);
                    this.end = new DataPoint(xmax, ymax);
                }
            }
        }

        _render(sender: SeriesBase, args: RenderEventArgs): void {
            this._updateLevels();

            var startX = this._getX(0),
                startY = this._getY(0),
                endX = this._getX(1),
                endY = this._getY(1);

            if (super._getLength() <= 1 || !isValid(startX) || !isValid(startY) || !isValid(endX) || !isValid(endY)) {
                return;
            }

            var ax = this._getAxisX(),
                ay = this._getAxisY(),
                si = this.chart.series.indexOf(this),
                engine = args.engine,
                swidth = 2,
                stroke = this._getSymbolStroke(si),
                lstyle = _BasePlotter.cloneStyle(this.style, ["fill"]),
                tstyle = _BasePlotter.cloneStyle(this.style, ["stroke"]);

            engine.stroke = stroke;
            engine.strokeWidth = swidth;
            engine.textFill = stroke;

            var yDiff = endY - startY,
                xDiff = endX - startX,
                clipPath = this.chart._plotrectId,
                x1: number, x2: number,
                y1: number, y2: number,
                pt1: Point, pt2: Point, cp: Point,
                m: number, b: number,
                lvl: number, lbl: string,
                size: Size, angle: number;

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
                } else if (yDiff < 0 && y2 < ay.actualMin) {
                    y2 = ay.actualMin;
                    x2 = (y2 - b) / m;
                }

                if (isValid(x1) && isValid(y1) && isValid(x2) && isValid(y2)) {
                    // convert once per fan line & associated label
                    pt1 = new Point(ax.convert(x1), ay.convert(y1));
                    pt2 = new Point(ax.convert(x2), ay.convert(y2));

                    // draw fan line
                    engine.drawLines([pt1.x, pt2.x], [pt1.y, pt2.y], null, lstyle);

                    // draw fan label
                    if (this.labelPosition != LabelPosition.None) {

                        // get label and determine its size
                        lbl = Globalize.format(lvl, "p1");
                        size = engine.measureString(lbl, null, null, tstyle);

                        // find angle for label
                        angle = Math.atan((pt2.y - pt1.y) / (pt2.x - pt1.x)) * 180 / Math.PI;

                        // get center point by cloning the label point
                        cp = pt2.clone();

                        // update label point for axis boundx
                        pt2.x = xDiff > 0 ? pt2.x - size.width : pt2.x;

                        var a = angle * Math.PI / 180,
                            tl = new Point(), bl = new Point(),
                            tr = new Point(), br = new Point(),
                            ymin = ay.convert(ay.actualMin), ymax = ay.convert(ay.actualMax),
                            xmin = ax.convert(ax.actualMin), xmax = ax.convert(ax.actualMax),
                            limit: number, acp = cp.clone();

                        // adjust pt2.y based on label position property
                        switch (this.labelPosition) {
                            // top is the default by nature
                            case LabelPosition.Center:
                                pt2.y += size.height * 0.5;

                                // todo: this works okay, but corners should be calculated in this case
                                acp.y += size.height * 0.5;
                                break;
                            case LabelPosition.Bottom:
                                pt2.y += size.height;
                                break;
                        }

                        // http://math.stackexchange.com/questions/170650/how-to-get-upper-left-upper-right-lower-left-and-lower-right-corners-xy-coordi
                        // attempt to keep labels in bounds
                        if (xDiff > 0) {
                            // todo: center is slightly off because the corners aren't correct
                            // calculate coordinates of label's corners
                            if (this.labelPosition === LabelPosition.Top || this.labelPosition === LabelPosition.Center) {
                                br = acp.clone();

                                tr.x = br.x + size.height * Math.sin(a);
                                tr.y = br.y - size.height * Math.cos(a);

                                tl.x = br.x - size.width * Math.cos(a) + size.height * Math.sin(a);
                                tl.y = br.y - size.width * Math.sin(a) - size.height * Math.cos(a);

                                bl.x = br.x - size.width * Math.cos(a);
                                bl.y = br.y - size.width * Math.sin(a);
                            } else if (this.labelPosition === LabelPosition.Bottom) {
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
                            } else if (yDiff < 0) {
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
                        } else if (xDiff < 0) {
                            // todo: center is slightly off because the corners aren't correct
                            if (this.labelPosition === LabelPosition.Top || this.labelPosition === LabelPosition.Center) {
                                bl = acp.clone();

                                tl.x = bl.x + size.height * Math.sin(a);
                                tl.y = bl.y - size.height * Math.cos(a);

                                br.x = bl.x + size.width * Math.cos(a);
                                br.y = bl.y + size.width * Math.sin(a);

                                tr.x = tl.x + size.width * Math.cos(a);
                                tr.y = tl.y + size.width * Math.sin(a);
                            } else if (this.labelPosition === LabelPosition.Bottom) {
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
                            } else if (yDiff < 0) {
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
                        } else {
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
        }

        private _getX(dim: number): number {
            var retval = null;

            if (dim === 0 && this.start) {
                retval = this.start.x;
            } else if (dim === 1 && this.end) {
                retval = this.end.x;
            }

            if (isDate(retval)) {
                retval = asDate(retval).valueOf();
            }

            return retval;
        }

        private _getY(dim: number): number {
            var retval = null;

            if (dim === 0 && this.start) {
                retval = this.start.y;
            } else if (dim === 1 && this.end) {
                retval = this.end.y;
            }

            return retval;
        }

        _getChartType(): ChartType {
            return ChartType.Line;
        }
    }

    /**
     * Represents a Fibonacci Time Zones tool for the @see:FinancialChart.
     */
    export class FibonacciTimeZones extends SeriesBase {
        private _startX: any;
        private _endX: any;
        private _levels: number[] = [0, 1, 2, 3, 5, 8, 13, 21, 34];
        private _labelPosition: LabelPosition = LabelPosition.Right;

        /**
         * Initializes a new instance of the @see:FibonacciTimeZones class.
         *
         * @param options A JavaScript object containing initialization data.
         */
        constructor(options?: any) {
            super();

            // copy options
            if (options) {
                wijmo.copy(this, options);
            }

            this.rendering.addHandler(this._render, this);
        }

        /**
         * Gets or sets the starting X data point for the time zones.
         *
         * If not set, the starting X data point is calculated automatically. The
         * value can be a number or a Date object (for time-based data).
         */
        get startX(): any {
            return this._startX;
        }
        set startX(value: any) {
            if (value !== this.startX) {
                if (isDate(value)) {
                    this._startX = asDate(value);
                } else {
                    this._startX = asNumber(value);
                }
                this._invalidate();
            }
        }

        /**
         * Gets or sets the ending X data point for the time zones.
         *
         * If not set, the ending X data point is calculated automatically. The
         * value can be a number or a Date object (for time-based data).
         */
        get endX(): any {
            return this._endX;
        }
        set endX(value: any) {
            if (value !== this.endX) {
                if (isDate(value)) {
                    this._endX = asDate(value);
                } else {
                    this._endX = asNumber(value);
                }
                this._invalidate();
            }
        }

        /**
         * Gets or sets the array of levels for plotting.
         *
         * Default value is [0, 1, 2, 3, 5, 8, 13, 21, 34].
         */
        get levels(): number[] {
            return this._levels;
        }
        set levels(value: number[]) {
            if (value !== this._levels) {
                this._levels = asArray(value, false);
                this._invalidate();
            }
        }

        /**
         * Gets or sets the @see:LabelPosition for levels in @see:FibonacciTimeZones tool.
         */
        get labelPosition(): LabelPosition {
            return this._labelPosition;
        }
        set labelPosition(value: LabelPosition) {
            if (value !== this.labelPosition) {
                this._labelPosition = asEnum(value, LabelPosition);
                this._invalidate();
            }
        }

        _render(sender: SeriesBase, args: RenderEventArgs): void {
            this._updateLevels();

            var start = this._getX(0),
                end = this._getX(1);

            if (super._getLength() <= 1 || !isValid(start) || !isValid(end)) {
                return;
            }

            var diff = end - start,
                ax = this._getAxisX(),
                ay = this._getAxisY(),
                si = this._chart.series.indexOf(this),
                engine = args.engine,
                swidth = 2,
                stroke = this._getSymbolStroke(si),
                lstyle = _BasePlotter.cloneStyle(this.style, ["fill"]),
                tstyle = _BasePlotter.cloneStyle(this.style, ["stroke"]),
                ymin = ay.convert(ay.actualMin),
                ymax = ay.convert(ay.actualMax),
                lvl: number, x: number,
                size: Size, lbl: string,
                clipPath = this.chart._plotrectId;

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
                if (this.labelPosition !== LabelPosition.None) {
                    // get label and determine its size
                    lbl = Globalize.format(lvl, "n0");
                    size = engine.measureString(lbl, null, null, tstyle);

                    // get label's x position
                    switch (this.labelPosition) {
                        case LabelPosition.Left:
                            x -= size.width + swidth;
                            break;
                        case LabelPosition.Center:
                            x -= size.width / 2;
                            break;
                        case LabelPosition.Right:
                            x += swidth;
                            break;
                        default:
                            x = diff < 0 ? x - size.width - swidth : x + swidth;
                            break;
                    }

                    engine.drawString(lbl, new Point(x, ymin), null, tstyle);
                }
            }

            engine.stroke = null;
            engine.strokeWidth = null;
            engine.textFill = null;
            // end group
            engine.endGroup();
        }

        _updateLevels(): void {
            var plotter = this.chart._getPlotter(this),
                xvals = super.getValues(1) || plotter.dataInfo.getXVals();

            if (super._getLength() <= 1) {
                return;
            }

            // get startX & endX as numbers; both must be define or both are ignored
            var start = this._getX(0),
                end = this._getX(1),
                defined = isNumber(start) && isNumber(end);

            // automatically init startX & endX if not defined
            if (!defined && !xvals) {
                this._startX = 0;
                this._endX = 1;
            } else if (!defined && xvals) {
                this._startX = xvals[0];
                this._endX = xvals[1];
            }
        }

        private _getX(dim: number): number {
            var retval = null;

            if (dim === 0) {
                retval = this.startX;
            } else if (dim === 1) {
                retval = this.endX;
            }

            if (isDate(retval)) {
                retval = asDate(retval).valueOf();
            }

            return retval;
        }

        _getChartType(): ChartType {
            return ChartType.Line;
        }
    }
}
module wijmo.chart.finance.analytics {
    "use strict";

    /**
     * Base class for overlay and indicator series (abstract).
     */
    export class OverlayIndicatorBase extends SeriesBase {
        private __plotter: _IPlotter;
        private __hitTester: _HitTester;

        _styles: any;

        // internal field for series that need multiple legend entries
        // in that case, set value to number of legend entries in ctor
        _seriesCount: number = 1;

        constructor() {
            super();
        }

        // access _IPlotter instance
        get _plotter(): _IPlotter {
            if (this.chart && !this.__plotter) {
                this.__plotter = this.chart._getPlotter(this);
            }
            return this.__plotter;
        }

        // access _HitTester instance
        get _hitTester(): _HitTester {
            if (this._plotter && !this.__hitTester) {
                this.__hitTester = this._plotter.hitTester;
            }
            return this.__hitTester;
        }

        // return ChartType
        _getChartType(): ChartType {
            return ChartType.Line;
        }

        // return original X-Values, if available
        _getXValues(): number[] {
            return (super.getValues(1) || this._plotter.dataInfo.getXVals());
        }

        // helper method to get a _DataPoint object for hit testing
        _getDataPoint(dataX: number, dataY: number, seriesIndex: number, pointIndex: number, ax: Axis, ay: Axis): _DataPoint {
            var dpt = new _DataPoint(seriesIndex, pointIndex, dataX, dataY);

            // set x & y related data
            dpt["y"] = dataY;
            dpt["yfmt"] = ay._formatValue(dataY);
            dpt["x"] = dataX;
            dpt["xfmt"] = ax._formatValue(dataX);

            return dpt;
        }

        // abstract method that determines whether or not calculations need to be ran
        _shouldCalculate(): boolean { return true; }

        // initialize internal collections
        _init(): void { }

        // responsible for calculating values
        _calculate(): void { }

        _clearValues(): void {
            super._clearValues();
            this.__plotter = null;
            this.__hitTester = null;
        }

        // helper for series with multiple names (csv)
        // Returns undefined or the name.
        _getName(dim: number): string {
            var retval: string = undefined;

            if (this.name) {
                if (this.name.indexOf(",")) {
                    var names = this.name.split(",");
                    if (names && names.length - 1 >= dim) {
                        retval = names[dim].trim();
                    }
                } else {
                    retval = this.name;
                }
            }

            return retval;
        }

        // helper for series with multiple styles
        // Returns the appropriate style for the given index, if
        // ones exists; null is returned otherwise.
        _getStyles(dim: number): any {
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
        }

        /* overrides for multiple legend items */
        legendItemLength(): number {
            return this._seriesCount;
        }

        measureLegendItem(engine: IRenderEngine, index: number): Size {
            var name = this._getName(index),
                retval = new Size(0, 0);

            if (name) {
                retval = this._measureLegendItem(engine, this._getName(index));
            }

            return retval;
        }

        drawLegendItem(engine: IRenderEngine, rect: Rect, index: number): void {
            var style = this._getStyles(index) || this.style,
                name = this._getName(index);

            if (name) {
                this._drawLegendItem(engine, rect, this._getChartType(), this._getName(index), style, this.symbolStyle);
            }
        }
    }

    /**
     * Base class for overlay and indicator series that render a single series (abstract).
     */
    export class SingleOverlayIndicatorBase extends OverlayIndicatorBase {
        private _period = 14;

        _xvals: number[];
        _yvals: number[];

        constructor() {
            super();
        }

        /**
         * Gets or sets the period for the calculation as an integer value.
         */
        get period(): any {
            return this._period;
        }
        set period(value: any) {
            if (value !== this._period) {
                this._period = asInt(value, false, true);
                this._clearValues();
                this._invalidate();
            }
        }

        // return the derived values
        getValues(dim: number): number[] {
            var retval: number[] = null;
            if (super._getLength() <= 0) {
                return retval;
            } else if (this._shouldCalculate()) {
                this._init();
                this._calculate();
            }

            if (dim === 0) {
                retval = this._yvals;
            } else if (dim === 1) {
                retval = this._xvals;
            }

            return retval;
        }

        // return limits for the derived values
        getDataRect(): Rect {
            var rect: Rect = null;
            if (super._getLength() <= 0) {
                return rect;
            } else if (this._shouldCalculate()) {
                this._init();
                this._calculate();
            }

            var xmin = _minimum(this._xvals),
                xmax = _maximum(this._xvals),
                ymin = _minimum(this._yvals),
                ymax = _maximum(this._yvals);

            if (_DataInfo.isValid(xmin) && _DataInfo.isValid(xmax) && _DataInfo.isValid(ymin) && _DataInfo.isValid(ymax)) {
                rect = new Rect(xmin, ymin, xmax - xmin, ymax - ymin);
            }

            return rect;
        }

        // clear the internal collections for the derived values
        _clearValues(): void {
            super._clearValues();
            this._xvals = null;
            this._yvals = null;
        }

        // determine if the derived values need to be calculated
        _shouldCalculate(): boolean {
            return !this._yvals || !this._xvals;
        }

        // initialize internal collections for the derived values
        _init(): void {
            super._init();
            this._yvals = [];
            this._xvals = [];
        }

        // override to get correct item for hit testing
        _getItem(pointIndex: number): any {
            if (super._getLength() <= 0) {
                return super._getItem(pointIndex);
            } else if (this._shouldCalculate()) {
                this._init();
                this._calculate();
            }

            var originalLen = super._getLength(),
                len = _minimum(this._yvals.length, this._xvals.length);

            // data index
            pointIndex = originalLen - len + pointIndex;
            return super._getItem(pointIndex);
        }
    }
}
module wijmo.chart.finance.analytics {
    "use strict";

    /**
     * Represents an Average True Range indicator series for the @see:FinancialChart.
     *
     * Average true range is used to measure the volatility of an asset. Average true range
     * does not provide any indication of the price's trend, but rather the degree of price
     * volatility.
     */
    export class ATR extends SingleOverlayIndicatorBase {
        constructor() {
            super();
            this.period = 14;
        }

        _calculate(): void {
            if (super._getLength() <= 0) {
                return;
            }

            var highs = super._getBindingValues(0),
                lows = super._getBindingValues(1),
                closes = super._getBindingValues(3),
                xs = this._getXValues();

            this._yvals = _avgTrueRng(highs, lows, closes, this.period);
            this._xvals = xs ? xs.slice(this.period - 1) : _range(this.period - 1, highs.length);
        }
    }
}
module wijmo.chart.finance.analytics {
    "use strict";

    /**
     * Represents a Commodity Channel Index indicator series for the @see:FinancialChart.
     *
     * The commodity channel index is an oscillator that measures an asset's current price
     * level relative to an average price level over a specified period of time.
     */
    export class CCI extends SingleOverlayIndicatorBase {
        private _constant = 0.015;

        constructor() {
            super();
            this.period = 20;
        }

        /**
         * Gets or sets the constant value for the CCI calculation.  The default
         * value is 0.015.
         */
        get constant(): number {
            return this._constant;
        }
        set constant(value: number) {
            if (value !== this._constant) {
                this._constant = asNumber(value, false);
                this._clearValues();
                this._invalidate();
            }
        }

        _calculate(): void {
            var originalLen = super._getLength();
            if (originalLen <= 0) {
                return;
            }

            var highs = super._getBindingValues(0),
                lows = super._getBindingValues(1),
                closes = super._getBindingValues(3),
                xs = this._getXValues();

            this._yvals = _cci(highs, lows, closes, this.period, this.constant);
            this._xvals = xs ? xs.slice(this.period - 1) : _range(this.period - 1, originalLen - 1);
        }
    }

    // calculate Commodity Channel Index for a set of financial data
    export function _cci(highs: number[], lows: number[], closes: number[], period: number, constant: number): number[] {
        asArray(highs, false); asArray(lows, false); asArray(closes, false);
        asInt(period, false, true);
        asNumber(constant, false, true);

        var len = _minimum(highs.length, lows.length, closes.length),
            typicalPrices: number[] = [],
            meanDeviations: number[] = [],
            smas: number[], i: number,
            ccis: number[] = [];

        assert(len > period && period > 1, "CCI period must be an integer less than the length of the data and greater than one.");

        // typical prices
        for (i = 0; i < len; i++) {
            typicalPrices.push(_average(highs[i], lows[i], closes[i]));
        }

        // simple moving average of typical prices
        smas = _sma(typicalPrices, period);

        // mean deviation
        var temp: number;
        for (i = 0; i < smas.length; i++) {
            temp = typicalPrices.slice(i, period + i)
                                .reduce((prev: number, curr: number) => prev + Math.abs(smas[i] - curr), 0);
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
}
module wijmo.chart.finance.analytics {
    "use strict";

    /**
     * Represents a Willaims %R indicator series for the @see:FinancialChart.
     *
     * Williams %R is a momentum indicator that is the inverse of a fast stochastic
     * oscillator (@see:Stochastic).  The Williams %R indicator is designed to
     * tell whether an asset is trading near the high or low of its trading range.
     */
    export class WilliamsR extends SingleOverlayIndicatorBase {
        constructor() {
            super();
            this.period = 14;
        }

        _calculate(): void {
            var originalLen = super._getLength();
            if (originalLen <= 0) {
                return;
            }

            var highs = super._getBindingValues(0),
                lows = super._getBindingValues(1),
                closes = super._getBindingValues(3),
                xs = this._getXValues();

            this._yvals = _williamsR(highs, lows, closes, this.period);
            this._xvals = xs ? xs.slice(this.period - 1) : _range(this.period - 1, originalLen - 1);
        }
    }

    // calculate Williams %R for a set of financial data
    export function _williamsR(highs: number[], lows: number[], closes: number[], period: number): number[] {
        asArray(highs, false); asArray(lows, false); asArray(closes, false);
        asInt(period, false, true);

        var len = _minimum(highs.length, lows.length, closes.length),
            extremeHighs: number[] = [],
            extremeLows: number[] = [],
            williamsRs: number[] = [],
            i: number;

        assert(len > period && period > 1, "Williams %R period must be an integer less than the length of the data and greater than one.");

        // get extreme high/low for each period
        for (i = period; i <= highs.length; i++) {
            extremeHighs.push(_maximum(highs.slice(i - period, i)));
            extremeLows.push(_minimum(lows.slice(i - period, i)));
        }

        // get subset of closing prices
        closes.splice(0, period - 1);

        // williams %r
        for (i = 0; i < extremeHighs.length; i++) {
            williamsRs.push((extremeHighs[i] - closes[i]) / (extremeHighs[i] - extremeLows[i]) * -100);
        }

        return williamsRs;
    }
}
module wijmo.chart.finance.analytics {
    "use strict";

    export enum MovingAverageType {
        Simple,
        Exponential
    }

    /**
     * Represents a Moving Average Envelopes overlay series for the @see:FinancialChart.
     *
     * Moving average envelopes are moving averages set above and below a standard moving
     * average.  The amount above/below the standard moving average is percentage based and
     * dictated by the @see:size property.
     */
    export class Envelopes extends OverlayIndicatorBase {
        private _upperYVals: number[];
        private _lowerYVals: number[];
        private _xVals: number[];

        private _period = 20;
        private _type = MovingAverageType.Simple;
        private _size = 0.025;

        constructor() {
            super();
            this.rendering.addHandler(this._rendering, this);
        }

        /**
         * Gets or sets the period for the calculation as an integer value.
         */
        get period(): any {
            return this._period;
        }
        set period(value: any) {
            if (value !== this._period) {
                this._period = asInt(value, false, true);
                this._clearValues();
                this._invalidate();
            }
        }

        /**
         * Gets or sets the moving average type for the
         * envelopes.  The default value is Simple.
         */
        get type(): MovingAverageType {
            return this._type;
        }
        set type(value: MovingAverageType) {
            if (value !== this._type) {
                this._type = asEnum(value, MovingAverageType, false);
                this._clearValues();
                this._invalidate();
            }
        }

        /**
         * Gets or set the size of the moving average
         * envelopes.  The default value is 2.5 percent (0.025).
         */
        get size(): number {
            return this._size;
        }
        set size(value: number) {
            if (value !== this._size) {
                this._size = asNumber(value, false, true);
                this._clearValues();
                this._invalidate();
            }
        }

        getDataRect(): Rect {
            var rect: Rect = null;
            if (super._getLength() <= 0) {
                return rect;
            } else if (this._shouldCalculate()) {
                this._init();
                this._calculate();
            }

            var ys = this._upperYVals.concat(this._lowerYVals),
                xmin = _minimum(this._xVals),
                xmax = _maximum(this._xVals),
                ymin = _minimum(ys),
                ymax = _maximum(ys);

            if (_DataInfo.isValid(xmin) && _DataInfo.isValid(xmax) && _DataInfo.isValid(ymin) && _DataInfo.isValid(ymax)) {
                rect = new Rect(xmin, ymin, xmax - xmin, ymax - ymin);
            }

            return rect;
        }

        _clearValues(): void {
            super._clearValues();
            this._upperYVals = null;
            this._lowerYVals = null;
            this._xVals = null;
        }

        _init(): void {
            super._init();
            this._upperYVals = [];
            this._lowerYVals = [];
            this._xVals = [];
        }

        _shouldCalculate(): boolean {
            return !this._upperYVals || !this._lowerYVals || !this._xVals;
        }

        // creates calculated values
        _calculate(): void {
            if (super._getLength() <= 0) {
                return;
            }

            var ys = super.getValues(0),
                xs = this._getXValues(),
                avgs: number[];

            // moving average calculations
            switch (this.type) {
                case MovingAverageType.Exponential:
                    avgs = _ema(ys, this.period);
                    break;
                case MovingAverageType.Simple:
                default:
                    avgs = _sma(ys, this.period);
                    break;
            }

            this._xVals = xs ? xs.slice(this.period - 1) : _range(this.period - 1, super._getLength() - 1);
            this._upperYVals = avgs.map((value: number) => value + (value * this.size));
            this._lowerYVals = avgs.map((value: number) => value - (value * this.size));
        }

        // custom rendering in order to draw multiple lines for a single SeriesBase object
        private _rendering(sender: Envelopes, args: RenderEventArgs): void {
            if (super._getLength() <= 0) {
                return;
            } else if (this._shouldCalculate()) {
                this._init();
                this._calculate();
            }

            var si = this.chart.series.indexOf(this),
                engine = args.engine,
                ax = this._getAxisX(), ay = this._getAxisY(),
                len = _minimum(this._upperYVals.length, this._lowerYVals.length, this._xVals.length),
                style = _BasePlotter.cloneStyle(this.style, ["fill"]),
                stroke = this._getSymbolStroke(si),
                clipPath = this.chart._plotrectId,
                swidth = 2;

            if (!len || len <= 0) { return; }

            engine.stroke = stroke;
            engine.strokeWidth = swidth;

            var xs: number[] = [],
                uys: number[] = [],
                lys: number[] = [],
                originalLen = this._getLength(),
                dpt: _DataPoint, area: _IHitArea, di: number;

            for (var i = 0; i < len; i++) {
                // data index
                di = originalLen - len + i;

                // x values
                xs.push(ax.convert(this._xVals[i]));

                // upper
                uys.push(ay.convert(this._upperYVals[i]));
                dpt = this._getDataPoint(this._xVals[i], this._upperYVals[i], si, di, ax, ay);
                area = new _CircleArea(new Point(xs[i], uys[i]), 0.5 * engine.strokeWidth);
                area.tag = dpt;
                this._hitTester.add(area, si);

                // lower
                lys.push(ay.convert(this._lowerYVals[i]));
                dpt = this._getDataPoint(this._xVals[i], this._lowerYVals[i], si, di, ax, ay);
                area = new _CircleArea(new Point(xs[i], lys[i]), 0.5 * engine.strokeWidth);
                area.tag = dpt;
                this._hitTester.add(area, si);
            }

            this._hitTester.add(new _LinesArea(xs, uys), si);
            this._hitTester.add(new _LinesArea(xs, lys), si);

            engine.drawLines(xs, uys, null, style, clipPath);
            engine.drawLines(xs, lys, null, style, clipPath);
        }

        getCalculatedValues(key: string): any[] {
            key = asString(key, false);

            var retval: any[] = [],
                i = 0;

            if (super._getLength() <= 0) {
                return retval;
            } else if (this._shouldCalculate()) {
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
        }
    }
}
module wijmo.chart.finance.analytics {
    "use strict";

    /**
     * Represents a Bollinger Bands&reg; overlay series for the @see:FinancialChart.
     *
     * <i>Bollinger Bands is a registered trademark of John Bollinger.</i>
     */
    export class BollingerBands extends OverlayIndicatorBase {
        private _upperYVals: number[];
        private _middleYVals: number[];
        private _lowerYVals: number[];
        private _xVals: number[];

        private _period = 20;
        private _multiplier = 2;

        constructor() {
            super();
            this.rendering.addHandler(this._rendering, this);
        }

        /**
         * Gets or sets the period for the calculation as an integer value.
         */
        get period(): any {
            return this._period;
        }
        set period(value: any) {
            if (value !== this._period) {
                this._period = asInt(value, false, true);
                this._clearValues();
                this._invalidate();
            }
        }

        /**
         * Gets or sets the standard deviation multiplier.
         */
        get multiplier(): number {
            return this._multiplier;
        }
        set multiplier(value: number) {
            if (value !== this._multiplier) {
                this._multiplier = asNumber(value, false, true);
                this._clearValues();
                this._invalidate();
            }
        }

        getDataRect(): Rect {
            if (super._getLength() <= 0) {
                return null;
            } else if (this._shouldCalculate()) {
                this._init();
                this._calculate();
            }

            var ys = this._upperYVals.concat(this._lowerYVals),
                xmin = _minimum(this._xVals),
                xmax = _maximum(this._xVals),
                ymin = _minimum(ys),
                ymax = _maximum(ys);

            if (_DataInfo.isValid(xmin) && _DataInfo.isValid(xmax) && _DataInfo.isValid(ymin) && _DataInfo.isValid(ymax)) {
                return new Rect(xmin, ymin, xmax - xmin, ymax - ymin);
            } else {
                return null;
            }
        }

        _clearValues(): void {
            super._clearValues();
            this._upperYVals = null;
            this._middleYVals = null;
            this._lowerYVals = null;
            this._xVals = null;
        }

        _shouldCalculate(): boolean {
            return !this._upperYVals || !this._middleYVals || !this._lowerYVals || !this._xVals;
        }

        _init(): void {
            super._init();
            this._upperYVals = [];
            this._middleYVals = [];
            this._lowerYVals = [];
            this._xVals = [];
        }

        _calculate(): void {
            var originalLen = super._getLength();
            if (originalLen <= 0) {
                return;
            }

            var ys = super.getValues(0),
                xs = this._getXValues();

            var values = _bollingerBands(ys, this.period, this.multiplier);
            this._upperYVals = values.uppers;
            this._middleYVals = values.middles;
            this._lowerYVals = values.lowers;
            this._xVals = xs ? xs.slice(this.period - 1) : _range(this.period - 1, originalLen - 1);
        }

        private _rendering(sender: SeriesBase, args: RenderEventArgs): void {
            if (super._getLength() <= 0) {
                return;
            } else if (this._shouldCalculate()) {
                this._init();
                this._calculate();
            }

            var si = this.chart.series.indexOf(this),
                engine = args.engine,
                ax = this._getAxisX(), ay = this._getAxisY(),
                len = _minimum(this._upperYVals.length, this._middleYVals.length, this._lowerYVals.length, this._xVals.length),
                style = _BasePlotter.cloneStyle(this.style, ["fill"]),
                stroke = this._getSymbolStroke(si),
                clipPath = this.chart._plotrectId,
                swidth = 2;

            if (!len || len <= 0) { return; }

            engine.stroke = stroke;
            engine.strokeWidth = swidth;

            var xs: number[] = [],
                uys: number[] = [],
                mys: number[] = [],
                lys: number[] = [],
                originalLen = this._getLength(),
                dpt: _DataPoint, area: _IHitArea, di: number;

            for (var i = 0; i < len; i++) {
                // data index
                di = originalLen - len + i;

                // x values
                xs.push(ax.convert(this._xVals[i]));

                // upper
                uys.push(ay.convert(this._upperYVals[i]));
                dpt = this._getDataPoint(this._xVals[i], this._upperYVals[i], si, di, ax, ay);
                area = new _CircleArea(new Point(xs[i], uys[i]), 0.5 * engine.strokeWidth);
                area.tag = dpt;
                this._hitTester.add(area, si);

                // middle
                mys.push(ay.convert(this._middleYVals[i]));
                dpt = this._getDataPoint(this._xVals[i], this._middleYVals[i], si, di, ax, ay);
                area = new _CircleArea(new Point(xs[i], mys[i]), 0.5 * engine.strokeWidth);
                area.tag = dpt;
                this._hitTester.add(area, si);

                // lower
                lys.push(ay.convert(this._lowerYVals[i]));
                dpt = this._getDataPoint(this._xVals[i], this._lowerYVals[i], si, di, ax, ay);
                area = new _CircleArea(new Point(xs[i], lys[i]), 0.5 * engine.strokeWidth);
                area.tag = dpt;
                this._hitTester.add(area, si);
            }

            this._hitTester.add(new _LinesArea(xs, uys), si);
            this._hitTester.add(new _LinesArea(xs, mys), si);
            this._hitTester.add(new _LinesArea(xs, lys), si);

            engine.drawLines(xs, uys, null, style, clipPath);
            engine.drawLines(xs, mys, null, style, clipPath);
            engine.drawLines(xs, lys, null, style, clipPath);
        }

        getCalculatedValues(key: string): any[] {
            key = asString(key, false);

            var retval: any[] = [],
                i = 0;

            if (super._getLength() <= 0) {
                return retval;
            } else if (this._shouldCalculate()) {
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
        }
    }

    // calculate Bollinger Bands for a set of financial data
    export function _bollingerBands(ys: number[], period: number, multiplier: number): any {
        asArray(ys, false);
        asInt(period, false, true);
        asNumber(multiplier, false, true);
        assert(ys.length > period && period > 1, "Bollinger Bands period must be an integer less than the length of the data and greater than one.");

        var avgs = _sma(ys, period),
            devs: number[] = [],
            i: number;

        // get standard deviations
        for (i = period; i <= ys.length; i++) {
            devs.push(_stdDeviation(ys.slice(i - period, i)));
        }

        var middles = avgs,
            uppers = avgs.map((value: number, index: number) => value + (devs[index] * multiplier)),
            lowers = avgs.map((value: number, index: number) => value - (devs[index] * multiplier));

        return {
            lowers: lowers,
            middles: middles,
            uppers: uppers
        };
    }
}
module wijmo.chart.finance.analytics {
    "use strict";

    /**
     * Represents a Relative Strength Index indicator series for the @see:FinancialChart.
     *
     * Relative strength index is a momentum osciallator designed to measure the current
     * and historical strength or weakness of an asset based on the closing prices of a
     * recent trading period.
     */
    export class RSI extends SingleOverlayIndicatorBase {
        constructor() {
            super();
            this.period = 14;
        }

        _calculate(): void {
            var originalLen = super._getLength();
            if (originalLen <= 0) {
                return;
            }

            var ys = super._getBindingValues(0), // getValues(0) is overridden
                xs = this._getXValues();

            this._yvals = _rsi(ys, this.period);
            this._xvals = xs ? xs.slice(this.period) : _range(this.period, originalLen);
        }
    }

    // calculate Relative Strength Index for a set of financial data
    export function _rsi(ys: number[], period: number): number[] {
        asArray(ys, false);
        asInt(period, true, false);
        assert(ys.length > period && period > 1, "RSI period must be an integer less than the length of the data and greater than one.");

        var changes: number[] = [],
            avgGains: number[] = [],
            avgLosses: number[] = [],
            gains: number[], losses: number[],
            rsis: number[] = [],
            rs: number, i: number;

        // calculate changes
        for (i = 1; i < ys.length; i++) {
            changes.push(ys[i] - ys[i - 1]);
        }

        // get gains and losses
        gains = changes.map((value: number) => value > 0 ? value : 0);
        losses = changes.map((value: number) => value < 0 ? Math.abs(value) : 0);

        // calculate rs and rsi
        for (i = period; i <= changes.length; i++) {
            if (i === period) {
                avgGains.push(_sum(gains.slice(i - period, i)) / period);
                avgLosses.push(_sum(losses.slice(i - period, i)) / period);
            } else {
                avgGains.push((gains[i - 1] + (avgGains[i - period - 1] * (period - 1))) / period);
                avgLosses.push((losses[i - 1] + (avgLosses[i - period - 1] * (period - 1))) / period);
            }

            rs = avgGains[i - period] / avgLosses[i - period];
            rs = isFinite(rs) ? rs : 0;
            rsis.push(100 - (100 / (1 + rs)));
        }

        return rsis;
    }
}
module wijmo.chart.finance.analytics {
    "use strict";

    /**
     * Base class for @see:Macd and @see:MacdHistogram series (abstract).
     */
    export class MacdBase extends OverlayIndicatorBase {
        _macdXVals: number[];
        _macdVals: number[];
        _signalXVals: number[];
        _signalVals: number[];
        _histogramXVals: number[];
        _histogramVals: number[];

        private _fastPeriod = 12;
        private _slowPeriod = 26;
        private _smoothingPeriod = 9;

        constructor() {
            super();
        }

        /**
         * Gets or sets the fast exponential moving average period
         * for the MACD line.
         */
        get fastPeriod(): number {
            return this._fastPeriod;
        }
        set fastPeriod(value: number) {
            if (value !== this._fastPeriod) {
                this._fastPeriod = asInt(value, false, true);
                this._clearValues();
                this._invalidate();
            }
        }

        /**
         * Gets or sets the slow exponential moving average period
         * for the MACD line.
         */
        get slowPeriod(): number {
            return this._slowPeriod;
        }
        set slowPeriod(value: number) {
            if (value !== this._slowPeriod) {
                this._slowPeriod = asInt(value, false, true);
                this._clearValues();
                this._invalidate();
            }
        }

        /**
         * Gets or sets the exponential moving average period
         * for the signal line.
         */
        get smoothingPeriod(): number {
            return this._smoothingPeriod;
        }
        set smoothingPeriod(value: number) {
            if (value !== this._smoothingPeriod) {
                this._smoothingPeriod = asInt(value, false, true);
                this._clearValues();
                this._invalidate();
            }
        }

        _clearValues(): void {
            super._clearValues();
            this._macdVals = null;
            this._macdXVals = null;
            this._signalVals = null;
            this._signalXVals = null;
            this._histogramVals = null;
            this._histogramXVals = null;
        }

        _shouldCalculate(): boolean {
            return !this._macdVals || !this._macdXVals ||
                !this._signalVals || !this._signalXVals ||
                !this._histogramVals || !this._histogramXVals;
        }

        _init(): void {
            super._init();
            this._macdVals = [];
            this._macdXVals = [];
            this._signalVals = [];
            this._signalXVals = [];
            this._histogramVals = [];
            this._histogramXVals = [];
        }

        _calculate(): void {
            var originalLen = super._getLength();
            if (originalLen <= 0) {
                return;
            }

            var ys = super.getValues(0),
                xs = this._getXValues();

            var values = _macd(ys, this.fastPeriod, this.slowPeriod, this.smoothingPeriod);
            this._macdVals = values.macds;
            this._signalVals = values.signals;
            this._histogramVals = values.histograms;

            this._macdXVals = xs ? xs.slice(originalLen - this._macdVals.length, originalLen) : _range(originalLen - this._macdVals.length, originalLen - 1);
            this._signalXVals = xs ? xs.slice(originalLen - this._signalVals.length, originalLen) : _range(originalLen - this._signalVals.length, originalLen - 1);
            this._histogramXVals = xs ? xs.slice(originalLen - this._histogramVals.length, originalLen) : _range(originalLen - this._histogramVals.length, originalLen - 1);
        }
    }

    /**
     * Represents a Moving Average Convergence/Divergence (MACD) indicator series
     * for the @see:FinancialChart.
     *
     * The MACD indicator is designed to reveal changes in strength, direction, momentum,
     * and duration of an asset's price trend.
     */
    export class Macd extends MacdBase {
        constructor() {
            super();

            this._seriesCount = 2;

            this.rendering.addHandler(this._rendering, this);
        }

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
        get styles(): any {
            return this._styles;
        }
        set styles(value: any) {
            if (value !== this._styles) {
                this._styles = value;
                this._invalidate();
            }
        }

        getDataRect(): Rect {
            var rect: Rect = null;
            if (super._getLength() <= 0) {
                return rect;
            } else if (this._shouldCalculate()) {
                this._init();
                this._calculate();
            }

            var ys: number[] = [],
                xs: number[] = [];

            xs.push.apply(xs, this._macdXVals);
            xs.push.apply(xs, this._signalXVals);
            ys.push.apply(ys, this._macdVals);
            ys.push.apply(ys, this._signalVals);

            var xmin = _minimum(xs),
                xmax = _maximum(xs),
                ymin = _minimum(ys),
                ymax = _maximum(ys);

            if (_DataInfo.isValid(xmin) && _DataInfo.isValid(xmax) && _DataInfo.isValid(ymin) && _DataInfo.isValid(ymax)) {
                rect = new Rect(xmin, ymin, xmax - xmin, ymax - ymin);
            }

            return rect;
        }

        private _rendering(sender: SeriesBase, args: RenderEventArgs): void {
            if (super._getLength() <= 0) {
                return;
            } else if (this._shouldCalculate()) {
                this._init();
                this._calculate();
            }

            var si = this.chart.series.indexOf(this),
                engine = args.engine,
                ax = this._getAxisX(), ay = this._getAxisY(),
                style = _BasePlotter.cloneStyle(this.style, ["fill"]),
                stroke = this._getSymbolStroke(si),
                clipPath = this.chart._plotrectId,
                swidth = 2,
                macdStyle = null, macdStroke = stroke, macdStrokeWidth = swidth,
                signalStyle = null, signalStroke = stroke, signalStrokeWidth = swidth;

            // handle "styles"
            if (this.styles && isObject(this.styles)) {
                if (this.styles.macdLine && isObject(this.styles.macdLine)) {
                    macdStyle = _BasePlotter.cloneStyle(this.styles.macdLine, ["fill"]);
                    macdStroke = macdStyle.stroke ? macdStyle.stroke : stroke;
                    macdStrokeWidth = macdStyle.strokeWidth ? macdStyle.strokeWidth : swidth;
                }

                if (this.styles.signalLine && isObject(this.styles.signalLine)) {
                    signalStyle = _BasePlotter.cloneStyle(this.styles.signalLine, ["fill"]);
                    signalStroke = signalStyle.stroke ? signalStyle.stroke : stroke;
                    signalStrokeWidth = signalStyle.strokeWidth ? signalStyle.strokeWidth : swidth;
                }
            }

            var macdVals: number[] = [],
                macdXVals: number[] = [],
                signalVals: number[] = [],
                signalXVals: number[] = [],
                dpt: _DataPoint, area: _IHitArea,
                originalLen = this._getLength(),
                i: number, di: number;

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
                area = new _CircleArea(new Point(macdXVals[i], macdVals[i]), 0.5 * engine.strokeWidth);
                area.tag = dpt;
                this._hitTester.add(area, si);
            }
            this._hitTester.add(new _LinesArea(macdXVals, macdVals), si);
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
                area = new _CircleArea(new Point(signalXVals[i], signalVals[i]), 0.5 * engine.strokeWidth);
                area.tag = dpt;
                this._hitTester.add(area, si);
            }
            this._hitTester.add(new _LinesArea(signalXVals, signalVals), si);
            engine.stroke = signalStroke;
            engine.strokeWidth = signalStrokeWidth;
            engine.drawLines(signalXVals, signalVals, null, style, clipPath);
        }

        getCalculatedValues(key: string): any[] {
            key = asString(key, false);

            var retval: any[] = [],
                i = 0;

            if (super._getLength() <= 0) {
                return retval;
            } else if (this._shouldCalculate()) {
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
        }
    }

    /**
     * Represents a Moving Average Convergence/Divergence (MACD) Histogram indicator series
     * for the @see:FinancialChart.
     *
     * The MACD indicator is designed to reveal changes in strength, direction, momentum,
     * and duration of an asset's price trend.
     */
    export class MacdHistogram extends MacdBase {
        constructor() {
            super();
        }

        getValues(dim: number): number[] {
            var retval: number[] = null;
            if (super._getLength() <= 0) {
                return retval;
            } else if (this._shouldCalculate()) {
                this._init();
                this._calculate();
            }

            if (dim === 0) {
                retval = this._histogramVals;
            } else if (dim === 1) {
                retval = this._histogramXVals;
            }

            return retval;
        }

        getDataRect(): Rect {
            var rect: Rect = null;
            if (super._getLength() <= 0) {
                return rect;
            } else if (this._shouldCalculate()) {
                this._init();
                this._calculate();
            }

            var xmin = _minimum(this._histogramXVals),
                xmax = _maximum(this._histogramXVals),
                ymin = _minimum(this._histogramVals),
                ymax = _maximum(this._histogramVals);

            if (_DataInfo.isValid(xmin) && _DataInfo.isValid(xmax) && _DataInfo.isValid(ymin) && _DataInfo.isValid(ymax)) {
                rect = new Rect(xmin, ymin, xmax - xmin, ymax - ymin);
            }

            return rect;
        }

        _getChartType(): ChartType {
            return ChartType.Column;
        }

        // override to get correct item for hit testing
        _getItem(pointIndex: number): any {
            var originalLen = super._getLength(),
                len = _minimum(this._histogramVals.length, this._histogramXVals.length);

            // data index
            pointIndex = originalLen - len + pointIndex;
            return super._getItem(pointIndex);
        }
    }


    // calculate MACD for a set of financial data
    export function _macd(ys: number[], fastPeriod: number, slowPeriod: number, smoothingPeriod: number): any {
        asArray(ys, false);
        asInt(fastPeriod, false, true); asInt(slowPeriod, false, true); asInt(smoothingPeriod, false, true);

        var opposite = fastPeriod > slowPeriod,
            temp: number;
        if (opposite) {
            temp = slowPeriod;
            slowPeriod = fastPeriod;
            fastPeriod = temp;
        }

        var fastEmas = _ema(ys, fastPeriod),
            slowEmas = _ema(ys, slowPeriod),
            macds: number[] = [],
            histograms: number[] = [],
            signals: number[], i: number;

        // get subset of fast emas for macd line calculation
        fastEmas.splice(0, slowPeriod - fastPeriod);

        // macd line
        for (i = 0; i < fastEmas.length; i++) {
            temp = fastEmas[i] - slowEmas[i];
            if (opposite) temp *= -1;
            macds.push(temp);
        }

        // signal line
        signals = _ema(macds, smoothingPeriod);

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
}
module wijmo.chart.finance.analytics {
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
    export class Stochastic extends OverlayIndicatorBase {
        private _kVals: number[];
        private _kXVals: number[];
        private _dVals: number[];
        private _dXVals: number[];

        private _kPeriod = 14;
        private _dPeriod = 3;
        private _smoothingPeriod = 1;

        constructor() {
            super();

            this._seriesCount = 2;

            this.rendering.addHandler(this._rendering, this);
        }

        /**
         * Gets or sets the period for the %K calculation.
         */
        get kPeriod(): number {
            return this._kPeriod;
        }
        set kPeriod(value: number) {
            if (value !== this._kPeriod) {
                this._kPeriod = asInt(value, false, true);
                this._clearValues();
                this._invalidate();
            }
        }

        /**
         * Gets or sets the period for the %D simple moving average.
         */
        get dPeriod(): number {
            return this._dPeriod;
        }
        set dPeriod(value: number) {
            if (value !== this._dPeriod) {
                this._dPeriod = asInt(value, false, true);
                this._clearValues();
                this._invalidate();
            }
        }

        /**
         * Gets or sets the smoothing period for full %K.
         */
        get smoothingPeriod(): number {
            return this._smoothingPeriod;
        }
        set smoothingPeriod(value: number) {
            if (value !== this._smoothingPeriod) {
                this._smoothingPeriod = asInt(value, false, true);
                this._clearValues();
                this._invalidate();
            }
        }

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
        get styles(): any {
            return this._styles;
        }
        set styles(value: any) {
            if (value !== this._styles) {
                this._styles = value;
                this._invalidate();
            }
        }

        getDataRect(): Rect {
            var rect: Rect = null;
            if (super._getLength() <= 0) {
                return rect;
            } else if (this._shouldCalculate()) {
                this._init();
                this._calculate();
            }

            var ys = this._kVals.concat(this._dVals),
                xs = this._kXVals.concat(this._dXVals),
                xmin = _minimum(xs),
                xmax = _maximum(xs),
                ymin = _minimum(ys),
                ymax = _maximum(ys);

            if (_DataInfo.isValid(xmin) && _DataInfo.isValid(xmax) && _DataInfo.isValid(ymin) && _DataInfo.isValid(ymax)) {
                rect = new Rect(xmin, ymin, xmax - xmin, ymax - ymin);
            }

            return rect;
        }

        _clearValues(): void {
            super._clearValues();
            this._kVals = null;
            this._kXVals = null;
            this._dVals = null;
            this._dXVals = null;
        }

        _shouldCalculate(): boolean {
            return !this._kVals || !this._kXVals ||
                !this._dVals || !this._dXVals;
        }

        _init(): void {
            super._init();
            this._kVals = [];
            this._kXVals = [];
            this._dVals = [];
            this._dXVals = [];
        }

        _calculate(): void {
            var originalLen = super._getLength();
            if (super._getLength() <= 0) {
                return;
            }

            var highs = super._getBindingValues(0),
                lows = super._getBindingValues(1),
                closes = super._getBindingValues(3),
                xs = this._getXValues();

            var values = _stochastic(highs, lows, closes, this.kPeriod, this.dPeriod, this.smoothingPeriod);
            this._kVals = values.ks;
            this._dVals = values.ds;

            // get %K x-values
            this._kXVals = xs ? xs.slice(this.kPeriod - 1) : _range(this.kPeriod - 1, originalLen - 1);
            if (this.smoothingPeriod && this.smoothingPeriod > 1) {
                this._kXVals = this._kXVals.slice(this._kXVals.length - this._kVals.length, this._kXVals.length);
            }

            // get %D x-values
            this._dXVals = this._kXVals.slice(this._kXVals.length - this._dVals.length, this._kXVals.length);
        }

        private _rendering(sender: SeriesBase, args: RenderEventArgs): void {
            if (super._getLength() <= 0) {
                return;
            } else if (this._shouldCalculate()) {
                this._init();
                this._calculate();
            }

            var si = this.chart.series.indexOf(this),
                engine = args.engine,
                ax = this._getAxisX(), ay = this._getAxisY(),
                style = _BasePlotter.cloneStyle(this.style, ["fill"]),
                stroke = this._getSymbolStroke(si),
                clipPath = this.chart._plotrectId,
                swidth = 2,
                kStyle = null, kStroke = stroke, kStrokeWidth = swidth,
                dStyle = null, dStroke = stroke, dStrokeWidth = swidth;

            // handle "styles"
            if (this.styles && isObject(this.styles)) {
                if (this.styles.kLine && isObject(this.styles.kLine)) {
                    kStyle = _BasePlotter.cloneStyle(this.styles.kLine, ["fill"]);
                    kStroke = kStyle.stroke ? kStyle.stroke : stroke;
                    kStrokeWidth = kStyle.strokeWidth ? kStyle.strokeWidth : swidth;
                }

                if (this.styles.dLine && isObject(this.styles.dLine)) {
                    dStyle = _BasePlotter.cloneStyle(this.styles.dLine, ["fill"]);
                    dStroke = dStyle.stroke ? dStyle.stroke : stroke;
                    dStrokeWidth = dStyle.strokeWidth ? dStyle.strokeWidth : swidth;
                }
            }

            var kVals: number[] = [],
                kXVals: number[] = [],
                dVals: number[] = [],
                dXVals: number[] = [],
                originalLen = this._getLength(),
                dpt: _DataPoint, area: _IHitArea,
                i: number, di: number;

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
                area = new _CircleArea(new Point(kXVals[i], kVals[i]), 0.5 * engine.strokeWidth);
                area.tag = dpt;
                this._hitTester.add(area, si);
            }
            this._hitTester.add(new _LinesArea(kXVals, kVals), si);
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
                area = new _CircleArea(new Point(dXVals[i], dVals[i]), 0.5 * engine.strokeWidth);
                area.tag = dpt;
                this._hitTester.add(area, si);
            }
            this._hitTester.add(new _LinesArea(dXVals, dVals), si);
            engine.stroke = dStroke;
            engine.strokeWidth = dStrokeWidth;
            engine.drawLines(dXVals, dVals, null, style, clipPath);
        }

        getCalculatedValues(key: string): any[] {
            key = asString(key, false);

            var retval: any[] = [],
                i = 0;

            if (super._getLength() <= 0) {
                return retval;
            } else if (this._shouldCalculate()) {
                this._init();
                this._calculate();
            }

            switch(key) {
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
        }
    }

    // calculate Stochastics for a set of financial data
    export function _stochastic(highs: number[], lows: number[], closes: number[], kPeriod: number, dPeriod: number, smoothingPeriod: number): any {
        asArray(highs, false); asArray(lows, false); asArray(closes, false);
        asInt(kPeriod, false, true); asInt(dPeriod, false, true); asInt(smoothingPeriod, true, true);

        var extremeHighs: number[] = [],
            extremeLows: number[] = [],
            kvals: number[] = [],
            dvals: number[], i: number;

        // get extreme highs/lows for each period
        for (i = kPeriod; i <= highs.length; i++) {
            extremeHighs.push(_maximum(highs.slice(i - kPeriod, i)));
            extremeLows.push(_minimum(lows.slice(i - kPeriod, i)));
        }

        // get subset of closing prices
        closes = closes.slice(kPeriod - 1);

        // %K
        for (i = 0; i < closes.length; i++) {
            kvals.push((closes[i] - extremeLows[i]) / (extremeHighs[i] - extremeLows[i]) * 100);
        }

        // %K in slow/full
        if (smoothingPeriod && smoothingPeriod > 1) {
            kvals = _sma(kvals, smoothingPeriod);
        }

        // %D
        dvals = _sma(kvals, dPeriod);

        return {
            ks: kvals,
            ds: dvals
        };
    }
}
