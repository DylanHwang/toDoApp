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
module wijmo.chart.interaction {
    'use strict';
    
    /**
     * Range Slider.
     */
    export class _RangeSlider {
        // Static class
        private static _HRANGESLIDER = 'wj-chart-hrangeslider';
        private static _VRANGESLIDER = 'wj-chart-vrangeslider';
        private static _RANGESLIDER_DECBTN = 'wj-rangeslider-decbtn';
        private static _RANGESLIDER_INCBTN = 'wj-rangeslider-incbtn';
        private static _RANGESLIDER_RANGEHANDLE = 'wj-rangeslider-rangehandle';
        private static _RANGESLIDER_MINHANDLE = 'wj-rangeslider-minhandle';
        private static _RANGESLIDER_MAXHANDLE = 'wj-rangeslider-maxhandle';
        private static _RANGESLIDER_HANDLE_ACTIVE = 'wj-rangeslider-handle-active';

        // fields
        private _isVisible: boolean = true;
        private _buttonsVisible: boolean = true;
        private _minScale: number = 0;
        private _maxScale: number = 1;
        private _seamless: boolean = false;

        // elements
        private _rsContainer: HTMLElement = null;
        private _rsEle: HTMLElement = null;
        private _decBtn: HTMLElement = null;
        private _incBtn: HTMLElement = null;
        private _rsContent: HTMLElement = null;
        private _minHandler: HTMLElement = null;
        private _rangeHandler: HTMLElement = null;
        private _maxHandler: HTMLElement = null;

        // event
        private _wrapperSliderMousedown = null;
        private _wrapperDocMouseMove = null;
        private _wrapperDocMouseup = null;
        private _wrapperBtnMousedown = null;
        private _wrapperRangeSpaceMousedown = null;
        private _wrapperRangeMouseleave = null;

        // helper field
        private _isTouch: boolean = false;
        private _slidingInterval = null;
        private _rangeSliderRect = null;
        private _isHorizontal: boolean = true;
        private _isBtnMousedown: boolean = false;
        private _needSpaceClick: boolean = false;
        private _hasButtons: boolean = true;
        private _movingEle: HTMLElement = null;
        private _movingOffset: Rect = null;
        private _range: number = null;
        private _plotBox;
        private _startPt: Point = null;

        _minPos: number = 0;
        _maxPos: number = 1;

        constructor(container: HTMLElement, needSpaceClick: boolean, hasButtons?: boolean, options?) {
            if (!container) {
                assert(false, 'The container cannot be null.');
            }

            this._isTouch = 'ontouchstart' in window; //isTouchDevice();

            this._needSpaceClick = needSpaceClick; // whether has space click function
            this._hasButtons = hasButtons;  //whether has dec and inc buttons
            wijmo.copy(this, options);
            this._createSlider(container);
        }

        /**
         * Gets or sets whether the increase/decrease buttons are displayed or not.
         */
        get buttonsVisible(): boolean {
            return this._buttonsVisible;
        }
        set buttonsVisible(value: boolean) {
            if (value != this._buttonsVisible) {
                this._buttonsVisible = asBoolean(value);

                if (!this._rsContainer || !this._hasButtons) {
                    return;
                }
                this._refresh();
            }
        }

        /**
         * Gets or sets the orientation of the range slider.
         */
        get isHorizontal(): boolean {
            return this._isHorizontal;
        }
        set isHorizontal(value: boolean) {
            if (value != this._isHorizontal) {
                this._isHorizontal = asBoolean(value);
                if (!this._rsContainer) {
                    return;
                }
                this._invalidate();
            }
        }

        /**
         * Gets or sets the visibility of the range slider.
         */
        get isVisible(): boolean {
            return this._isVisible;
        }
        set isVisible(value: boolean) {
            if (value != this._isVisible) {
                this._isVisible = asBoolean(value);
                if (!this._rsContainer) {
                    return;
                }
                this._rsContainer.style.visibility = this._isVisible ? 'visible' : 'hidden';
            }
        }

        /**
         * Gets or sets the minimum range scale of the range slider.
         */
        get minScale(): number {
            return this._minScale;
        }
        set minScale(value: number) {
            if (value >= 0 && value != this._minScale) {
                this._minScale = asNumber(value);           
            }
        }

        /**
         * Gets or sets the maximum range scale of the range slider.
         */
        get maxScale(): number {
            return this._maxScale;
        }
        set maxScale(value: number) {
            if (value >= 0 && value != this._maxScale) {
                this._maxScale = asNumber(value);
            }
        }

        /**
         * Gets or sets a value that determines whether the minimal and 
         * maximal handler will move seamlessly.
         */
        get seamless(): boolean {
            return this._seamless;
        }
        set seamless(value: boolean) {
            if ( value != this._seamless) {
                this._seamless = asBoolean(value);
            }
        }

        /**
        * Occurs after the range changes.
        */
        rangeChanged = new Event();

        /**
         * Raises the @see:rangeChanged event.
         */
        onRangeChanged(e?: EventArgs) {
            this.rangeChanged.raise(this, e);
        }

        /**
        * Occurs while the range is changing.
        */
        rangeChanging = new Event();

        /**
         * Raises the @see:rangeChanging event.
         */
        onRangeChanging(e?: EventArgs) {
            this.rangeChanging.raise(this, e);
        }

        get _isSliding() {
            return this._startPt !== null;
        }

        get _handleWidth(): number {
            return this._minHandler.offsetWidth;
        }

        private _createSlider(container: HTMLElement) {
            var sCss = this._isHorizontal ? _RangeSlider._HRANGESLIDER : _RangeSlider._VRANGESLIDER,
                decBtnCss = this._isHorizontal ? 'wj-glyph-left' : 'wj-glyph-down',
                incBtnCss = this._isHorizontal ? 'wj-glyph-right' : 'wj-glyph-up',
                off, box;

            this._rsContainer = container;
            this._rsContainer.style.visibility = this._isVisible ? 'visible' : 'hidden';
            this._rsEle = createElement('<div class="wj-chart-rangeslider ' + sCss + '"></div>');
            this._rsContainer.appendChild(this._rsEle);

            if (this._hasButtons) {
                 //decrease button
                this._decBtn = createElement(
                    '<button class="wj-rangeslider-decbtn wj-btn wj-btn-default" type="button" tabindex="-1">' +
                    '<span class="' + decBtnCss + ' ' + _RangeSlider._RANGESLIDER_DECBTN + '"></span>' +
                    '</button>');
                this._rsEle.appendChild(this._decBtn);

                //increase button
                this._incBtn = createElement(
                    '<button class="wj-rangeslider-incbtn wj-btn wj-btn-default" type="button" tabindex="-1">' +
                    '<span class="' + incBtnCss + ' ' + _RangeSlider._RANGESLIDER_INCBTN + '"></span>' +
                    '</button>');
                this._rsEle.appendChild(this._incBtn);
            }

            //creating range slider
            this._rsContent = createElement('<div class="wj-rangeslider-content">' +
                '<div class="wj-rangeslider-rangehandle"></div>' +
                '<div class="wj-rangeslider-minhandle"></div>' +
                '<div class="wj-rangeslider-maxhandle"></div>');
            this._rsEle.appendChild(this._rsContent);

            this._minHandler = <HTMLElement>this._rsContent.querySelector('.' + _RangeSlider._RANGESLIDER_MINHANDLE);
            this._rangeHandler = <HTMLElement>this._rsContent.querySelector('.' + _RangeSlider._RANGESLIDER_RANGEHANDLE);
            this._maxHandler = <HTMLElement>this._rsContent.querySelector('.' + _RangeSlider._RANGESLIDER_MAXHANDLE);

            //bind event
            this._wrapperSliderMousedown = this._onSliderMousedown.bind(this);
            this._wrapperDocMouseMove = this._onDocMouseMove.bind(this);
            this._wrapperDocMouseup = this._onDocMouseup.bind(this);
            this._wrapperRangeSpaceMousedown = this._onRangeSpaceMousedown.bind(this);
            this._wrapperRangeMouseleave = this._onRangeMouseleave.bind(this);
            this._wrapperBtnMousedown = this._onBtnMousedown.bind(this);
            this._switchEvent(true);
        }

        private _switchEvent(isOn: boolean) {
            var eventListener = isOn ? 'addEventListener' : 'removeEventListener',
                eventHandler = isOn ? 'addHandler' : 'removeHandler';

            if (this._rsContainer) {
                if (this._needSpaceClick) {
                    this._rsEle[eventListener]('mousedown', this._wrapperRangeSpaceMousedown);
                }
                this._rsEle[eventListener]('mouseleave', this._wrapperRangeMouseleave);
                this._rsContent[eventListener]('mousedown', this._wrapperSliderMousedown);

                if (this._hasButtons) {
                    this._decBtn[eventListener]('mousedown', this._wrapperBtnMousedown);
                    this._incBtn[eventListener]('mousedown', this._wrapperBtnMousedown);
                }

                document[eventListener]('mousemove', this._wrapperDocMouseMove);
                document[eventListener]('mouseup', this._wrapperDocMouseup);

                if ('ontouchstart' in window) {
                    if (this._needSpaceClick) {
                        this._rsEle[eventListener]('touchstart', this._wrapperRangeSpaceMousedown);
                    }
                    this._rsContent[eventListener]('touchstart', this._wrapperSliderMousedown);

                    if (this._hasButtons) {
                        this._decBtn[eventListener]('touchstart', this._wrapperBtnMousedown);
                        this._incBtn[eventListener]('touchstart', this._wrapperBtnMousedown);
                    }

                    document[eventListener]('touchmove', this._wrapperDocMouseMove);
                    document[eventListener]('touchend', this._wrapperDocMouseup);
                }
            }
        }

        private _onSliderMousedown(e) {
            if (!this._isVisible) {
                return;
            }

            this._movingEle = e.srcElement || e.target;
            this._startPt = e instanceof MouseEvent ?
            new wijmo.Point(e.pageX, e.pageY) :
            new wijmo.Point(e.changedTouches[0].pageX, e.changedTouches[0].pageY);
            wijmo.removeClass(this._minHandler, _RangeSlider._RANGESLIDER_HANDLE_ACTIVE);
            wijmo.removeClass(this._maxHandler, _RangeSlider._RANGESLIDER_HANDLE_ACTIVE);
            this._movingOffset = wijmo.getElementRect(this._movingEle);
            if (this._movingEle != this._rangeHandler) {
                if (this._isHorizontal) {
                    this._movingOffset.left += 0.5 * this._movingEle.offsetWidth;
                } else {
                    this._movingOffset.top += 0.5 * this._movingEle.offsetHeight;
                }
                wijmo.addClass(this._movingEle, _RangeSlider._RANGESLIDER_HANDLE_ACTIVE);
            } else {
                this._range = this._maxPos - this._minPos;
            }

            e.preventDefault();
        }

        private _onDocMouseMove(e) {
            if (!this._isVisible || !this._startPt) {
                return;
            }

            var movingPt = e instanceof MouseEvent ?
                new wijmo.Point(e.pageX, e.pageY) :
                new wijmo.Point(e.changedTouches[0].pageX, e.changedTouches[0].pageY);

            this._onMove(movingPt);
            //e.preventDefault();
        }

        private _onMove(mvPt: wijmo.Point) {
            var self = this,
                strPt = this._startPt, movingOffset = this._movingOffset,
                plotBox = this._plotBox, range = this._range, moving = this._movingEle,
                left = this._minHandler, middle = this._rangeHandler, right = this._maxHandler,
                x, y, pos;

            if (strPt && movingOffset) {
                if (this._isHorizontal) {
                    x = movingOffset.left + mvPt.x - strPt.x;
                    pos = (x - plotBox.x) / plotBox.width;
                } else {
                    y = movingOffset.top + mvPt.y - strPt.y;
                    pos = 1 - (y - plotBox.y) / plotBox.height;
                }

                if (pos < 0) {
                    pos = 0;
                } else if (pos > 1) {
                    pos = 1;
                } 

                if (moving === left) {
                    if (this._seamless && this._minScale === 0 && pos >= this._maxPos) {
                          self._minPos = self._maxPos;
                          self._movingEle = right;
                          wijmo.removeClass(this._minHandler, _RangeSlider._RANGESLIDER_HANDLE_ACTIVE);
                          wijmo.addClass(this._maxHandler, _RangeSlider._RANGESLIDER_HANDLE_ACTIVE);
                        } else {
                        if (pos > this._maxPos - this._minScale) {
                            pos = this._maxPos - this._minScale;
                        }

                        if (pos < this._maxPos - this._maxScale) {
                            pos = this._maxPos - this._maxScale;
                        }

                        this._minPos = pos;
                    }                  
                    
                } else if (moving === right) {

                    if (this._seamless && this._minScale === 0 && pos <= this._minPos) {
                            self._maxPos = self._minPos;
                            self._movingEle = left;
                            wijmo.removeClass(this._maxHandler, _RangeSlider._RANGESLIDER_HANDLE_ACTIVE);
                            wijmo.addClass(this._minHandler, _RangeSlider._RANGESLIDER_HANDLE_ACTIVE);
                        } else {
                      if (pos < this._minPos + this._minScale) {
                        pos = this._minPos + this._minScale;
                    }
                        if (pos > this._minPos + this._maxScale) {
                            pos = this._minPos + this._maxScale;
                        }
                    this._maxPos = pos;
                  }

                } else if (moving === middle) {
                    if (this._isHorizontal) {
                        this._minPos = pos;
                        this._maxPos = this._minPos + range;
                        if (this._maxPos >= 1) {
                            this._maxPos = 1;
                            this._minPos = this._maxPos - range;
                        }
                    } else {
                        this._maxPos = pos;
                        this._minPos = this._maxPos - range;
                        if (this._minPos <= 0) {
                            this._minPos = 0;
                            this._maxPos = this._minPos + range;
                        }
                    }
                }

                this._updateElesPosition();
                this.onRangeChanging();
            }
        }

        private _onDocMouseup(e) {
            var chart, axis, actualMin, actualMax, range;

            if (!this._isVisible) {
                return;
            }

            // fire event
            this._clearInterval();            
            this._isBtnMousedown = false;
            if (this._startPt) {
                this.onRangeChanged();
                this._startPt = null;
                this._movingOffset = null;
            }
            wijmo.removeClass(this._minHandler, _RangeSlider._RANGESLIDER_HANDLE_ACTIVE);
            wijmo.removeClass(this._maxHandler, _RangeSlider._RANGESLIDER_HANDLE_ACTIVE);
        }

        private _onRangeSpaceMousedown(e) {
            var pt = e instanceof MouseEvent ?
                new wijmo.Point(e.pageX, e.pageY) :
                new wijmo.Point(e.changedTouches[0].pageX, e.changedTouches[0].pageY),
                sOffset = wijmo.getElementRect(this._rsContent),
                rOffset = wijmo.getElementRect(this._rangeHandler),
                clickEle = e.srcElement || e.target,
                offset = 0;

            e.stopPropagation();
            e.preventDefault();
            if (clickEle !== this._rsContent && clickEle !== this._rsEle) {
                return;
            }

            if (this._isHorizontal) {
                offset = rOffset.width / sOffset.width;
                if (pt.x < rOffset.left) {
                    offset = -1 * offset;
                } else if (pt.x > rOffset.left + rOffset.width) {
                    offset = 1 * offset;
                }
            } else {
                offset = rOffset.height / sOffset.height;
                if (pt.y < rOffset.top) {
                    offset = 1 * offset;
                } else if (pt.y > rOffset.top + rOffset.height) {
                    offset = -1 * offset;
                }
            }

            if (offset !== 0) {
                this._doSliding(offset, pt);
            }
        }

        private _onRangeMouseleave(e) {
            e.stopPropagation();
            e.preventDefault();

            if (!this._isBtnMousedown) {
                return;
            }
            //fire event
            this._clearInterval();
            this.onRangeChanged();
        }

        private _onBtnMousedown(e) {
            var targetEle = e.srcElement || e.target, offset = 0;

            e.stopPropagation();
            e.preventDefault();

            if (hasClass(targetEle, _RangeSlider._RANGESLIDER_DECBTN)) {
                if (this._minPos === 0) {
                    return;
                }
                offset = -0.05;
            } else if (hasClass(targetEle, _RangeSlider._RANGESLIDER_INCBTN)) {
                if (this._maxPos === 1) {
                    return;
                }
                offset = 0.05;
            }

            this._isBtnMousedown = true;
            if (offset !== 0) {
                this._doSliding(offset);
            }
        }

        _refresh(rsRect?) {
            var sliderOffset = 0, containerOffset = 0,
                slbarCss, rangeSliderEleCss,
                rOffset = wijmo.getElementRect(this._rsContainer);

            if (rsRect) {
                this._rangeSliderRect = rsRect;
            }

            if (!this._rangeSliderRect) {
                return;
            }

            if (this._hasButtons && this._buttonsVisible) {
                this._decBtn.style.display = 'block';
                this._incBtn.style.display = 'block';
                sliderOffset = this._isHorizontal ? this._decBtn.offsetWidth + this._minHandler.offsetWidth / 2 :
                this._decBtn.offsetHeight + this._minHandler.offsetHeight / 2;
            } else {
                if (this._hasButtons) {
                    this._decBtn.style.display = 'none';
                    this._incBtn.style.display = 'none';
                }
                sliderOffset = this._isHorizontal ? this._minHandler.offsetWidth / 2 : this._minHandler.offsetHeight / 2;
            }

            slbarCss = this._getRsRect();
            if (this._isHorizontal) {               
                slbarCss.left -= this._minHandler.offsetWidth / 2;
                slbarCss.width += this._minHandler.offsetWidth;       
                rangeSliderEleCss = { left: sliderOffset, width: slbarCss.width - sliderOffset * 2 };     
            } else {
                //slbarCss.left -= this._minHandler.offsetWidth;
                slbarCss.top -= this._minHandler.offsetHeight/2;
                slbarCss.height += this._minHandler.offsetHeight;
                rangeSliderEleCss = { top: sliderOffset, height: slbarCss.height - sliderOffset * 2 };
            }

            wijmo.setCss(this._rsEle, slbarCss);
            wijmo.setCss(this._rsContent, rangeSliderEleCss);

            rOffset = wijmo.getElementRect(this._rsContent);
            this._plotBox = { x: rOffset.left, y: rOffset.top, width: rOffset.width, height: rOffset.height };
            this._updateElesPosition();
        }

        private _updateElesPosition() {
            var minHandle = this._minHandler, rangeHandle = this._rangeHandler,
                maxHandle = this._maxHandler, box = this._plotBox,
                rangeCss, minCss, rangeCss, maxCss,
                isHorizontal = this._isHorizontal;

            if (box) {
                minCss = isHorizontal ?
                { left: this._minPos * box.width - 0.5 * minHandle.offsetWidth } :
                { top: (1 - this._minPos) * box.height - 0.5 * maxHandle.offsetHeight };

                rangeCss = isHorizontal ?
                { left: this._minPos * box.width, width: (this._maxPos - this._minPos) * box.width } :
                { top: (1 - this._maxPos) * box.height, height: (this._maxPos - this._minPos) * box.height };

                maxCss = isHorizontal ?
                { left: this._maxPos * (box.width) - 0.5 * maxHandle.offsetWidth } :
                { top: (1 - this._maxPos) * box.height - 0.5 * minHandle.offsetHeight };

                this._refreshSlider(minCss, rangeCss, maxCss);
            }
        }

        private _refreshSlider(minCss, rangeCss, maxCss) {
            wijmo.setCss(this._minHandler, minCss);
            wijmo.setCss(this._rangeHandler, rangeCss);
            wijmo.setCss(this._maxHandler, maxCss);
        }

        private _invalidate() {
            var addClass, rmvClass;

            if (!this._rsContainer) {
                return;
            }
            //get needed adding and removing class
            addClass = this._isHorizontal ?
            _RangeSlider._HRANGESLIDER : _RangeSlider._VRANGESLIDER;
            rmvClass = this._isHorizontal?
            _RangeSlider._VRANGESLIDER : _RangeSlider._HRANGESLIDER;

            wijmo.removeClass(this._rsEle, rmvClass);
            wijmo.addClass(this._rsEle, addClass);
      
            //clear inline style
            [this._rsEle, this._rsContent, this._minHandler,
                this._maxHandler, this._rangeHandler].forEach((ele) => {
                ele.removeAttribute("style");
            })
            this._refresh();
        }

        private _changeRange(offset) {
            var range = this._maxPos - this._minPos;

            if ((offset < 0 && this._minPos === 0) || ((offset > 0 && this._maxPos === 1))) {
                return;
            }
            if (offset < 0) {
                this._minPos += offset;
                this._minPos = this._minPos < 0 ? 0 : this._minPos;
                this._maxPos = this._minPos + range;
            } else {
                this._maxPos += offset;
                this._maxPos = this._maxPos > 1 ? 1 : this._maxPos;
                this._minPos = this._maxPos - range;
            }

            this._updateElesPosition();
        }

        private _doSliding(offset, pt?: Point) {
            var sOffset = wijmo.getElementRect(this._rsContent),
                rOffset = wijmo.getElementRect(this._rangeHandler);

            this._clearInterval();

            this._startPt = new Point();
            this._changeRange(offset);
            this.onRangeChanged();
            this._setSlidingInterval(offset, pt);
        }

        private _setSlidingInterval(offset, pt?: Point) {
            var self = this,
                sOffset, rOffset;
           
            this._slidingInterval = window.setInterval(function () {
                if (pt) {
                    //clear the interval when the rangeslider is on mouse position.
                    sOffset = wijmo.getElementRect(self._rsContent);
                    rOffset = wijmo.getElementRect(self._rangeHandler);
                    if (self._isHorizontal) {
                        if (pt.x >= rOffset.left && pt.x <= rOffset.left + rOffset.width) {
                            self._clearInterval();
                            return;
                        }
                    } else {
                        if (pt.y >= rOffset.top && pt.y <= rOffset.top + rOffset.height) {
                            self._clearInterval();
                            return;
                        }
                    }
                }
                self._changeRange(offset);
                self.onRangeChanged();            
            }, 200);
        }

        private _clearInterval() {
            if (this._slidingInterval) {
                window.clearInterval(this._slidingInterval);
            }
        }

        private _getRsRect() {
            var rsRect = this._rangeSliderRect, rect = {};
            if (!rsRect) {
                return;
            }
            ['left', 'top', 'width', 'height'].forEach(function (key) {
                if (rsRect[key]) {
                    rect[key] = rsRect[key];
                }
            })
            return rect;
        }
    }
}
/**
 * Defines classes that add interactive features to charts.
 */
module wijmo.chart.interaction {
    'use strict';

    /**
    * Specifies the orientation of the range selector.
    */
    export enum Orientation {
        /** Horizontal, x-data range. */
        X = 0,
        /** Vertical, y-data range. */
        Y = 1,
    }

    /**
     * The @see:RangeSelector control displays a range selector that allows the user to
     * choose the range of data to display on the specified @see:FlexChart.
     *
     * To use the @see:RangeSelector control, specify the @see:FlexChart
     * control to display the selected range of data.
     *
     * The @see:rangeChanged event is fired when there is a change in min or max value.
     * For example:
     * <pre>
     *  var rangeSelector = new wijmo.chart.interaction.RangeSelector(chart);
     *  rangeSelector.rangeChanged.addHandler(function () {
     *     // perform related updates
     *     // e.g. modify displayed range of another chart
     *     update(rangeSelector.min, rangeSelector.max);
     *   });
     * </pre>
     */
    export class RangeSelector {
        // fields
        private _isVisible = true;                              // range selector is visible or not
        private _min: number = null;                            // minimum value
        private _max: number = null;                            // maximum value
        private _orientation: Orientation = Orientation.X;      // range selector's orientation
        private _seamless: boolean = false;                     // seamless with min and max
        private _minScale: number = 0;                          // minimum range limitation
        private _maxScale: number = 1;                          // maximum range limitation

        private _chart: wijmo.chart.FlexChartCore = null;       // chart host
        private _rangeSelectorEle: HTMLElement = null;          // range selector div element

        private _rangeSlider: _RangeSlider = null;

        /**
         * Initializes a new instance of the @see:RangeSelector class.
         *
         * @param chart The @see:FlexChart that displays the selected range.
         * @param options A JavaScript object containing initialization data for the control.
         */
        constructor(chart: wijmo.chart.FlexChartCore, options?) {
            if (!chart) {
                assert(false, 'The FlexChart cannot be null.');
            }

            this._chart = chart;
            wijmo.copy(this, options);
            this._createRangeSelector();
        }

        /**
         * Gets or sets the visibility of the range selector.
         */
        get isVisible(): boolean {
            return this._isVisible;
        }
        set isVisible(value: boolean) {
            if (value != this._isVisible) {
                this._isVisible = asBoolean(value);
                if (!this._rangeSlider) {
                    return;
                }
                this._rangeSlider.isVisible = value;
            }
        }

        /**
         * Gets or sets the minimum value of the range.
         * If not set, the minimum is calculated automatically.
         */
        get min(): number {
            return this._min;
        }
        set min(value: number) {
            if (value !== this._min) {
                var tmpValue = asNumber(value, true, false); 
                var rangeChged = false;
                if (tmpValue !== null && tmpValue !== undefined && !isNaN(tmpValue) &&
                    this._max !== undefined && this._max !== null) {
                    if (tmpValue <= this._max) {
                        this._min = tmpValue;
                        rangeChged = true;
                    }
                } else {
                    this._min = tmpValue;
                    rangeChged = true;
                }                
                               
                if (!this._rangeSlider) {
                    return;
                }
                if (rangeChged) {
                    this._changeRange();
                }  
            }
        }

        /**
         * Gets or sets the maximum value of the range.
         * If not set, the maximum is calculated automatically.
         */
        get max(): number {
            return this._max;
        }
        set max(value: number) {
            if (value !== this._max) {
                var tmpValue = asNumber(value, true, false); 
                var rangeChged = false;
                if (tmpValue !== null && tmpValue !== undefined && !isNaN(tmpValue)) {
                    if (tmpValue >= this._min) {
                        this._max = tmpValue;
                        rangeChged = true;
                    }
                } else {
                    this._max = tmpValue;
                    rangeChged = true;
                }

                if (!this._rangeSlider) {
                    return;
                }
                if (rangeChged) {
                    this._changeRange();
                }                
            }
        }

        /**
         * Gets or sets the orientation of the range selector.
         */
        get orientation(): Orientation {
            return this._orientation;
        }
        set orientation(value: Orientation) {
            if (value !== this._orientation) {
                this._orientation = asEnum(value, Orientation);
                if (!this._rangeSlider) {
                    return;
                }
                this._rangeSlider.isHorizontal = value === Orientation.X;
            }
        }

        /**
         * Gets or sets a value that determines whether the minimal and maximal 
         * handler will move seamlessly.
         */
        get seamless(): boolean {
            return this._seamless;
        }
        set seamless(value: boolean) {
            if (value !== this._seamless) {
                this._seamless = asBoolean(value, true);
                if (!this._rangeSlider) {
                    return;
                }
                this._rangeSlider.seamless = value;
            }
        }

        /**
         * Gets or sets the minimum range scale of the range selector.
         * The minimum scale is between 0 and 1.
         */
        get minScale(): number {
            return this._minScale;
        }
        set minScale(value: number) {
            if (value <= 1 && value >= 0 && value != this._minScale && value < this._maxScale) {
                this._minScale = asNumber(value);
                if (!this._rangeSlider) {
                    return;
                }
                this._rangeSlider.minScale = asNumber(value);
                this._updateMinAndMaxWithScale(true);
            }
        }

        /**
         * Gets or sets the maximum range scale of the range selector.
         * The maximum scale is between 0 and 1.
         */
        get maxScale(): number {
            return this._maxScale;
        }
        set maxScale(value: number) {
            if (value <=1 && value >= 0 && value != this._maxScale && value > this._minScale) {
                this._maxScale = asNumber(value);
                if (!this._rangeSlider) {
                    return;
                }
                this._rangeSlider.maxScale = asNumber(value);
                this._updateMinAndMaxWithScale(true);
            }
        }

        /**
         * Occurs after the range changes.
         */
        rangeChanged = new Event();

        /**
         * Raises the @see:rangeChanged event.
         */
        onRangeChanged(e?: EventArgs) {
            this.rangeChanged.raise(this, e);
        }

        /**
         * Removes the @see:RangeSelector control from the chart.
         */
        remove() {
            if (this._rangeSelectorEle) {
                this._chart.hostElement.removeChild(this._rangeSelectorEle);
                this._switchEvent(false);
                this._rangeSelectorEle = null;
                this._rangeSlider = null;
            }
        }

        private _createRangeSelector() {
            var chart = this._chart,
                chartHostEle = chart.hostElement,
                isHorizontal = this._orientation === Orientation.X;

            this._rangeSelectorEle = createElement('<div class="wj-chart-rangeselector-container"></div>');
            this._rangeSlider = new _RangeSlider(this._rangeSelectorEle,
                false, //no range click
                false, //no buttons
                {
                    //options settings
                    isHorizontal: isHorizontal,
                    isVisible: this._isVisible,
                    seamless: this._seamless
                }
                );
            chartHostEle.appendChild(this._rangeSelectorEle);

            this._switchEvent(true);
        }

        private _switchEvent(isOn: boolean) {
            var eventHandler = isOn ? 'addHandler' : 'removeHandler';

            if (this._chart.hostElement) {
                this._rangeSlider.rangeChanged[eventHandler](this._updateRange, this);                
                this._chart.rendered[eventHandler](this._refresh, this);
            }
        }

        private _refresh() {
            var chartHostEle = this._chart.hostElement,
                pa, pOffset, plotBox, rOffset = wijmo.getElementRect(this._rangeSelectorEle);

            pa = chartHostEle.querySelector('.' + FlexChart._CSS_PLOT_AREA);
            pOffset = wijmo.getElementRect(pa);
            plotBox = pa.getBBox();

            if (plotBox && plotBox.width && plotBox.height) {
                this._adjustMinAndMax();

                // position and sized rangeslider
                this._rangeSlider._refresh(
                    {
                        left: plotBox.x,
                        top: pOffset.top - rOffset.top,
                        width: plotBox.width,
                        height: plotBox.height
                    });
            }
        }

        private _adjustMinAndMax() {
            var self = this, chart = self._chart, rangeSlider = self._rangeSlider,
                min = self._min, max = self._max,
                axis = self._orientation === Orientation.X ? chart.axisX : chart.axisY,
                actualMin = isDate(axis.actualMin) ? axis.actualMin.valueOf() : axis.actualMin,
                actualMax = isDate(axis.actualMax) ? axis.actualMax.valueOf() : axis.actualMax,
                range = actualMax - actualMin;

            self._min = (min === null || isNaN(min) || min === undefined || min < actualMin || min > actualMax) ? actualMin : min;
            self._max = (max === null || isNaN(max) || max === undefined || max < actualMin || max > actualMax) ? actualMax : max;

            // removed
            //rangeSlider._minPos = (self._min - actualMin) / range;
            //rangeSlider._maxPos = (self._max - actualMin) / range;
            //
            // The previous code is only for regular(linear) axis.
            // Take into account non-linear axis:
            var plotRect = this._chart._plotRect;
            if (!plotRect) {
                return;
            }
            if (this._orientation === Orientation.X) {
                var minPos = (axis.convert(self._min) - plotRect.left) / plotRect.width;
                rangeSlider._minPos = minPos;
                var maxPos = (axis.convert(self._max) - plotRect.left) / plotRect.width;
                rangeSlider._maxPos = maxPos;
            } else {
                var minPos = (plotRect.top - axis.convert(self._min)) / plotRect.height + 1;
                rangeSlider._minPos = minPos;
                var maxPos = (plotRect.top - axis.convert(self._max)) / plotRect.height + 1;
                rangeSlider._maxPos = maxPos;
            }

            this._updateMinAndMaxWithScale(false);
        }

        private _updateMinAndMaxWithScale(fireEvent:boolean) {
            var rangeSlider = this._rangeSlider, max, updated = false;

            if (this._minScale !== 0 &&
                rangeSlider._minPos + this._minScale > rangeSlider._maxPos) {
                max = rangeSlider._minPos + this._minScale;
                if (max > 1) {
                    rangeSlider._maxPos = 1;
                    rangeSlider._minPos = 1 - this._minScale;
                } else {
                    rangeSlider._maxPos = max;
                }
                updated = true;
            }

            if (this._maxScale !== 1 &&
                rangeSlider._minPos + this._maxScale < rangeSlider._maxPos) {
                max = rangeSlider._minPos + this._maxScale;
                if (max > 1) {
                    rangeSlider._maxPos = 1;
                    rangeSlider._minPos = 1 - this._maxScale;
                } else {
                    rangeSlider._maxPos = max;
                }
                updated = true;
            }

            if (updated) {
                this._min = this._getMinAndMax()['min'];
                this._max = this._getMinAndMax()['max'];
                if (fireEvent) {
                    if (!this._rangeSelectorEle) {
                        return;
                    }

                    this._rangeSlider._refresh();
                    this.onRangeChanged();
                }
            }

        }

        private _changeRange() {
            this._adjustMinAndMax();

            if (!this._rangeSelectorEle) {
                return;
            }

            this._rangeSlider._refresh();
            this.onRangeChanged();
        }

        private _updateRange() {
            var rangeSlider = this._rangeSlider,
                chart, axis, actualMin, actualMax, range;
            chart = this._chart;
            axis = this._orientation === Orientation.X ? chart.axisX : chart.axisY;
            
            //raise event

            // removed
            // this._min = actualMin + this._minPos * range;
            // this._max = actualMin + this._maxPos * range;
            //
            // The previous code is only for regular(linear) axis.
            // take into account non-linear axis
            this._min = this._getMinAndMax()['min'];
            this._max = this._getMinAndMax()['max'];
            this.onRangeChanged();
        }

        private _getMinAndMax(): any {
            var rangeSlider = this._rangeSlider,
                chart = this._chart,
                axis = this._orientation === Orientation.X ? chart.axisX : chart.axisY,
                plotRect = this._chart._plotRect,
                min, max;
            if (this._orientation === Orientation.X) {
                min = axis.convertBack(plotRect.left + rangeSlider._minPos * plotRect.width);
                max = axis.convertBack(plotRect.left + rangeSlider._maxPos * plotRect.width);
            } else {
                min = axis.convertBack(plotRect.top + (1 - rangeSlider._minPos) * plotRect.height);
                max = axis.convertBack(plotRect.top + (1 - rangeSlider._maxPos) * plotRect.height);
            }
            return { min:min, max:max };
        }
    } 
}
module wijmo.chart.interaction {

    /**
      * Specifies the mouse action of the chart gestures.
      */
    export enum MouseAction {
        /** Zoom chart by mouse. */
        Zoom = 0,
        /** Pan chart by mouse. */
        Pan = 1,
    }

    /**
      * Specifies the interactive axes of the chart gestures.
      */
    export enum InteractiveAxes {
        /** Interactive Axis X. */
        X = 0,
        /** Interactive Axis Y. */
        Y = 1,
        /** Interactive Both Axis X and Axis Y. */
        XY = 2,
    }    

    /**
     * The @see:ChartGestures control allows the user to zoom or pan on  
     * the specified @see:FlexChart.
     *
     * To use the @see:ChartGestures control, specify the @see:FlexChart
     * control on which to zoom or pan.
     *
     * <pre>
     *  var chartGestures = new wijmo.chart.interaction.ChartGestures(chart);
     * </pre>
     */
    export class ChartGestures {
        static _CSS_ZOOM = 'wj-zoom';
        static _CSS_ZOOM_OVERLAY = 'wj-zoom-overlay';
        static _CSS_PANABLE = 'wj-panable';
        static _CSS_TOUCH_DISABLED = 'wj-flexchart-touch-disabled';
        static _CSS_BLOCK_INTERACTION = 'wj-block-other-interaction';

        private _chart: wijmo.chart.FlexChartCore = null;
        private _zoomEle: HTMLElement = null;
        private _overlayEle: HTMLElement = null;
        private _zoomEleOffset: any;

        //events
        private _wrapperMousedown = null;
        private _wrapperMouseMove = null;
        private _wrapperMouseup = null;
        private _wrapperPointerdown = null;
        private _wrapperPointerMove = null;
        private _wrapperPointerup = null;
        private _wrapperTouchStart = null;
        private _wrapperTouchMove = null;
        private _wrapperTouchEnd = null;
        private _wrapperMouseWheel = null;
          
        //help                         
        private _plotBox: any;
        private _startFirstPt: Point = null;
        private _minX: number = null;
        private _maxX: number = null;
        private _minY: number = null;
        private _maxY: number = null;
        private _seriesGroup: any;
        private _threadHold: number = 20;
        private _scaling: boolean;
        private _panning: boolean;
        private _startDistance: any;
        private _clip = {};
        private _selection = {};
        private _startPointers = [];
        private _mvPointers = [];
        private _plotOffset: any;
        private _endPoint: Point;
        private _pinchStartEvents = [];
        private _minXRange: number = null;
        private _minYRange: number = null;
        private _innerUpdating: boolean = false;
        private _lastMinX: number = null;
        private _lastMaxX: number = null;
        private _lastMinY: number = null;
        private _lastMaxY: number = null;

        //option
        private _mouseAction = MouseAction.Zoom;
        private _interactiveAxes = InteractiveAxes.X;
        private _enable = true;
        private _scaleX: number = 1;
        private _scaleY: number = 1;
        private _posX: number = 0;
        private _posY: number = 0;

        /**
         * Initializes a new instance of the @see:ChartGestures class.
         *
         * @param chart The @see:FlexChart that allows the user to zoom or pan.
         * @param options A JavaScript object containing initialization data for the control.
         */
        constructor(chart: wijmo.chart.FlexChartCore, options?) {
            if (!chart) {
                assert(false, 'The FlexChart cannot be null.');
            }

            this._chart = chart;
            wijmo.copy(this, options);
            this._initialize();
        }

        /**
          * Gets or sets the mouse action of the ChartGestures.
          */
        get mouseAction(): MouseAction {
            return this._mouseAction;
        }

        set mouseAction(value: MouseAction) {
            if (value !== this._mouseAction) {
                this._mouseAction = asEnum(value, MouseAction);
            }
        }

        /**
          * Gets or sets the interactive axes of the ChartGestures.
          */
        get interactiveAxes(): InteractiveAxes {
            return this._interactiveAxes;
        }

        set interactiveAxes(value: InteractiveAxes) {
            if (value !== this._interactiveAxes) {
                this._interactiveAxes = asEnum(value, InteractiveAxes);
            }
        }

        /**
          * Gets or sets the enable of the ChartGestures.
          */
        get enable(): boolean {
            return this._enable;
        }

        set enable(value: boolean) {
            if (value !== this._enable) {
                this._enable = asBoolean(value, true);
            }
        }

        /**
          * Gets or sets the initial scale of axis X.
          * The scale should be more than 0 and less than or equal to 1.
          * The scale specifies which part of the range between Min and Max
          * is shown. When scale is 1 (default value), the whole axis range
          * is visible.
          */
        get scaleX(): number {
            return this._scaleX;
        }

        set scaleX(value: number) {
            if (value !== this._scaleX) {
                if (value < 0) {
                    this._scaleX = 0;
                } else if (value > 1) {
                    this._scaleX = 1;
                } else {
                    this._scaleX = asNumber(value);
                } 
                if (this._seriesGroup) {
                    this._initAxisRangeWithPosAndScale(true);     
                }
            }
        }

        /**
          * Gets or sets the initial scale of axis Y.
          * The scale should be more than 0 and less than or equal to 1.
          * The scale specifies which part of the range between Min and Max
          * is shown. When scale is 1 (default value), the whole axis range
          * is visible.
          */
        get scaleY(): number {
            return this._scaleY;
        }
        set scaleY(value: number) {
            if (value !== this._scaleY) {
                if (value < 0) {
                    this._scaleY = 0;
                } else if (value > 1) {
                    this._scaleY = 1;
                } else {
                    this._scaleY = asNumber(value);
                }
                if (this._seriesGroup) {
                    this._initAxisRangeWithPosAndScale(false);  
                }
            }
        }

        /**
          * Gets or sets the initial position of the axis X.
          * The value represents initial position on the axis when the Scale
          * is less than 1. Otherwise, the Value has no effect. The Value should
          * lie between 0 to 1.
         */
        get posX(): number {
            return this._posX;
        }
        set posX(value: number) {
            if (value !== this._posX) {
                if (value < 0) {
                    this._posX = 0;
                } else if (value > 1) {
                    this._posX = 1;
                } else {
                    this._posX = asNumber(value);
                } 
                if (this._seriesGroup) {
                    this._initAxisRangeWithPosAndScale(true);   
                }
            }
        }

        /**
          * Gets or sets the initial position of the axis Y.
          * The value represents initial position on the axis when the Scale
          * is less than 1. Otherwise, the Value has no effect. The Value should
          * lie between 0 to 1.
         */
        get posY(): number {
            return this._posY;
        }
        set posY(value: number) {
            if (value !== this._posY) {
                if (value < 0) {
                    this._posY = 0;
                } else if (value > 1) {
                    this._posY = 1;
                } else {
                    this._posY = asNumber(value);
                }
                if (this._seriesGroup) {
                    this._initAxisRangeWithPosAndScale(false);
                }
            }
        }

        /**
         * Removes the @see:ChartGestures control from the chart.
         */
        remove() {
            if (this._zoomEle) {
                this._chart.hostElement.removeChild(this._zoomEle);
                this._zoomEle = null;
            }
            wijmo.removeClass(this._chart.hostElement, ChartGestures._CSS_TOUCH_DISABLED);
            this._switchEvent(false);
            this._wrapperMousedown = null;
            this._wrapperMouseMove = null;
            this._wrapperMouseup = null;
            this._wrapperPointerdown = null;
            this._wrapperPointerMove = null;
            this._wrapperPointerup = null;
            this._wrapperTouchStart = null;
            this._wrapperTouchMove = null;
            this._wrapperTouchEnd = null;
            this._wrapperMouseWheel = null;
        }

        /** 
         * Reset the axis of the chart.
         */
        reset() {
            var chart = this._chart,
                axisX = chart.axisX,
                axisY = chart.axisY;

            if (this._maxX) {
                axisX.max = this._maxX;
            }
            if (this._minX) {
                axisX.min = this._minX;
            }
            if (this._maxY) {
                axisY.max = this._maxY;
            }
            if (this._minY) {
                axisY.min = this._minY;
            }
            // Axis X
            this._initAxisRangeWithPosAndScale(true);
            // Axis Y
            this._initAxisRangeWithPosAndScale(false);
        }
        
        /** 
         * Refreshes the @see:FlexChart with the gestures settings.
         */
        _refreshChart() {
            var chart = this._chart,
                axisX = chart.axisX, axisY = chart.axisY;

            this._minX = this._getAxisMin(axisX);
            this._maxX = this._getAxisMax(axisX);
            this._minY = this._getAxisMin(axisY);
            this._maxY = this._getAxisMax(axisY); 

            //setting the min&max scale range
            this._minXRange = (this._maxX - this._minX) * 0.005;
            this._minYRange = (this._maxY - this._minY) * 0.005;
 
            // initialize Axis X
            this._initAxisRangeWithPosAndScale(true);
            // initialize Axis Y
            this._initAxisRangeWithPosAndScale(false);
        }

        private _initialize() {
            var chart = this._chart,
                chartHostEle = chart.hostElement;

            this._zoomEle = createElement('<div class="' + ChartGestures._CSS_ZOOM + '">' +
                '<div class="' + ChartGestures._CSS_ZOOM_OVERLAY + '"></div>');
            this._zoomEle.style.visibility = 'visible';

            chartHostEle.appendChild(this._zoomEle);
            wijmo.addClass(chartHostEle, ChartGestures._CSS_TOUCH_DISABLED);
            this._overlayEle = <HTMLElement>this._zoomEle.querySelector('.' + ChartGestures._CSS_ZOOM_OVERLAY);
           
            //bind event
            this._wrapperMousedown = this._onMousedown.bind(this);
            this._wrapperMouseMove = this._onMouseMove.bind(this);
            this._wrapperMouseup = this._onMouseup.bind(this);
            this._wrapperPointerdown = this._onPointerdown.bind(this);
            this._wrapperPointerMove = this._onPointerMove.bind(this);
            this._wrapperPointerup = this._onPointerup.bind(this);
            this._wrapperMouseWheel = this._onMouseWheel.bind(this);
            this._wrapperTouchStart = this._onTouchStart.bind(this);
            this._wrapperTouchMove = this._onTouchMove.bind(this);
            this._wrapperTouchEnd = this._onTouchEnd.bind(this);
            this._switchEvent(true);
        }

        private _switchEvent(isOn: boolean) {
            var chartHostEle = this._chart.hostElement,
                eventListener = isOn ? 'addEventListener' : 'removeEventListener',
                eventHandler = isOn ? 'addHandler' : 'removeHandler';

            if (chartHostEle) {
                chartHostEle[eventListener]('mousedown', this._wrapperMousedown);
                chartHostEle[eventListener]('mousemove', this._wrapperMouseMove);
                document[eventListener]('mouseup', this._wrapperMouseup);
                if ('onpointerdown' in window) {
                    chartHostEle[eventListener]('pointerdown', this._wrapperPointerdown);
                    chartHostEle[eventListener]('pointermove', this._wrapperPointerMove);
                    document[eventListener]('pointerup', this._wrapperPointerup);
                }
                // change to 'wheel' event
               // if ('onmousewheel' in window) {
                //    chartHostEle[eventListener]('mousewheel', this._wrapperMouseWheel);
                //}
                chartHostEle[eventListener]('wheel', this._wrapperMouseWheel);
                if ('ontouchstart' in window) {
                    chartHostEle[eventListener]('touchstart', this._wrapperTouchStart);
                    chartHostEle[eventListener]('touchmove', this._wrapperTouchMove);
                    document[eventListener]('touchend', this._wrapperTouchEnd);
                }
                this._chart.rendered[eventHandler](this._refresh, this);
            }
        }

        private _refresh() {
            var chart = this._chart,
                axisX = chart.axisX, axisY = chart.axisY,
                chartHostEle = chart.hostElement, pa,
                rangeXChged, rangeYChged;

            this._seriesGroup = chartHostEle.querySelector('.wj-series-group');
            pa = chartHostEle.querySelector('.' + FlexChart._CSS_PLOT_AREA);
            this._plotOffset = wijmo.getElementRect(pa);
            this._plotBox = pa.getBBox();
            this._zoomEleOffset = wijmo.getElementRect(this._zoomEle);
            if (this._overlayEle) {
                this._overlayEle.removeAttribute('style');
            }

            if (this._innerUpdating) {
                this._innerUpdating = false;
                return;
            }

            rangeXChged = false;
            rangeYChged = false;
           
            if (this._minX === null || isNaN(this._minX) || this._minX === 0
                || this._minX === -1 || this._lastMinX !== this._getAxisMin(axisX)) {
                this._minX = this._getAxisMin(axisX);
                if (this._minX !== null && !isNaN(this._minX) && this._minX !== 0 && this._minX !== -1) {
                    rangeXChged = true;
                }                
            }
            if (this._maxX === null || isNaN(this._maxX) || this._maxX === 0
                || this._maxX === -1 || this._lastMaxX !== this._getAxisMax(axisX)) {
                this._maxX = this._getAxisMax(axisX);
                if (this._maxX !== null && !isNaN(this._maxX) && this._maxX !== 0 && this._maxX !== -1) {
                    rangeXChged = true;
                }   
            }

            if (this._minY === null || isNaN(this._minY) || this._lastMinY !== this._getAxisMin(axisY)) {
                this._minY = this._getAxisMin(axisY);               
                if (!isNaN(this._minY)) {
                    rangeYChged = true;
                }                
            }
            if (this._maxY === null || isNaN(this._maxY) || this._lastMaxY !== this._getAxisMax(axisY)) {
                this._maxY = this._getAxisMax(axisY);
                if (!isNaN(this._maxY)) {
                    rangeYChged = true;
                } 
            }
            
            //setting the min&max scale range
            this._minXRange = (this._maxX - this._minX) * 0.005;
            this._minYRange = (this._maxY - this._minY) * 0.005;

            //initialize axisX and axisY      
            if (rangeXChged && this._scaleX !== null && this._scaleX !== undefined && this._scaleX !== 1 &&
                this._posX !== null && this._posX !== undefined && this._posX !== 0) {
                this._initAxisRangeWithPosAndScale(true);              
            }
            if (rangeYChged && this._scaleY !== null && this._scaleY !== undefined && this._scaleY !== 1 &&
                this._posY !== null && this._posY !== undefined && this._posY !== 0) {
                this._initAxisRangeWithPosAndScale(false);              
            }
        }

        /** mouse event*/
        private _onMousedown(e: MouseEvent) {
            if (!this._enable) {
                return;
            }
            this._disabledOthersInteraction(true);
            this._mouseDown(e);
            e.preventDefault();
        }

        private _onMouseMove(e: MouseEvent) {
            if (!this._enable) {
                return;
            }
            this._mouseMove(e);
            e.preventDefault();
        }

        private _onMouseup(e: MouseEvent) {
            if (!this._enable) {
                return;
            }
            this._mouseup(e);
            this._disabledOthersInteraction(false);
            //e.preventDefault();
        }

        private _onMouseWheel(e: WheelEvent) {
            //var delta = e.detail || e.wheelDelta,
            var delta = -e.deltaY,
                chg = delta > 0 ? 0.05 : -0.05;
            if (!this._enable) {
                return;
            }
            this._scaling = true;
            if (this._interactiveAxes === InteractiveAxes.X ||
                this._interactiveAxes === InteractiveAxes.XY) {
                this._updateAxisByChg(true, chg, -chg);
            }

            if (this._interactiveAxes === InteractiveAxes.Y ||
                this._interactiveAxes === InteractiveAxes.XY) {
                this._updateAxisByChg(false, chg, -chg);
            }
            this._scaling = false;
            e.preventDefault();
        }

        private _mouseDown(e: MouseEvent) {
            this._startFirstPt = this._getPoint(e);
            this._updatePoint(this._startFirstPt);
            if (this._mouseAction === MouseAction.Zoom) {
                this._initOverlay();
            } else {
                this._seriesGroup.setAttribute('clip-path', 'url(#' + this._chart._plotrectId + ')');
                wijmo.toggleClass(this._chart.hostElement, ChartGestures._CSS_PANABLE, this._mouseAction === MouseAction.Pan);
            }
        }

        private _mouseMove(e: MouseEvent) {
            var mvPt;
            if (!this._startFirstPt) {
                return;
            }

            mvPt = this._getPoint(e);
            this._updatePoint(mvPt);
            this._endPoint = new Point(mvPt.x, mvPt.y);

            if (this._mouseAction === MouseAction.Zoom) {
                this._updateOverLay(mvPt);
            } else {
                this._panning = true;
                this._panningChart(mvPt.x - this._startFirstPt.x, mvPt.y - this._startFirstPt.y);
            }
        }

        private _mouseup(e: MouseEvent) {
            var endPt = this._endPoint,
                axisX = this._chart.axisX;

            if (!this._startFirstPt || !endPt) {
                wijmo.removeClass(this._chart.hostElement, ChartGestures._CSS_PANABLE);
                this._reset();
                return;
            }

            if (this._mouseAction === MouseAction.Zoom) {
                this._zoomedChart(endPt);
                this._reset();
            } else {
                this._pannedChart(endPt.x - this._startFirstPt.x, endPt.y - this._startFirstPt.y);
                this._reset();
            }
            wijmo.removeClass(this._chart.hostElement, ChartGestures._CSS_PANABLE);
        }

        /** ms pointer event*/
        private _onPointerdown(e: PointerEvent) {
            if (!this._enable) {
                return;
            }
            this._disabledOthersInteraction(true);
            switch (e.pointerType) {
                case "touch":
                    this._pointerDown(e);
                    break;
                case "mouse":
                    this._mouseDown(e);
                    break;
            }
            e.preventDefault();
        }

        private _onPointerMove(e: PointerEvent) {
            if (!this._enable) {
                return;
            }
            switch (e.pointerType) {
                case "touch":
                    this._pointerMove(e);
                    break;
                case "mouse":
                    this._mouseMove(e);
                    break;
            }
            e.preventDefault();
        }

        private _onPointerup(e: PointerEvent) {
            if (!this._enable) {
                return;
            }
            switch (e.pointerType) {
                case "touch":
                    this._pointerUp(e);
                    break;
                case "mouse":
                    this._mouseup(e);
                    break;
            }
            this._disabledOthersInteraction(false);
            e.preventDefault();
        }

        private _pointerDown(e: any) {
            if (e.preventManipulation)
                e.preventManipulation();

            this._seriesGroup.setAttribute('clip-path', 'url(#' + this._chart._plotrectId + ')');
            this._startPointers.push({ id: e.pointerId, x: e.pageX, y: e.pageY });
            if (this._startPointers.length === 1) {
                this._scaling = false;
                this._panning = true;
            } else if (this._startPointers.length === 2) {
                this._panning = false;
                this._scaling = true;
                this._startDistance = {
                    x: this._startPointers[0].x - this._startPointers[1].x,
                    y: this._startPointers[0].y - this._startPointers[1].y,
                };
            }
        }

        private _pointerMove(e: any) {
            var pt1, pt2,
                mvPt = new Point(e.pageX, e.pageY),
                rNowCordinate, rStartCordinate,
                offset = {}, scale = {};
            if (e.preventManipulation)
                e.preventManipulation();

            if (this._panning) {
                if (!this._pointInPlotArea(mvPt)) {
                    return;
                }
                this._endPoint = new Point(e.pageX, e.pageY);
                this._panningChart(this._endPoint.x - this._startPointers[0].x, this._endPoint.y - this._startPointers[0].y);
            } else if (this._scaling) {
                pt1 = this._startPointers[0].id + '';
                pt2 = this._startPointers[1].id + '';

                this._mvPointers[e.pointerId + ''] = { x: e.pageX, y: e.pageY };

                if (this._mvPointers[pt1] && this._mvPointers[pt2]) {
                    if (Math.abs(this._startDistance.x) > this._threadHold &&
                        this._interactiveAxes !== InteractiveAxes.Y) {
                        rNowCordinate = this._mvPointers[pt1].x - this._plotOffset.left;
                        rStartCordinate = this._startPointers[0].x - this._plotOffset.left;
                        scale['x'] = Math.abs((this._mvPointers[pt1].x - this._mvPointers[pt2].x) / this._startDistance.x);
                        offset['x'] = rNowCordinate - scale['x'] * rStartCordinate;
                        this._clip['x'] = (this._plotBox.x - rNowCordinate) / scale['x'] + rStartCordinate;
                        this._selection['w'] = this._plotBox.width / scale['x'];
                    }
                    if (Math.abs(this._startDistance.y) > this._threadHold &&
                        this._interactiveAxes !== InteractiveAxes.X) {
                        rNowCordinate = this._mvPointers[pt1].y - this._plotOffset.top;
                        rStartCordinate = this._startPointers[0].y - this._plotOffset.top;
                        scale['y'] = Math.abs((this._mvPointers[pt1].y - this._mvPointers[pt2].y) / this._startDistance.y);
                        offset['y'] = rNowCordinate - scale['y'] * rStartCordinate;
                        this._clip['y'] = (this._plotBox.y - rNowCordinate) / scale['y'] + rStartCordinate;
                        this._selection['h'] = this._plotBox.height / scale['y'];
                    }
                    this._scalingChart(scale, offset);
                }
            }
        }

        private _pointerUp(e: any) {
            if (e.preventManipulation)
                e.preventManipulation();

            if (this._panning) {
                if (this._endPoint) {
                    this._pannedChart(this._endPoint.x - this._startPointers[0].x, this._endPoint.y - this._startPointers[0].y);
                }                
                this._reset();
            } else if (this._scaling) {
                this._scaledChart(e);
                this._reset();
            }
        }

        /** touch event*/
        private _onTouchStart(e: any) {
            if (!this._enable) {
                return;
            }
            this._disabledOthersInteraction(true);
            if (e.touches.length == 1) {
                this._scaling = false;
                this._panning = true;
                this._startFirstPt = this._getPoint(e);
            } else if (e.touches.length == 2) {
                this._pinchStartEvents = this._getTouchPair(e);
                this._startDistance = this._touchDistance(e);
                this._panning = false;
                this._scaling = true;
            }
            if (this._seriesGroup) {
                this._seriesGroup.setAttribute('clip-path', 'url(#' + this._chart._plotrectId + ')');
            }

            e.preventDefault();
            return true;
        }

        private _onTouchMove(e: any) {
            if (!this._enable) {
                return;
            }
            var scale = {}, offset = {},
                touchs = e.touches[0],
                mvPt = new Point(touchs.pageX, touchs.pageY),
                rNowCordinate, rStartCordinate,
                nowDist, nowPos, startPos;
            e.preventDefault();

            if (this._panning) {
                if (this._startFirstPt) {
                    if (!this._pointInPlotArea(mvPt)) {
                        return;
                    }
                    this._endPoint = new Point(touchs.pageX, touchs.pageY);
                    this._panningChart(this._endPoint.x - this._startFirstPt.x,
                        this._endPoint.y - this._startFirstPt.y);
                }
            } else if (this._scaling) {

                nowDist = this._touchDistance(e);
                nowPos = this._getTouchPair(e)[0];
                startPos = this._pinchStartEvents[0];
                
                //horizontal
                if (Math.abs(this._startDistance.x) > this._threadHold &&
                    this._interactiveAxes !== InteractiveAxes.Y) {
                    rNowCordinate = nowPos.pageX - this._plotOffset.left;
                    rStartCordinate = startPos.pageX - this._plotOffset.left;
                    scale['x'] = Math.abs(nowDist.x / this._startDistance.x);
                    offset['x'] = rNowCordinate - (scale['x'] * rStartCordinate);
                    this._clip['x'] = (this._plotBox.x - rNowCordinate) / scale['x'] + rStartCordinate;
                    this._selection['w'] = this._plotBox.width / scale['x'];
                }
               
                //vertical
                if (Math.abs(this._startDistance.y) > this._threadHold &&
                    this._interactiveAxes !== InteractiveAxes.X) {
                    rNowCordinate = nowPos.pageY - this._plotOffset.top;
                    rStartCordinate = startPos.pageY - this._plotOffset.top;
                    scale['y'] = Math.abs(nowDist.y / this._startDistance.y);
                    offset['y'] = rNowCordinate - (scale['y'] * rStartCordinate);
                    this._clip['y'] = (this._plotBox.y - rNowCordinate) / scale['y'] + rStartCordinate;
                    this._selection['h'] = this._plotBox.height / scale['y'];
                }
                this._scalingChart(scale, offset);
            }

            return true;
        }

        private _onTouchEnd(e: any) {
            if (!this._enable) {
                return;
            }
            var endPt = this._endPoint;

            if (this._panning) {
                if (!this._startFirstPt || !endPt) {
                    this._reset();
                    return;
                }
                this._pannedChart(endPt.x - this._startFirstPt.x, endPt.y - this._startFirstPt.y);
            } else if (this._scaling) {
                this._scaledChart(e);
            }
            this._reset();
            this._disabledOthersInteraction(false);        
            //e.preventDefault();
            return true;
        }

        /** help method of zooming chart by mouse */
        private _initOverlay() {
            this._zoomEle.style.visibility = 'visible';
            switch (this._interactiveAxes) {
                case InteractiveAxes.X:
                    this._overlayEle.style.left = (this._startFirstPt.x - this._zoomEleOffset.left) + 'px';
                    this._overlayEle.style.top = (this._plotOffset.top - this._zoomEleOffset.top) + 'px';
                    break;
                case InteractiveAxes.Y:
                    this._overlayEle.style.left = (this._plotBox.x) + 'px';
                    this._overlayEle.style.top = (this._startFirstPt.y - this._zoomEleOffset.top) + 'px';
                    break;
                case InteractiveAxes.XY:
                    this._overlayEle.style.left = (this._startFirstPt.x - this._zoomEleOffset.left) + 'px';
                    this._overlayEle.style.top = (this._startFirstPt.y - this._zoomEleOffset.top) + 'px';
                    break;
            }
        }

        private _updateOverLay(mvPt: Point) {
            var distanceX = this._startFirstPt.x - mvPt.x,
                distanceY = this._startFirstPt.y - mvPt.y,
                style = {};

            switch (this._interactiveAxes) {
                case InteractiveAxes.X:
                    if ((Math.abs(distanceX)) < this._threadHold) {
                        return;
                    }
                    style = distanceX <= 0 ?
                        { width: Math.abs(distanceX) + 'px', height: this._plotBox.height + 'px' } :
                        { left: (mvPt.x - this._zoomEleOffset.left) + 'px', width: distanceX + 'px', height: this._plotBox.height + 'px' };
                    break;
                case InteractiveAxes.Y:
                    if ((Math.abs(distanceY)) < this._threadHold) {
                        return;
                    }
                    style = distanceY <= 0 ?
                        { height: Math.abs(distanceY) + 'px', width: this._plotBox.width + 'px' } :
                        { top: (mvPt.y - this._zoomEleOffset.top) + 'px', height: distanceY + 'px', width: this._plotBox.width + 'px' };
                    break;
                case InteractiveAxes.XY:
                    //horizontal
                    if ((Math.abs(distanceX)) >= this._threadHold) {
                        style['width'] = Math.abs(distanceX) + 'px';
                        if (distanceX > 0) {
                            style['left'] = (mvPt.x - this._zoomEleOffset.left) + 'px';
                        }
                    }
                    //vertical
                    if ((Math.abs(distanceY)) >= this._threadHold) {
                        style['height'] = Math.abs(distanceY) + 'px';
                        if (distanceY > 0) {
                            style['top'] = (mvPt.y - this._zoomEleOffset.top) + 'px';
                        }
                    }
                    break;
            }
            wijmo.setCss(this._overlayEle, style);
        }

        _updatePoint(mvPt: Point) {
            var plotRect = this._plotOffset;
            if (mvPt.x < plotRect.left) {
                mvPt.x = plotRect.left;
            }
            if (mvPt.x > plotRect.left + plotRect.width) {
                mvPt.x = plotRect.left + plotRect.width;
            }
            if (mvPt.y < plotRect.top) {
                mvPt.y = plotRect.top;
            }
            if (mvPt.y > plotRect.top + plotRect.height) {
                mvPt.y = plotRect.top + plotRect.height;
            }
        }

        _pointInPlotArea(mvPt: Point) {
            var plotRect = this._plotOffset;
            if (mvPt.x >= plotRect.left && mvPt.x <= plotRect.left + plotRect.width &&
                mvPt.y >= plotRect.top && mvPt.y <= plotRect.top + plotRect.height) {
                return true;
            }
            return false;
        }

        private _zoomedChart(endPt: Point) {
            if (!endPt) {
                return;
            }
            //horizontal
            if (this._interactiveAxes === InteractiveAxes.X ||
                this._interactiveAxes === InteractiveAxes.XY) {
                this._zoomedAxis(endPt, true);
            }
            //vertical
            if (this._interactiveAxes === InteractiveAxes.Y ||
                this._interactiveAxes === InteractiveAxes.XY) {
                this._zoomedAxis(endPt, false);
            }
            this._startFirstPt = null;
            //this._refresh();
        }

        private _zoomedAxis(endPt: Point, isX: boolean) {
            var htStart, htEnd, min, max,
                axis = isX ? this._chart.axisX : this._chart.axisY,
                xy = isX ? 'x' : 'y', lt = isX ? 'left' : 'top';

            if (!endPt) {
                return;
            }
            if (Math.abs(this._startFirstPt[xy] - endPt[xy]) > this._threadHold) {

                min = axis.convertBack(this._startFirstPt[xy] - this._plotOffset[lt] + this._plotBox[xy]);
                max = axis.convertBack(endPt[xy] - this._plotOffset[lt] + this._plotBox[xy]);

                if (max - min !== 0) {
                    this._updateAxisRange(axis, Math.min(min, max), Math.max(min, max));
                }
            }
        }

        private _panningChart(distanceX: number, distanceY: number) {
            var axisX = this._chart.axisX, axisY = this._chart.axisY,
                sgs = this._getTransFormGroups();

            distanceX = (Math.abs(distanceX)) < this._threadHold ? 0 : distanceX;
            distanceY = (Math.abs(distanceY)) < this._threadHold ? 0 : distanceY;

            if (this._interactiveAxes === InteractiveAxes.X) {
                distanceY = 0;
            }

            if (this._interactiveAxes === InteractiveAxes.Y) {
                distanceX = 0;
            }
            
            // check x axis range
            if (distanceX > 0 && axisX.actualMin.valueOf() === this._minX) {
                distanceX = 0;
            }
            if (distanceX < 0 && axisX.actualMax.valueOf() === this._maxX) {
                distanceX = 0;
            }
            // check y axis range
            if (distanceY > 0 && axisY.actualMax.valueOf() === this._maxY) {
                distanceY = 0;
            }
            if (distanceY < 0 && axisY.actualMin.valueOf() === this._minY) {
                distanceY = 0;
            }

            for (var i = 0; i < sgs.length; i++) {
                sgs[i].setAttribute('transform', 'translate(' + distanceX + ',' + distanceY + ')');
            }
        }

        private _pannedChart(distanceX: number, distanceY: number) {

            if (this._interactiveAxes === InteractiveAxes.X ||
                this._interactiveAxes === InteractiveAxes.XY) {
                this._updateAxisByDistance(true, distanceX);
            }

            if (this._interactiveAxes === InteractiveAxes.Y ||
                this._interactiveAxes === InteractiveAxes.XY) {
                this._updateAxisByDistance(false, -distanceY);
            }
        }

        private _scalingChart(scale, offset) {
            var axisX = this._chart.axisX, axisY = this._chart.axisY,
                seriesGroups,
                offsetX = offset.x !== undefined ? offset.x : 0,
                offsetY = offset.y !== undefined ? offset.y : 0,
                scaleX, scaleY;

            if (!scale) {
                return;
            }
            seriesGroups = this._getTransFormGroups();
            //check x axis range
            if (scale.x !== undefined) {
                if (scale.x < 1) {
                    if (axisX.actualMin.valueOf() === this._minX &&
                        axisX.actualMax.valueOf() === this._maxX) {
                        scale.x = 1;
                        offsetX = 0;
                    }
                }
            }
            //check y axis range
            if (scale.y !== undefined) {
                if (scale.y < 1) {
                    if (axisY.actualMin.valueOf() === this._minY &&
                        axisY.actualMax.valueOf() === this._maxY) {
                        scale.y = 1;
                        offsetY = 0;
                    }
                }
            }

            scaleX = scale.x !== undefined ? scale.x : 1;
            scaleY = scale.y !== undefined ? scale.y : 1;

            for (var i = 0; i < seriesGroups.length; i++) {
                seriesGroups[i].setAttribute('transform', 'translate(' + offsetX + ', ' + offsetY + ') ' +
                    'scale(' + scaleX + ', ' + scaleY + ')');
            }
        }

        private _scaledChart(e: any) {
            var min, max,
                chart = this._chart,
                axisX = chart.axisX,
                axisY = chart.axisY;

            if (!this._clip) {
                return;
            }

            if (this._interactiveAxes !== InteractiveAxes.Y) {
                if (this._clip['x'] !== undefined) {
                    min = Math.max(this._minX, axisX.convertBack(this._clip['x']));
                    max = Math.min(this._maxX, axisX.convertBack(this._clip['x'] + this._selection['w']));
                    if (min - max !== 0) {
                        this._updateAxisRange(axisX, min, max);
                    }
                }
            }

            if (this._interactiveAxes !== InteractiveAxes.X) {
                if (this._clip['y'] !== undefined) {
                    max = Math.min(this._maxY, axisY.convertBack(this._clip['y']));
                    min = Math.max(this._minY, axisY.convertBack(this._clip['y'] + this._selection['h']));
                    if (min - max !== 0) {
                        this._updateAxisRange(axisY, min, max);
                    }
                }
            }
        }

        //help method
        private _updateAxisByDistance(isX: boolean, distance: number) {
            var axis = isX ? this._chart.axisX : this._chart.axisY,
                oriMin = isX ? this._minX : this._minY,
                oriMax = isX ? this._maxX : this._maxY,
                min = axis.actualMin.valueOf(),
                max = axis.actualMax.valueOf(),
                change;

            if (distance === 0) {
                return;
            }
            if ((distance > 0 && oriMin === min) || (
                distance < 0 && oriMax === max)) {
                this._innerUpdating = true;
                this._chart.invalidate();
                return;
            }
            change = distance / (isX ? this._plotBox.width : this._plotBox.height);
            this._updateAxisByChg(isX, -change, -change);
        }

        private _updateAxisByChg(isX: boolean, chgMin: number, chgMax: number) {
            var axis = isX ? this._chart.axisX : this._chart.axisY,
                oriMin = isX ? this._minX : this._minY,
                oriMax = isX ? this._maxX : this._maxY,
                min = axis.actualMin.valueOf(),
                max = axis.actualMax.valueOf(),
                range = max - min,
                plotRect = this._chart._plotRect,
                lt = isX ? plotRect.left : plotRect.top,
                wh = isX ? plotRect.width : plotRect.height,
                minRange = isX ? this._minXRange : this._minYRange,
                tMin, tMax;

            if (isNaN(chgMin) || isNaN(chgMax)) {
                return;
            }

            if (this._panning) {
                if (chgMin < 0) {
                    //tMin = min + chgMin * range;
                    tMin = isX ? axis.convertBack(lt + chgMin * wh) : axis.convertBack(lt + wh - chgMin * wh);
                    if (tMin < oriMin) {
                        tMin = oriMin;
                        //tMax = tMin + range;
                        tMax = isX ? axis.convertBack(axis.convert(tMin) + wh) : axis.convertBack(axis.convert(tMin) - wh);
                    } else {
                        //tMax = max + chgMax * range;
                        tMax = isX ? axis.convertBack(lt + wh + chgMax * wh) : axis.convertBack(lt - chgMax * wh);
                    }
                } else {
                    //tMax = max + chgMax * range;
                    tMax = isX ? axis.convertBack(lt + wh + chgMax * wh) : axis.convertBack(lt - chgMax * wh);
                    if (tMax > oriMax) {
                        tMax = oriMax;
                        //tMin = tMax - range;
                        tMin = isX ? axis.convertBack(axis.convert(tMax) - wh) : axis.convertBack(axis.convert(tMax) + wh);
                    } else {
                        //tMin = min + chgMin * range;
                        tMin = isX ? axis.convertBack(lt + chgMin * wh) : axis.convertBack(lt + wh - chgMin * wh);
                    }
                }

            } else if (this._scaling) {
                //scaling: control the range 
                //tMin = min + chgMin * range;
                //tMax = max + chgMax * range;
                tMin = isX ? axis.convertBack(lt + chgMin * wh) : axis.convertBack(lt + wh - chgMin * wh);
                tMax = isX ? axis.convertBack(lt + wh + chgMax * wh) : axis.convertBack(lt - chgMax * wh);
                if (tMin < oriMin) {
                    tMin = oriMin;
                }
                if (tMax > oriMax) {
                    tMax = oriMax;
                }
                if ((tMax - tMin) < minRange) {
                    tMin = tMax - minRange;
                }
            }
            
            this._updateAxisRange(axis, tMin, tMax);      
        }

        private _initAxisRangeWithPosAndScale(isX: boolean) {
            var range, initRange, initMin, initMax;
            if (isX) {
                range = this._maxX - this._minX;
                initRange = range * this._scaleX;
                initMin = this._minX + this._posX * (range - initRange);
                initMax = this._minX + initRange;
                this._innerUpdating = true;
                this._chart.axisX.min = initMin;
                this._chart.axisX.max = initMax;
                this._lastMinX = initMin;
                this._lastMaxX = initMax;
            } else {
                range = this._maxY - this._minY;
                initRange = range * this._scaleY;
                initMin = this._minY + this._posY * (range - initRange);
                initMax = this._minY + initRange;
                this._innerUpdating = true;
                this._chart.axisY.min = initMin;
                this._chart.axisY.max = initMax; 
                this._lastMinY = initMin;
                this._lastMaxY = initMax;
            }  
        }

        private _updateAxisRange(axis: Axis, tMin: number, tMax: number) {
            this._chart.beginUpdate();
            axis.min = tMin;
            axis.max = tMax;
            if (axis === this._chart.axisX) {
                this._lastMinX = tMin;
                this._lastMaxX = tMax;
            } else {
                this._lastMinY = tMin;
                this._lastMaxY = tMax;
            }
            this._innerUpdating = true;
            this._chart.endUpdate();   
        }

        private _reset() {
            this._scaling = false;
            this._panning = false;
            this._startDistance = 0;
            this._startFirstPt = null;
            this._pinchStartEvents = [];
            this._startPointers = [];
            this._mvPointers = [];
            this._endPoint = null;
            this._clip = {};
            this._selection = {};
        }

        private _getAxisMin(axis: Axis) {
            return isDate(axis.actualMin) ? axis.actualMin.valueOf() : axis.actualMin;
        }

        private _getAxisMax(axis: Axis) {
            return isDate(axis.actualMax) ? axis.actualMax.valueOf() : axis.actualMax;
        }

        private _getTransFormGroups() {
            var seriesGroups = this._seriesGroup.querySelectorAll('g[clip-path]');

            //for Line chart: line chart's group has no clip-path attribute
            if (seriesGroups.length === 0) {
                seriesGroups = this._seriesGroup.querySelectorAll('g');
            }
            return seriesGroups;
        }

        private _disabledOthersInteraction(disabled: boolean) {
            //disabled the line marker
            var hostEle = this._chart.hostElement;
            if (hostEle === null || hostEle === undefined) {
                return;
            }
            var lineMarks = hostEle.querySelectorAll('.wj-chart-linemarker-container');
            for (var i = 0; i < lineMarks.length; i++) {
                if (disabled) {
                    wijmo.addClass(<HTMLElement>lineMarks[i], ChartGestures._CSS_BLOCK_INTERACTION);
                } else {
                    wijmo.removeClass(<HTMLElement>lineMarks[i], ChartGestures._CSS_BLOCK_INTERACTION);
                }
            }
        }

        private _getPoint(e: any): Point {
            return e instanceof MouseEvent ?
                new wijmo.Point(e.pageX, e.pageY) :
                new wijmo.Point(e.changedTouches[0].pageX, e.changedTouches[0].pageY);
        }

        private _getTouchPair(event: any) {
            var touches = [];

            // array of touches is supplied
            if (isArray(event)) {
                touches[0] = event[0];
                touches[1] = event[1];
            }
            // an event
            else {
                if (event.type === 'touchend') {
                    if (event.touches.length === 1) {
                        touches[0] = event.touches[0];
                        touches[1] = event.changedTouches[0];
                    }
                    else if (event.touches.length === 0) {
                        touches[0] = event.changedTouches[0];
                        touches[1] = event.changedTouches[1];
                    }
                }
                else {
                    touches[0] = event.touches[0];
                    touches[1] = event.touches[1];
                }
            }

            return touches;
        }

        private _touchDistance(event: any) {
            var touches = this._getTouchPair(event),
                dx = 0, dy = 0;
            if (touches[0] && touches[0]['pageX'] !== undefined
                && touches[1] && touches[1]['pageX'] !== undefined) {
                dx = touches[0]['pageX'] - touches[1]['pageX'];
            }

            if (touches[0] && touches[0]['pageY'] !== undefined
                && touches[1] && touches[1]['pageY'] !== undefined) {
                dy = touches[0]['pageY'] - touches[1]['pageY'];
            }

            return { x: dx, y: dy };
        }
    }
}
