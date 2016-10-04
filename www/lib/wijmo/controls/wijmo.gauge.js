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
* Defines the @see:RadialGauge, @see:LinearGauge, and @see:BulletGraph
* controls.
*
* Unlike many gauge controls, Wijmo gauges concentrate on the data being
* displayed, with little extraneous color and markup elements. They were
* designed to be easy to use and to read, especially on small-screen devices.
*
* Wijmo gauges are composed of @see:Range objects. Every Wijmo gauge has
* at least two ranges: the "face" and the "pointer".
*
* <ul><li>
* The "face" represents the gauge background. The "min" and "max"
* properties of the face range correspond to the "min" and "max" properties
* of the gauge control, and limit the values that the gauge can display.
* </li><li>
* The "pointer" is the range that indicates the gauge's current value. The
* "max" property of the pointer range corresponds to the "value" property
* of the gauge.
* </li></ul>
*
* In addition to these two special ranges, gauges may have any number of
* additional ranges added to their "ranges" collection. These additional
* ranges can be used for two things:
*
* <ul><li>
* By default, the extra ranges appear as part of the gauge background.
* This way you can show 'zones' within the gauge, like 'good,' 'average,'
* and 'bad' for example.
* </li><li>
* If you set the gauge's "showRanges" property to false, the additional
* ranges are not shown. Instead, they are used to automatically set the
* color of the "pointer" based on the current value.
* </li></ul>
*/
var wijmo;
(function (wijmo) {
    var gauge;
    (function (gauge) {
        'use strict';
        /**
         * Specifies which values to display as text.
         */
        (function (ShowText) {
            /** Do not show any text in the gauge. */
            ShowText[ShowText["None"] = 0] = "None";
            /** Show the gauge's @see:Gauge.value as text. */
            ShowText[ShowText["Value"] = 1] = "Value";
            /** Show the gauge's @see:Gauge.min and @see:Gauge.max values as text. */
            ShowText[ShowText["MinMax"] = 2] = "MinMax";
            /** Show the gauge's @see:Gauge.value, @see:Gauge.min, and @see:Gauge.max as text. */
            ShowText[ShowText["All"] = 3] = "All";
        })(gauge.ShowText || (gauge.ShowText = {}));
        var ShowText = gauge.ShowText;
        /**
         * Base class for the Wijmo Gauge controls (abstract).
         */
        var Gauge = (function (_super) {
            __extends(Gauge, _super);
            /**
             * Initializes a new instance of the @see:Gauge class.
             *
             * @param element The DOM element that hosts the control, or a selector for the host element (e.g. '#theCtrl').
             * @param options The JavaScript object containing initialization data for the control.
             */
            function Gauge(element, options) {
                var _this = this;
                _super.call(this, element, null, true);
                // property storage
                this._ranges = new wijmo.collections.ObservableArray();
                this._rngElements = [];
                this._format = 'n0';
                this._showRanges = true;
                this._shadow = true;
                this._animated = true;
                this._readOnly = true;
                this._step = 1;
                this._showText = ShowText.None;
                this._showTicks = false;
                // protected
                this._thickness = 0.8;
                this._initialized = false;
                /**
                 * Occurs when the value shown on the gauge changes.
                 */
                this.valueChanged = new wijmo.Event();
                // scales a value to a percentage based on the gauge's min and max properties
                this._getPercent = function (value) {
                    var pct = (this.max > this.min) ? (value - this.min) / (this.max - this.min) : 0;
                    return Math.max(0, Math.min(1, pct));
                };
                Gauge._ctr++;
                // instantiate and apply template
                var tpl = this.getTemplate();
                this.applyTemplate('wj-control wj-gauge', tpl, {
                    _dSvg: 'dsvg',
                    _svg: 'svg',
                    _filter: 'filter',
                    _gFace: 'gface',
                    _gRanges: 'granges',
                    _gPointer: 'gpointer',
                    _gCover: 'gcover',
                    _pFace: 'pface',
                    _pPointer: 'ppointer',
                    _cValue: 'cvalue',
                    _tValue: 'value',
                    _tMin: 'min',
                    _tMax: 'max',
                    _pTicks: 'pticks'
                });
                // apply filter id to template
                this._filterID = 'wj-gauge-filter-' + Gauge._ctr.toString(36);
                this._filter.setAttribute('id', this._filterID);
                // initialize main ranges
                this.face = new gauge.Range();
                this.pointer = new gauge.Range();
                // invalidate control and re-create range elements when ranges change
                this._ranges.collectionChanged.addHandler(function () {
                    // check types
                    var arr = _this._ranges;
                    for (var i = 0; i < arr.length; i++) {
                        var rng = wijmo.tryCast(arr[i], gauge.Range);
                        if (!rng) {
                            throw 'ranges array must contain Range objects.';
                        }
                    }
                    // remember ranges are dirty and invalidate
                    _this._rangesDirty = true;
                    _this.invalidate();
                });
                // keyboard handling
                this.addEventListener(this.hostElement, 'keydown', this._keydown.bind(this));
                // mouse handling
                this.addEventListener(this.hostElement, 'click', function (e) {
                    if (e.button == 0) {
                        _this.focus();
                        _this._applyMouseValue(e);
                    }
                });
                this.addEventListener(this.hostElement, 'mousedown', function (e) {
                    if (e.button == 0) {
                        _this.focus();
                        _this._applyMouseValue(e);
                    }
                });
                this.addEventListener(this.hostElement, 'mousemove', function (e) {
                    if (e.buttons == 1) {
                        _this._applyMouseValue(e, true);
                    }
                });
                // touch handling
                if ('ontouchstart' in window) {
                    this.addEventListener(this.hostElement, 'touchstart', function (e) {
                        _this.focus();
                        if (!e.defaultPrevented && !_this.isReadOnly && _this._applyMouseValue(e, true)) {
                            e.preventDefault();
                        }
                    });
                    this.addEventListener(this.hostElement, 'touchmove', function (e) {
                        if (!e.defaultPrevented && !_this.isReadOnly && _this._applyMouseValue(e, true)) {
                            e.preventDefault();
                        }
                    });
                }
                // use wheel to increase/decrease the value
                this.addEventListener(this.hostElement, 'wheel', function (e) {
                    if (!e.defaultPrevented && !_this.isReadOnly && _this.containsFocus() && _this.value != null && _this.hitTest(e)) {
                        var step = wijmo.clamp(-e.deltaY, -1, +1);
                        _this.value = wijmo.clamp(_this.value + (_this.step || 1) * step, _this.min, _this.max);
                        e.preventDefault();
                    }
                });
                // initialize control options
                this.initialize(options);
                // ensure face and text are updated
                this.invalidate();
            }
            Object.defineProperty(Gauge.prototype, "value", {
                /**
                 * Gets or sets the value to display on the gauge.
                 */
                get: function () {
                    return this._pointer.max;
                },
                set: function (value) {
                    if (value != this._pointer.max) {
                        this._pointer.max = wijmo.asNumber(value, true);
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "min", {
                /**
                 * Gets or sets the minimum value that can be displayed on the gauge.
                 *
                 * For details about using the @see:min and @see:max properties, please see the
                 * <a href="static/minMax.html">Using the min and max properties</a> topic.
                 */
                get: function () {
                    return this._face.min;
                },
                set: function (value) {
                    this._face.min = wijmo.asNumber(value);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "max", {
                /**
                 * Gets or sets the maximum value that can be displayed on the gauge.
                 *
                 * For details about using the @see:min and @see:max properties, please see the
                 * <a href="static/minMax.html">Using the min and max properties</a> topic.
                 */
                get: function () {
                    return this._face.max;
                },
                set: function (value) {
                    this._face.max = wijmo.asNumber(value);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "origin", {
                /**
                 * Gets or sets the starting point used for painting the range.
                 *
                 * By default, this property is set to null, which causes the value range
                 * to start at the gauge's minimum value, or zero if the minimum is less
                 * than zero.
                 */
                get: function () {
                    return this._origin;
                },
                set: function (value) {
                    if (value != this._origin) {
                        this._origin = wijmo.asNumber(value, true);
                        this.invalidate();
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "isReadOnly", {
                /**
                 * Gets or sets a value that indicates whether the user can edit the value
                 * using the mouse and keyboard.
                 */
                get: function () {
                    return this._readOnly;
                },
                set: function (value) {
                    this._readOnly = wijmo.asBoolean(value);
                    this._setAttribute(this._svg, 'cursor', this._readOnly ? null : 'pointer');
                    wijmo.toggleClass(this.hostElement, 'wj-state-readonly', this.isReadOnly);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "step", {
                /**
                 * Gets or sets the amount to add to or subtract from the @see:value property
                 * when the user presses the arrow keys or moves the mouse wheel.
                 */
                get: function () {
                    return this._step;
                },
                set: function (value) {
                    this._step = wijmo.asNumber(value, true);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "format", {
                /**
                 * Gets or sets the format string used to display gauge values as text.
                 */
                get: function () {
                    return this._format;
                },
                set: function (value) {
                    if (value != this._format) {
                        this._format = wijmo.asString(value);
                        this.invalidate();
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "getText", {
                /**
                 * Gets or sets a callback that returns customized strings used to
                 * display gauge values.
                 *
                 * Use this property if you want to customize the strings shown on
                 * the gauge in cases where the @see:format property is not enough.
                 *
                 * If provided, the callback should be a function as that takes as
                 * parameters the gauge, the part name, the value, and the formatted
                 * value. The callback should return the string to be displayed on
                 * the gauge.
                 *
                 * For example:
                 *
                 * <pre>// callback to convert values into strings
                 * gauge.getText = function (gauge, part, value, text) {
                 *   switch (part) {
                 *     case 'value':
                 *       if (value &lt;= 10) return 'Empty!';
                 *       if (value &lt;= 25) return 'Low...';
                 *       if (value &lt;= 95) return 'Good';
                 *       return 'Full';
                 *     case 'min':
                 *       return 'EMPTY';
                 *     case 'max':
                 *       return 'FULL';
                 *   }
                 *   return text;
                 * }</pre>
                 */
                get: function () {
                    return this._getText;
                },
                set: function (value) {
                    if (value != this._getText) {
                        this._getText = wijmo.asFunction(value);
                        this.invalidate();
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "thickness", {
                /**
                 * Gets or sets the thickness of the gauge, on a scale between zero and one.
                 *
                 * Setting the thickness to one causes the gauge to fill as much of the
                 * control area as possible. Smaller values create thinner gauges.
                 */
                get: function () {
                    return this._thickness;
                },
                set: function (value) {
                    if (value != this._thickness) {
                        this._thickness = wijmo.clamp(wijmo.asNumber(value, false), 0, 1);
                        this.invalidate();
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "face", {
                /**
                 * Gets or sets the @see:Range used to represent the gauge's overall geometry
                 * and appearance.
                 */
                get: function () {
                    return this._face;
                },
                set: function (value) {
                    if (value != this._face) {
                        if (this._face) {
                            this._face.propertyChanged.removeHandler(this._rangeChanged);
                        }
                        this._face = wijmo.asType(value, gauge.Range);
                        if (this._face) {
                            this._face.propertyChanged.addHandler(this._rangeChanged, this);
                        }
                        this.invalidate();
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "pointer", {
                /**
                 * Gets or sets the @see:Range used to represent the gauge's current value.
                 */
                get: function () {
                    return this._pointer;
                },
                set: function (value) {
                    if (value != this._pointer) {
                        var gaugeValue = null;
                        if (this._pointer) {
                            gaugeValue = this.value;
                            this._pointer.propertyChanged.removeHandler(this._rangeChanged);
                        }
                        this._pointer = wijmo.asType(value, gauge.Range);
                        if (this._pointer) {
                            if (gaugeValue) {
                                this.value = gaugeValue;
                            }
                            this._pointer.propertyChanged.addHandler(this._rangeChanged, this);
                        }
                        this.invalidate();
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "showText", {
                /**
                 * Gets or sets the @see:ShowText values to display as text in the gauge.
                 */
                get: function () {
                    return this._showText;
                },
                set: function (value) {
                    if (value != this._showText) {
                        this._showText = wijmo.asEnum(value, ShowText);
                        this.invalidate();
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "showTicks", {
                /**
                 * Gets or sets a property that determines whether the gauge should display
                 * tickmarks at each @see:step value.
                 *
                 * The tickmarks can be formatted in CSS using the <b>wj-gauge</b> and
                 * <b>wj-ticks</b> class names. For example:
                 *
                 * <pre>.wj-gauge .wj-ticks {
                 *     stroke-width: 2px;
                 *     stroke: white;
                 * }</pre>
                 */
                get: function () {
                    return this._showTicks;
                },
                set: function (value) {
                    if (value != this._showTicks) {
                        this._showTicks = wijmo.asBoolean(value);
                        this.invalidate();
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "thumbSize", {
                /**
                 * Gets or sets the size of the element that shows the gauge's current value, in pixels.
                 */
                get: function () {
                    return this._thumbSize;
                },
                set: function (value) {
                    if (value != this._thumbSize) {
                        this._thumbSize = wijmo.asNumber(value, true, true);
                        this.invalidate();
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "showRanges", {
                /**
                 * Gets or sets a value that indicates whether the gauge displays the ranges contained in
                 * the @see:ranges property.
                 *
                 * If this property is set to false, the ranges contained in the @see:ranges property are not
                 * displayed in the gauge. Instead, they are used to interpolate the color of the @see:pointer
                 * range while animating value changes.
                 */
                get: function () {
                    return this._showRanges;
                },
                set: function (value) {
                    if (value != this._showRanges) {
                        this._showRanges = wijmo.asBoolean(value);
                        this._animColor = null;
                        this._rangesDirty = true;
                        this.invalidate();
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "hasShadow", {
                /**
                 * Gets or sets a value that indicates whether the gauge displays a shadow effect.
                 */
                get: function () {
                    return this._shadow;
                },
                set: function (value) {
                    if (value != this._shadow) {
                        this._shadow = wijmo.asBoolean(value);
                        this.invalidate();
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "isAnimated", {
                /**
                 * Gets or sets a value that indicates whether the gauge animates value changes.
                 */
                get: function () {
                    return this._animated;
                },
                set: function (value) {
                    if (value != this._animated) {
                        this._animated = wijmo.asBoolean(value);
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Gauge.prototype, "ranges", {
                /**
                 * Gets the collection of ranges in this gauge.
                 */
                get: function () {
                    return this._ranges;
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Raises the @see:valueChanged event.
             */
            Gauge.prototype.onValueChanged = function (e) {
                this.valueChanged.raise(this, e);
            };
            /**
             * Refreshes the control.
             *
             * @param fullUpdate Indicates whether to update the control layout as well as the content.
             */
            Gauge.prototype.refresh = function (fullUpdate) {
                if (fullUpdate === void 0) { fullUpdate = true; }
                _super.prototype.refresh.call(this, fullUpdate);
                // update ranges if they are dirty
                if (this._rangesDirty) {
                    this._rangesDirty = false;
                    var gr = this._gRanges;
                    // remove old elements and disconnect event handlers
                    for (var i = 0; i < this._rngElements.length; i++) {
                        var e = this._rngElements[i];
                        e.rng.propertyChanged.removeHandler(this._rangeChanged);
                    }
                    while (gr.lastChild) {
                        gr.removeChild(gr.lastChild);
                    }
                    this._rngElements = [];
                    // add elements for each range and listen to changes
                    if (this._showRanges) {
                        for (var i = 0; i < this.ranges.length; i++) {
                            var rng = this.ranges[i];
                            rng.propertyChanged.addHandler(this._rangeChanged, this);
                            this._rngElements.push({
                                rng: rng,
                                el: this._createElement('path', gr)
                            });
                        }
                    }
                }
                // update text elements
                this._showElement(this._tValue, (this.showText & ShowText.Value) != 0);
                this._showElement(this._tMin, (this.showText & ShowText.MinMax) != 0);
                this._showElement(this._tMax, (this.showText & ShowText.MinMax) != 0);
                this._showElement(this._cValue, (this.showText & ShowText.Value) != 0 || this._thumbSize > 0);
                this._updateText();
                // update face and pointer
                var filterUrl = this._getFilterUrl();
                this._setAttribute(this._pFace, 'filter', filterUrl);
                this._setAttribute(this._pPointer, 'filter', filterUrl);
                this._updateRange(this._face);
                this._updateRange(this._pointer);
                this._updateTicks();
                // update ranges
                for (var i = 0; i < this.ranges.length; i++) {
                    this._updateRange(this.ranges[i]);
                }
                // ready
                this._initialized = true;
            };
            /**
             * Gets a number that corresponds to the value of the gauge at a given point.
             *
             * For example:
             *
             * <pre>
             * // hit test a point when the user clicks on the gauge
             * gauge.hostElement.addEventListener('click', function (e) {
             *   var ht = gauge.hitTest(e.pageX, e.pageY);
             *   if (ht != null) {
             *     console.log('you clicked the gauge at value ' + ht.toString());
             *   }
             * });
             * </pre>
             *
             * @param pt The point to investigate, in window coordinates, or a MouseEvent object,
             * or the x coordinate of the point.
             * @param y The Y coordinate of the point (if the first parameter is a number).
             * @return Value of the gauge at the point, or null if the point is not on the gauge's face.
             */
            Gauge.prototype.hitTest = function (pt, y) {
                // get point in page coordinates
                if (wijmo.isNumber(pt) && wijmo.isNumber(y)) {
                    pt = new wijmo.Point(pt, y);
                }
                else if (!(pt instanceof wijmo.Point)) {
                    pt = wijmo.mouseToPage(pt);
                }
                pt = wijmo.asType(pt, wijmo.Point);
                // convert point to gauge client coordinates
                var rc = wijmo.Rect.fromBoundingRect(this._dSvg.getBoundingClientRect());
                pt.x -= rc.left + pageXOffset;
                pt.y -= rc.top + pageYOffset;
                // get gauge value from point
                return this._getValueFromPoint(pt);
            };
            // ** implementation
            // safe version of getBBox (TFS 129851, 144174)
            // (workaround for FF bug, see https://bugzilla.mozilla.org/show_bug.cgi?id=612118)
            Gauge._getBBox = function (e) {
                try {
                    return e.getBBox();
                }
                catch (x) {
                    return { x: 0, y: 0, width: 0, height: 0 };
                }
            };
            // gets the unique filter ID used by this gauge
            Gauge.prototype._getFilterUrl = function () {
                return this.hasShadow ? 'url(#' + this._filterID + ')' : null;
            };
            // gets the path element that represents a Range
            Gauge.prototype._getRangeElement = function (rng) {
                if (rng == this._face) {
                    return this._pFace;
                }
                else if (rng == this._pointer) {
                    return this._pPointer;
                }
                for (var i = 0; i < this._rngElements.length; i++) {
                    var rngEl = this._rngElements[i];
                    if (rngEl.rng == rng) {
                        return rngEl.el;
                    }
                }
                return null;
            };
            // handle changes to range objects
            Gauge.prototype._rangeChanged = function (rng, e) {
                var _this = this;
                // when pointer.max changes, raise valueChanged
                if (rng == this._pointer && e.propertyName == 'max') {
                    this.onValueChanged();
                    this._updateText();
                }
                // when face changes, invalidate the whole gauge
                if (rng == this._face) {
                    this.invalidate();
                    return;
                }
                // update pointer with animation
                if (rng == this._pointer && e.propertyName == 'max') {
                    // clear pending animations if any
                    if (this._animInterval) {
                        clearInterval(this._animInterval);
                    }
                    // animate
                    if (this.isAnimated && !this.isUpdating && this._initialized) {
                        var s1 = this._getPointerColor(e.oldValue), s2 = this._getPointerColor(e.newValue), c1 = s1 ? new wijmo.Color(s1) : null, c2 = s2 ? new wijmo.Color(s2) : null;
                        this._animInterval = wijmo.animate(function (pct) {
                            _this._animColor = (c1 && c2)
                                ? wijmo.Color.interpolate(c1, c2, pct).toString()
                                : null;
                            _this._updateRange(rng, e.oldValue + pct * (e.newValue - e.oldValue));
                            if (pct >= 1) {
                                _this._animColor = null;
                                _this._animInterval = null;
                                _this._updateRange(rng);
                                _this._updateText();
                            }
                        });
                        return;
                    }
                }
                // update range without animation
                this._updateRange(rng);
            };
            // creates an SVG element with the given tag and appends it to a given element
            Gauge.prototype._createElement = function (tag, parent, cls) {
                var e = document.createElementNS(Gauge._SVGNS, tag);
                if (cls) {
                    e.setAttribute('class', cls);
                }
                parent.appendChild(e);
                return e;
            };
            // centers an SVG text element at a given point
            Gauge.prototype._centerText = function (e, value, center) {
                if (e.getAttribute('display') != 'none') {
                    // get the text for the element
                    var text = wijmo.Globalize.format(value, this.format);
                    if (wijmo.isFunction(this.getText)) {
                        var part = e == this._tValue ? 'value' :
                            e == this._tMin ? 'min' :
                                e == this._tMax ? 'max' :
                                    null;
                        wijmo.assert(part != null, 'unknown element');
                        text = this.getText(this, part, value, text);
                    }
                    // set the text and center the element
                    e.textContent = text;
                    var box = wijmo.Rect.fromBoundingRect(Gauge._getBBox(e)), x = (center.x - box.width / 2), y = (center.y + box.height / 4);
                    e.setAttribute('x', this._fix(x));
                    e.setAttribute('y', this._fix(y));
                }
            };
            // method used in JSON-style initialization
            Gauge.prototype._copy = function (key, value) {
                if (key == 'ranges') {
                    var arr = wijmo.asArray(value);
                    for (var i = 0; i < arr.length; i++) {
                        var r = new gauge.Range();
                        wijmo.copy(r, arr[i]);
                        this.ranges.push(r);
                    }
                    return true;
                }
                else if (key == 'pointer') {
                    wijmo.copy(this.pointer, value);
                    return true;
                }
                return false;
            };
            // shows or hides an element
            Gauge.prototype._showElement = function (e, show) {
                this._setAttribute(e, 'display', show ? '' : 'none');
            };
            // sets or clears an attribute
            Gauge.prototype._setAttribute = function (e, att, value) {
                if (value) {
                    e.setAttribute(att, value);
                }
                else {
                    e.removeAttribute(att);
                }
            };
            // updates the element for a given range
            Gauge.prototype._updateRange = function (rng, value) {
                if (value === void 0) { value = rng.max; }
                // update pointer's min value
                if (rng == this._pointer) {
                    rng.min = this.origin != null
                        ? this.origin
                        : (this.min < 0 && this.max > 0) ? 0 : this.min;
                }
                // update the range's element
                var e = this._getRangeElement(rng);
                if (e) {
                    this._updateRangeElement(e, rng, value);
                    var color = rng.color;
                    if (rng == this._pointer) {
                        color = this._animColor ? this._animColor : this._getPointerColor(rng.max);
                    }
                    this._setAttribute(e, 'style', color ? 'fill:' + color : null);
                }
            };
            // gets the color for the pointer range based on the gauge ranges
            Gauge.prototype._getPointerColor = function (value) {
                var rng;
                if (!this._showRanges) {
                    for (var i = 0; i < this._ranges.length; i++) {
                        var r = this._ranges[i];
                        if (value >= r.min && value <= r.max) {
                            if (rng == null || rng.max - rng.min > r.max - r.min) {
                                rng = r;
                            }
                        }
                    }
                    if (rng) {
                        return rng.color;
                    }
                }
                return this._pointer.color;
            };
            // keyboard handling
            Gauge.prototype._keydown = function (e) {
                if (!this._readOnly && this._step) {
                    var handled = true;
                    switch (e.keyCode) {
                        case wijmo.Key.Left:
                        case wijmo.Key.Down:
                            this.value = wijmo.clamp(this.value - this.step, this.min, this.max);
                            break;
                        case wijmo.Key.Right:
                        case wijmo.Key.Up:
                            this.value = wijmo.clamp(this.value + this.step, this.min, this.max);
                            break;
                        case wijmo.Key.Home:
                            this.value = this.min;
                            break;
                        case wijmo.Key.End:
                            this.value = this.max;
                            break;
                        default:
                            handled = false;
                            break;
                    }
                    if (handled) {
                        e.preventDefault();
                    }
                }
            };
            // apply value based on mouse/pointer position
            Gauge.prototype._applyMouseValue = function (e, instant) {
                if (!this.isReadOnly && this.containsFocus()) {
                    var value = this.hitTest(e);
                    if (value != null) {
                        // disable animation for instant changes
                        var a = this._animated;
                        if (instant) {
                            this._animated = false;
                        }
                        // make the change
                        if (this._step != null) {
                            value = Math.round(value / this._step) * this._step;
                        }
                        this.value = wijmo.clamp(value, this.min, this.max);
                        // restore animation and return true
                        this._animated = a;
                        return true;
                    }
                }
                // not editable or hit-test off the gauge? return false
                return false;
            };
            // ** virtual methods (must be overridden in derived classes)
            // updates the range element
            Gauge.prototype._updateRangeElement = function (e, rng, value) {
                wijmo.assert(false, 'Gauge is an abstract class.');
            };
            // updates the text elements
            Gauge.prototype._updateText = function () {
                wijmo.assert(false, 'Gauge is an abstract class.');
            };
            // updates the tickmarks
            Gauge.prototype._updateTicks = function () {
                wijmo.assert(false, 'Gauge is an abstract class.');
            };
            // gets the value at a given point (in gauge client coordinates)
            Gauge.prototype._getValueFromPoint = function (pt) {
                return null;
            };
            // formats numbers or points with up to 4 decimals
            Gauge.prototype._fix = function (n) {
                return wijmo.isNumber(n)
                    ? parseFloat(n.toFixed(4)).toString()
                    : this._fix(n.x) + ' ' + this._fix(n.y);
            };
            Gauge._SVGNS = 'http://www.w3.org/2000/svg';
            Gauge._ctr = 0;
            /**
             * Gets or sets the template used to instantiate @see:Gauge controls.
             */
            Gauge.controlTemplate = '<div wj-part="dsvg" style="width:100%;height:100%">' +
                '<svg wj-part="svg" width="100%" height="100%" style="overflow:visible">' +
                '<defs>' +
                '<filter wj-part="filter">' +
                '<feOffset dx="3" dy="3"></feOffset>' +
                '<feGaussianBlur result="offset-blur" stdDeviation="5"></feGaussianBlur>' +
                '<feComposite operator="out" in="SourceGraphic" in2="offset-blur" result="inverse"></feComposite>' +
                '<feFlood flood-color="black" flood-opacity="0.2" result="color"></feFlood>' +
                '<feComposite operator="in" in="color" in2="inverse" result="shadow"></feComposite>' +
                '<feComposite operator="over" in="shadow" in2="SourceGraphic"></feComposite>' +
                '</filter>' +
                '</defs>' +
                '<g wj-part="gface" class="wj-face" style="cursor:inherit">' +
                '<path wj-part="pface"/>' +
                '</g>' +
                '<g wj-part="granges" style="cursor:inherit"/>' +
                '<g wj-part="gpointer" class="wj-pointer" style="cursor:inherit">' +
                '<path wj-part="ppointer"/>' +
                '</g>' +
                '<g wj-part="gcover" style="cursor:inherit">' +
                '<path wj-part="pticks" class="wj-ticks"/>' +
                '<circle wj-part="cvalue" class="wj-pointer wj-thumb"/>' +
                '<text wj-part="value" class="wj-value"/>' +
                '<text wj-part="min" class="wj-min"/>' +
                '<text wj-part="max" class="wj-max"/>' +
                '</g>' +
                '</svg>' +
                '</div>';
            return Gauge;
        }(wijmo.Control));
        gauge.Gauge = Gauge;
    })(gauge = wijmo.gauge || (wijmo.gauge = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Gauge.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var gauge;
    (function (gauge) {
        'use strict';
        /**
         * Represents the direction in which the pointer of a @see:LinearGauge
         * increases.
         */
        (function (GaugeDirection) {
            /** Gauge value increases from left to right. */
            GaugeDirection[GaugeDirection["Right"] = 0] = "Right";
            /** Gauge value increases from right to left. */
            GaugeDirection[GaugeDirection["Left"] = 1] = "Left";
            /** Gauge value increases from bottom to top. */
            GaugeDirection[GaugeDirection["Up"] = 2] = "Up";
            /** Gauge value increases from top to bottom. */
            GaugeDirection[GaugeDirection["Down"] = 3] = "Down";
        })(gauge.GaugeDirection || (gauge.GaugeDirection = {}));
        var GaugeDirection = gauge.GaugeDirection;
        /**
         * The @see:LinearGauge displays a linear scale with an indicator
         * that represents a single value and optional ranges to represent
         * reference values.
         *
         * If you set the gauge's @see:LinearGauge.isReadOnly property to
         * false, then users will be able to edit the value by clicking on
         * the gauge.
         *
         * @fiddle:t842jozb
         */
        var LinearGauge = (function (_super) {
            __extends(LinearGauge, _super);
            /**
             * Initializes a new instance of the @see:LinearGauge class.
             *
             * @param element The DOM element that will host the control, or a selector for the host element (e.g. '#theCtrl').
             * @param options JavaScript object containing initialization data for the control.
             */
            function LinearGauge(element, options) {
                _super.call(this, element, null);
                // property storage
                this._direction = GaugeDirection.Right;
                // customize
                wijmo.addClass(this.hostElement, 'wj-lineargauge');
                // initialize control options
                this.initialize(options);
            }
            Object.defineProperty(LinearGauge.prototype, "direction", {
                /**
                 * Gets or sets the direction in which the gauge is filled.
                 */
                get: function () {
                    return this._direction;
                },
                set: function (value) {
                    if (value != this._direction) {
                        this._direction = wijmo.asEnum(value, GaugeDirection);
                        this.invalidate();
                    }
                },
                enumerable: true,
                configurable: true
            });
            // virtual methods
            // updates the element for a given range
            LinearGauge.prototype._updateRangeElement = function (e, rng, value) {
                // update the path
                var rc = this._getRangeRect(rng, value);
                this._updateSegment(e, rc);
                // check whether we have to show text and/or thumb
                var showText = (rng == this._pointer) && (this.showText & gauge.ShowText.Value) != 0, showThumb = showText || (rng == this._pointer && this.thumbSize > 0);
                // calculate thumb center
                if (showText || showThumb) {
                    var x = rc.left + rc.width / 2, y = rc.top + rc.height / 2;
                    switch (this.direction) {
                        case GaugeDirection.Right:
                            x = rc.right;
                            break;
                        case GaugeDirection.Left:
                            x = rc.left;
                            break;
                        case GaugeDirection.Up:
                            y = rc.top;
                            break;
                        case GaugeDirection.Down:
                            y = rc.bottom;
                            break;
                    }
                }
                // update text
                if (showText) {
                    this._centerText(this._tValue, value, new wijmo.Point(x, y));
                }
                // update thumb
                if (showText || showThumb) {
                    rc = wijmo.Rect.fromBoundingRect(gauge.Gauge._getBBox(this._tValue));
                    var color = this._animColor ? this._animColor : this._getPointerColor(rng.max), radius = this.thumbSize != null ? this.thumbSize / 2 : Math.max(rc.width, rc.height) * .8, ce = this._cValue;
                    this._setAttribute(ce, 'cx', this._fix(x));
                    this._setAttribute(ce, 'cy', this._fix(y));
                    this._setAttribute(ce, 'style', color ? 'fill:' + color : null);
                    this._setAttribute(ce, 'r', this._fix(radius));
                }
            };
            // update the text elements
            LinearGauge.prototype._updateText = function () {
                var rc = this._getRangeRect(this._face);
                switch (this.direction) {
                    case GaugeDirection.Right:
                        this._setText(this._tMin, this.min, rc, 'left');
                        this._setText(this._tMax, this.max, rc, 'right');
                        break;
                    case GaugeDirection.Left:
                        this._setText(this._tMin, this.min, rc, 'right');
                        this._setText(this._tMax, this.max, rc, 'left');
                        break;
                    case GaugeDirection.Up:
                        this._setText(this._tMin, this.min, rc, 'bottom');
                        this._setText(this._tMax, this.max, rc, 'top');
                        break;
                    case GaugeDirection.Down:
                        this._setText(this._tMin, this.min, rc, 'top');
                        this._setText(this._tMax, this.max, rc, 'bottom');
                        break;
                }
            };
            // update the tickmarks
            LinearGauge.prototype._updateTicks = function () {
                var d = '';
                if (this.showTicks && this.step) {
                    var rc = this._getRangeRect(this._face);
                    for (var t = this.min + this.step; t < this.max; t += this.step) {
                        switch (this.direction) {
                            case GaugeDirection.Right:
                                var tx = this._fix(rc.left + rc.width * this._getPercent(t));
                                d += 'M ' + tx + ' ' + this._fix(rc.top) + ' L ' + tx + ' ' + this._fix(rc.bottom) + ' ';
                                break;
                            case GaugeDirection.Left:
                                var tx = this._fix(rc.right - rc.width * this._getPercent(t));
                                d += 'M ' + tx + ' ' + rc.top.toFixed(2) + ' L ' + tx + ' ' + rc.bottom.toFixed(2) + ' ';
                                break;
                            case GaugeDirection.Up:
                                var ty = (rc.bottom - rc.height * this._getPercent(t)).toFixed(2);
                                d += 'M ' + this._fix(rc.left) + ' ' + ty + ' L ' + this._fix(rc.right) + ' ' + ty + ' ';
                                break;
                            case GaugeDirection.Down:
                                var ty = (rc.top + rc.height * this._getPercent(t)).toFixed(2);
                                d += 'M ' + rc.left.toFixed(2) + ' ' + ty + ' L ' + rc.right.toFixed(2) + ' ' + ty + ' ';
                                break;
                        }
                    }
                }
                this._pTicks.setAttribute('d', d);
            };
            // ** private stuff
            // draws a rectangular segment at the specified position
            LinearGauge.prototype._updateSegment = function (path, rc) {
                var data = {
                    p1: this._fix(new wijmo.Point(rc.left, rc.top)),
                    p2: this._fix(new wijmo.Point(rc.right, rc.top)),
                    p3: this._fix(new wijmo.Point(rc.right, rc.bottom)),
                    p4: this._fix(new wijmo.Point(rc.left, rc.bottom))
                };
                var content = wijmo.format('M {p1} L {p2} L {p3} L {p4} Z', data);
                path.setAttribute('d', content);
            };
            // positions a text element
            LinearGauge.prototype._setText = function (e, value, rc, pos) {
                if (e.getAttribute('display') != 'none') {
                    e.textContent = wijmo.Globalize.format(value, this.format);
                    var box = wijmo.Rect.fromBoundingRect(gauge.Gauge._getBBox(e)), pt = new wijmo.Point(rc.left + rc.width / 2 - box.width / 2, rc.top + rc.height / 2 + box.height / 2);
                    switch (pos) {
                        case 'top':
                            pt.y = rc.top - 4;
                            break;
                        case 'left':
                            pt.x = rc.left - 4 - box.width;
                            break;
                        case 'right':
                            pt.x = rc.right + 4;
                            break;
                        case 'bottom':
                            pt.y = rc.bottom + 4 + box.height;
                            break;
                    }
                    e.setAttribute('x', this._fix(pt.x));
                    e.setAttribute('y', this._fix(pt.y));
                }
            };
            // gets a rectangle that represents a Range
            LinearGauge.prototype._getRangeRect = function (rng, value) {
                if (value === void 0) { value = rng.max; }
                // get gauge size
                var rc = new wijmo.Rect(0, 0, this.hostElement.clientWidth, this.hostElement.clientHeight);
                // get face rect (account for thickness, text or thumb at edges)
                var padding = this.thumbSize ? Math.ceil(this.thumbSize / 2) : 0;
                if (this.showText != gauge.ShowText.None) {
                    padding = Math.max(padding, 3 * parseInt(getComputedStyle(this.hostElement).fontSize));
                }
                switch (this.direction) {
                    case GaugeDirection.Right:
                    case GaugeDirection.Left:
                        rc = rc.inflate(-padding, -rc.height * (1 - this.thickness * rng.thickness) / 2);
                        break;
                    case GaugeDirection.Up:
                    case GaugeDirection.Down:
                        rc = rc.inflate(-rc.width * (1 - this.thickness * rng.thickness) / 2, -padding);
                        break;
                }
                // get range rect
                var face = rng == this._face, pctMin = face ? 0 : this._getPercent(rng.min), pctMax = face ? 1 : this._getPercent(rng.max);
                switch (this.direction) {
                    case GaugeDirection.Right:
                        rc.left += rc.width * pctMin;
                        rc.width *= (pctMax - pctMin);
                        break;
                    case GaugeDirection.Left:
                        rc.left = rc.right - rc.width * pctMax;
                        rc.width = rc.width * (pctMax - pctMin);
                        break;
                    case GaugeDirection.Down:
                        rc.top += rc.height * pctMin;
                        rc.height *= (pctMax - pctMin);
                        break;
                    case GaugeDirection.Up:
                        rc.top = rc.bottom - rc.height * pctMax;
                        rc.height = rc.height * (pctMax - pctMin);
                        break;
                }
                // done
                return rc;
            };
            // gets the gauge value at a given point (in gauge client coordinates)
            LinearGauge.prototype._getValueFromPoint = function (pt) {
                // get face rectangle to calculate coordinates
                var rc = this._getRangeRect(this._face);
                // accept clicks anywhere to be touch-friendly
                //if (!rc.contains(pt)) {
                //    return null;
                //}
                // get position in control coordinates (min to max)
                var pct = 0;
                switch (this.direction) {
                    case GaugeDirection.Right:
                        pct = rc.width > 0 ? (pt.x - rc.left) / rc.width : 0;
                        break;
                    case GaugeDirection.Left:
                        pct = rc.width > 0 ? (rc.right - pt.x) / rc.width : 0;
                        break;
                    case GaugeDirection.Up:
                        pct = rc.height > 0 ? (rc.bottom - pt.y) / rc.height : 0;
                        break;
                    case GaugeDirection.Down:
                        pct = rc.height > 0 ? (pt.y - rc.top) / rc.height : 0;
                        break;
                }
                // done
                return this.min + pct * (this.max - this.min);
            };
            return LinearGauge;
        }(gauge.Gauge));
        gauge.LinearGauge = LinearGauge;
    })(gauge = wijmo.gauge || (wijmo.gauge = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=LinearGauge.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var gauge;
    (function (gauge) {
        'use strict';
        /**
         * The @see:RadialGauge displays a circular scale with an indicator
         * that represents a single value and optional ranges to represent
         * reference values.
         *
         * If you set the gauge's @see:RadialGauge.isReadOnly property to
         * false, then users will be able to edit the value by clicking on
         * the gauge.
         *
         * @fiddle:7ec2144u
         */
        var RadialGauge = (function (_super) {
            __extends(RadialGauge, _super);
            /**
             * Initializes a new instance of the @see:RadialGauge class.
             *
             * @param element The DOM element that hosts the control, or a selector for the host element (e.g. '#theCtrl').
             * @param options The JavaScript object containing initialization data for the control.
             */
            function RadialGauge(element, options) {
                _super.call(this, element, null);
                // property storage
                this._startAngle = 0;
                this._sweepAngle = 180;
                this._autoScale = true;
                // customize
                wijmo.addClass(this.hostElement, 'wj-radialgauge');
                this._thickness = .4;
                this.showText = gauge.ShowText.All;
                // initialize control options
                this.initialize(options);
            }
            Object.defineProperty(RadialGauge.prototype, "startAngle", {
                /**
                 * Gets or sets the starting angle for the gauge, in degrees.
                 *
                 * Angles are measured in degrees, clockwise, starting from the 9 o'clock position.
                 */
                get: function () {
                    return this._startAngle;
                },
                set: function (value) {
                    if (value != this._startAngle) {
                        this._startAngle = wijmo.clamp(wijmo.asNumber(value, false), -360, 360);
                        this.invalidate();
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(RadialGauge.prototype, "sweepAngle", {
                /**
                 * Gets or sets the sweeping angle for the gauge, in degrees.
                 *
                 * Angles are measured in degrees, clockwise, starting from the 9 o'clock position.
                 */
                get: function () {
                    return this._sweepAngle;
                },
                set: function (value) {
                    if (value != this._sweepAngle) {
                        this._sweepAngle = wijmo.clamp(wijmo.asNumber(value, false), -360, 360);
                        this.invalidate();
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(RadialGauge.prototype, "autoScale", {
                /**
                 * Gets or sets a value that indicates whether the gauge automatically scales to
                 * fill the host element.
                 */
                get: function () {
                    return this._autoScale;
                },
                set: function (value) {
                    if (value != this._autoScale) {
                        this._autoScale = wijmo.asBoolean(value);
                        this.invalidate();
                    }
                },
                enumerable: true,
                configurable: true
            });
            // virtual methods
            /**
             * Refreshes the control.
             *
             * @param fullUpdate Indicates whether to update the control layout as well as the content.
             */
            RadialGauge.prototype.refresh = function (fullUpdate) {
                if (fullUpdate === void 0) { fullUpdate = true; }
                // clear viewbox
                this._setAttribute(this._svg, 'viewBox', null);
                // cache svg rect to work around a weird problem in Chrome
                this._rcSvg = wijmo.Rect.fromBoundingRect(this._dSvg.getBoundingClientRect());
                // update gauge
                _super.prototype.refresh.call(this, fullUpdate);
                // clear transform matrix
                this._ctmInv = null;
                this._ptSvg = null;
                // set viewbox to auto-scale
                if (this._autoScale) {
                    // clear viewbox first
                    this._setAttribute(this._svg, 'viewBox', '');
                    // measure
                    var rc = wijmo.Rect.fromBoundingRect(gauge.Gauge._getBBox(this._pFace));
                    if ((this.showText & gauge.ShowText.Value) != 0) {
                        rc = wijmo.Rect.union(rc, wijmo.Rect.fromBoundingRect(gauge.Gauge._getBBox(this._tValue)));
                    }
                    if ((this.showText & gauge.ShowText.MinMax) != 0) {
                        rc = wijmo.Rect.union(rc, wijmo.Rect.fromBoundingRect(gauge.Gauge._getBBox(this._tMin)));
                        rc = wijmo.Rect.union(rc, wijmo.Rect.fromBoundingRect(gauge.Gauge._getBBox(this._tMax)));
                    }
                    // apply viewbox
                    var viewBox = [this._fix(rc.left), this._fix(rc.top), this._fix(rc.width), this._fix(rc.height)].join(' ');
                    this._setAttribute(this._svg, 'viewBox', viewBox);
                    // save transform matrix for hit-testing (_getValueFromPoint)
                    this._ctmInv = this._pFace.getCTM().inverse();
                    this._ptSvg = this._svg.createSVGPoint();
                }
            };
            // updates the element for a given range
            RadialGauge.prototype._updateRangeElement = function (e, rng, value) {
                if (this._rcSvg) {
                    var rc = this._rcSvg, center = new wijmo.Point(rc.width / 2, rc.height / 2), radius = Math.min(rc.width, rc.height) / 2, fThick = radius * this.thickness, rThick = fThick * rng.thickness, outer = radius - (fThick - rThick) / 2, inner = outer - rThick, start = this.startAngle + 180, sweep = this.sweepAngle, face = rng == this._face, ps = face ? 0 : this._getPercent(rng.min), pe = face ? 1 : this._getPercent(value), rngStart = start + sweep * ps, rngSweep = sweep * (pe - ps);
                    // update path
                    this._updateSegment(e, center, outer, inner, rngStart, rngSweep);
                    // update thumb
                    if (rng == this._pointer && this.thumbSize > 0) {
                        var color = this._animColor ? this._animColor : this._getPointerColor(rng.max), pt = this._getPoint(center, start + sweep * this._getPercent(value), (outer + inner) / 2), ce = this._cValue;
                        this._setAttribute(ce, 'cx', this._fix(pt.x));
                        this._setAttribute(ce, 'cy', this._fix(pt.y));
                        this._setAttribute(ce, 'style', color ? 'fill:' + color : null);
                        this._setAttribute(ce, 'r', this._fix(this.thumbSize / 2));
                    }
                }
            };
            // update the content and position of the text elements
            RadialGauge.prototype._updateText = function () {
                if (this._rcSvg) {
                    var rc = this._rcSvg, center = new wijmo.Point(rc.width / 2, rc.height / 2), outer = Math.min(rc.width, rc.height) / 2, inner = Math.max(0, outer * (1 - this.thickness)), start = this.startAngle + 180, sweep = this.sweepAngle;
                    // show thumb if it has a size
                    this._showElement(this._cValue, this.thumbSize > 0);
                    // hide min/max if sweep angle > 300 degrees
                    var show = (this.showText & gauge.ShowText.MinMax) != 0 && Math.abs(sweep) <= 300;
                    this._showElement(this._tMin, show);
                    this._showElement(this._tMax, show);
                    // update text/position
                    this._centerText(this._tValue, this.value, center);
                    var offset = 10 * (this.sweepAngle < 0 ? -1 : +1);
                    this._centerText(this._tMin, this.min, this._getPoint(center, start - offset, (outer + inner) / 2));
                    this._centerText(this._tMax, this.max, this._getPoint(center, start + sweep + offset, (outer + inner) / 2));
                }
            };
            // update the tickmarks
            RadialGauge.prototype._updateTicks = function () {
                var d = '';
                if (this.showTicks && this.step) {
                    var rc = this._rcSvg, ctr = new wijmo.Point(rc.width / 2, rc.height / 2), radius = Math.min(rc.width, rc.height) / 2, fThick = radius * this.thickness, rThick = fThick * this._face.thickness, outer = radius - (fThick - rThick) / 2, inner = outer - rThick;
                    for (var t = this.min + this.step; t < this.max; t += this.step) {
                        var angle = this.startAngle + 180 + this.sweepAngle * this._getPercent(t), p1 = this._fix(this._getPoint(ctr, angle, inner)), p2 = this._fix(this._getPoint(ctr, angle, outer));
                        d += 'M ' + p1 + ' L ' + p2 + ' ';
                    }
                }
                this._pTicks.setAttribute('d', d);
            };
            // draws a radial segment at the specified position
            RadialGauge.prototype._updateSegment = function (path, ctr, rOut, rIn, start, sweep) {
                sweep = Math.min(Math.max(sweep, -359.99), 359.99);
                var p1 = this._getPoint(ctr, start, rIn), p2 = this._getPoint(ctr, start, rOut), p3 = this._getPoint(ctr, start + sweep, rOut), p4 = this._getPoint(ctr, start + sweep, rIn);
                var data = {
                    large: Math.abs(sweep) > 180 ? 1 : 0,
                    cw: sweep > 0 ? 1 : 0,
                    ccw: sweep > 0 ? 0 : 1,
                    or: this._fix(rOut),
                    ir: this._fix(rIn),
                    p1: this._fix(p1),
                    p2: this._fix(p2),
                    p3: this._fix(p3),
                    p4: this._fix(p4)
                };
                var content = wijmo.format('M {p1} ' +
                    'L {p2} A {or} {or} 0 {large} {cw} {p3} ' +
                    'L {p4} A {ir} {ir} 0 {large} {ccw} {p1} Z', data);
                path.setAttribute('d', content);
            };
            // converts polar to Cartesian coordinates
            RadialGauge.prototype._getPoint = function (ctr, angle, radius) {
                angle = angle * Math.PI / 180;
                return new wijmo.Point(ctr.x + radius * Math.cos(angle), ctr.y + radius * Math.sin(angle));
            };
            // gets the gauge value at a given point (in gauge client coordinates)
            RadialGauge.prototype._getValueFromPoint = function (pt) {
                // convert client coordinates to SVG viewport
                // the getCTM matrix transforms viewport into client coordinates
                // the inverse matrix transforms client into viewport, which is what we want
                if (this.autoScale && this._ctmInv) {
                    this._ptSvg.x = pt.x;
                    this._ptSvg.y = pt.y;
                    this._ptSvg = this._ptSvg.matrixTransform(this._ctmInv);
                    pt.x = this._ptSvg.x;
                    pt.y = this._ptSvg.y;
                }
                // sanity
                if (!this._rcSvg) {
                    return null;
                }
                // calculate geometry
                var rc = this._rcSvg, center = new wijmo.Point(rc.width / 2, rc.height / 2), outer = Math.min(rc.width, rc.height) / 2, inner = outer * (1 - this.thickness), dx = pt.x - center.x, dy = pt.y - center.y;
                // check that the point is within the face
                var r2 = dy * dy + dx * dx;
                if (r2 > outer * outer + 16 || r2 < inner * inner - 16) {
                    return null;
                }
                // calculate angle, percentage
                var ang = (Math.PI - Math.atan2(-dy, dx)) * 180 / Math.PI, start = this.startAngle, sweep = this.sweepAngle;
                if (sweep > 0) {
                    while (ang < start)
                        ang += 360;
                    while (ang > start + sweep)
                        ang -= 360;
                }
                else {
                    while (ang < start + sweep)
                        ang += 360;
                    while (ang > start)
                        ang -= 360;
                }
                var pct = Math.abs(ang - start) / Math.abs(sweep);
                return this.min + pct * (this.max - this.min);
            };
            return RadialGauge;
        }(gauge.Gauge));
        gauge.RadialGauge = RadialGauge;
    })(gauge = wijmo.gauge || (wijmo.gauge = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=RadialGauge.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var gauge;
    (function (gauge) {
        'use strict';
        /**
         * The @see:BulletGraph is a type of linear gauge designed specifically for use
         * in dashboards. It displays a single key measure along with a comparative
         * measure and qualitative ranges to instantly signal whether the measure is
         * good, bad, or in some other state.
         *
         * Bullet Graphs were created and popularized by dashboard design expert
         * Stephen Few. You can find more details and examples on
         * <a href="http://en.wikipedia.org/wiki/Bullet_graph">Wikipedia</a>.
         *
         * @fiddle:8uxb1vwf
         */
        var BulletGraph = (function (_super) {
            __extends(BulletGraph, _super);
            /**
             * Initializes a new instance of the @see:BulletGraph class.
             *
             * @param element The DOM element that hosts the control, or a selector for the host element (e.g. '#theCtrl').
             * @param options The JavaScript object containing initialization data for the control.
             */
            function BulletGraph(element, options) {
                _super.call(this, element, null);
                // customize
                wijmo.addClass(this.hostElement, 'wj-bulletgraph');
                this._pointer.thickness = .35;
                // add reference ranges
                this._rngTarget = new gauge.Range('target');
                this._rngTarget.thickness = .8;
                this._rngTarget.color = 'black';
                this._rngGood = new gauge.Range('good');
                this._rngGood.color = 'rgba(0,0,0,.15)';
                this._rngBad = new gauge.Range('bad');
                this._rngBad.color = 'rgba(0,0,0,.3)';
                this.ranges.push(this._rngBad);
                this.ranges.push(this._rngGood);
                this.ranges.push(this._rngTarget);
                // initialize control options
                this.initialize(options);
            }
            Object.defineProperty(BulletGraph.prototype, "target", {
                /**
                 * Gets or sets the target value for the measure.
                 */
                get: function () {
                    return this._rngTarget.max;
                },
                set: function (value) {
                    this._rngTarget.max = value;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(BulletGraph.prototype, "good", {
                /**
                 * Gets or sets a reference value considered good for the measure.
                 */
                get: function () {
                    return this._rngGood.max;
                },
                set: function (value) {
                    this._rngGood.max = value;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(BulletGraph.prototype, "bad", {
                /**
                 * Gets or sets a reference value considered bad for the measure.
                 */
                get: function () {
                    return this._rngBad.max;
                },
                set: function (value) {
                    this._rngBad.max = value;
                },
                enumerable: true,
                configurable: true
            });
            // ** implementation
            // gets a rectangle that represents a Range
            BulletGraph.prototype._getRangeRect = function (rng, value) {
                if (value === void 0) { value = rng.max; }
                // let base class calculate the rectangle
                var rc = _super.prototype._getRangeRect.call(this, rng, value);
                // make target range rect look like a bullet
                if (rng == this._rngTarget) {
                    switch (this.direction) {
                        case gauge.GaugeDirection.Right:
                            rc.left = rc.right - 1;
                            rc.width = 3;
                            break;
                        case gauge.GaugeDirection.Left:
                            rc.width = 3;
                            break;
                        case gauge.GaugeDirection.Up:
                            rc.height = 3;
                            break;
                        case gauge.GaugeDirection.Down:
                            rc.top = rc.bottom - 1;
                            rc.height = 3;
                            break;
                    }
                }
                // done
                return rc;
            };
            return BulletGraph;
        }(gauge.LinearGauge));
        gauge.BulletGraph = BulletGraph;
    })(gauge = wijmo.gauge || (wijmo.gauge = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=BulletGraph.js.map
var wijmo;
(function (wijmo) {
    var gauge;
    (function (gauge) {
        'use strict';
        /**
         * Defines ranges to be used with @see:Gauge controls.
         *
         * @see:Range objects have @see:min and @see:max properties that
         * define the range's domain, as well as @see:color and @see:thickness
         * properties that define the range's appearance.
         *
         * Every @see:Gauge control has at least two ranges:
         * the 'face' defines the minimum and maximum values for the gauge, and
         * the 'pointer' displays the gauge's current value.
         *
         * In addition to the built-in ranges, gauges may have additional
         * ranges used to display regions of significance (for example,
         * low, medium, and high values).
         */
        var Range = (function () {
            /**
             * Initializes a new instance of the @see:Range class.
             *
             * @param name The name of the range.
             */
            function Range(name) {
                this._min = 0;
                this._max = 100;
                this._thickness = 1;
                /**
                 * Occurs when the value of a property in this @see:Range changes.
                 */
                this.propertyChanged = new wijmo.Event();
                this._name = name;
            }
            Object.defineProperty(Range.prototype, "min", {
                /**
                 * Gets or sets the minimum value for this range.
                 */
                get: function () {
                    return this._min;
                },
                set: function (value) {
                    this._setProp('_min', wijmo.asNumber(value, true));
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Range.prototype, "max", {
                /**
                 * Gets or sets the maximum value for this range.
                 */
                get: function () {
                    return this._max;
                },
                set: function (value) {
                    this._setProp('_max', wijmo.asNumber(value, true));
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Range.prototype, "color", {
                /**
                 * Gets or sets the color used to display this range.
                 */
                get: function () {
                    return this._color;
                },
                set: function (value) {
                    this._setProp('_color', wijmo.asString(value));
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Range.prototype, "thickness", {
                /**
                 * Gets or sets the thickness of this range as a percentage of
                 * the parent gauge's thickness.
                 */
                get: function () {
                    return this._thickness;
                },
                set: function (value) {
                    this._setProp('_thickness', wijmo.clamp(wijmo.asNumber(value), 0, 1));
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(Range.prototype, "name", {
                /**
                 * Gets or sets the name of this @see:Range.
                 */
                get: function () {
                    return this._name;
                },
                set: function (value) {
                    this._setProp('_name', wijmo.asString(value));
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Raises the @see:propertyChanged event.
             *
             * @param e @see:PropertyChangedEventArgs that contains the property
             * name, old, and new values.
             */
            Range.prototype.onPropertyChanged = function (e) {
                this.propertyChanged.raise(this, e);
            };
            // ** implementation
            // sets property value and notifies about the change
            Range.prototype._setProp = function (name, value) {
                var oldValue = this[name];
                if (value != oldValue) {
                    this[name] = value;
                    var e = new wijmo.PropertyChangedEventArgs(name.substr(1), oldValue, value);
                    this.onPropertyChanged(e);
                }
            };
            Range._ctr = 0;
            return Range;
        }());
        gauge.Range = Range;
    })(gauge = wijmo.gauge || (wijmo.gauge = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Range.js.map
