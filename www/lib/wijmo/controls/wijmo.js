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
* Contains utilities used by all controls and modules, as well as the
* @see:Control and @see:Event classes.
*/
var wijmo;
(function (wijmo) {
    'use strict';
    // major (ECMAScript version required).
    // year/trimester.
    // sequential
    var _VERSION = '5.20162.211';
    /**
     * Gets the version of the Wijmo library that is currently loaded.
     */
    function getVersion() {
        return _VERSION;
    }
    wijmo.getVersion = getVersion;
    /**
     * Specifies constants that represent keyboard codes.
     *
     * This enumeration is useful when handling <b>keyDown</b> events.
     */
    (function (Key) {
        /** The backspace key. */
        Key[Key["Back"] = 8] = "Back";
        /** The tab key. */
        Key[Key["Tab"] = 9] = "Tab";
        /** The enter key. */
        Key[Key["Enter"] = 13] = "Enter";
        /** The escape key. */
        Key[Key["Escape"] = 27] = "Escape";
        /** The space key. */
        Key[Key["Space"] = 32] = "Space";
        /** The page up key. */
        Key[Key["PageUp"] = 33] = "PageUp";
        /** The page down key. */
        Key[Key["PageDown"] = 34] = "PageDown";
        /** The end key. */
        Key[Key["End"] = 35] = "End";
        /** The home key. */
        Key[Key["Home"] = 36] = "Home";
        /** The left arrow key. */
        Key[Key["Left"] = 37] = "Left";
        /** The up arrow key. */
        Key[Key["Up"] = 38] = "Up";
        /** The right arrow key. */
        Key[Key["Right"] = 39] = "Right";
        /** The down arrow key. */
        Key[Key["Down"] = 40] = "Down";
        /** The delete key. */
        Key[Key["Delete"] = 46] = "Delete";
        /** The F1 key. */
        Key[Key["F1"] = 112] = "F1";
        /** The F2 key. */
        Key[Key["F2"] = 113] = "F2";
        /** The F3 key. */
        Key[Key["F3"] = 114] = "F3";
        /** The F4 key. */
        Key[Key["F4"] = 115] = "F4";
        /** The F5 key. */
        Key[Key["F5"] = 116] = "F5";
        /** The F6 key. */
        Key[Key["F6"] = 117] = "F6";
        /** The F7 key. */
        Key[Key["F7"] = 118] = "F7";
        /** The F8 key. */
        Key[Key["F8"] = 119] = "F8";
        /** The F9 key. */
        Key[Key["F9"] = 120] = "F9";
        /** The F10 key. */
        Key[Key["F10"] = 121] = "F10";
        /** The F11 key. */
        Key[Key["F11"] = 122] = "F11";
        /** The F12 key. */
        Key[Key["F12"] = 123] = "F12";
    })(wijmo.Key || (wijmo.Key = {}));
    var Key = wijmo.Key;
    /**
     * Specifies constants that represent data types.
     *
     * Use the @see:getType method to get a @see:DataType from a value.
     */
    (function (DataType) {
        /** Object (anything). */
        DataType[DataType["Object"] = 0] = "Object";
        /** String. */
        DataType[DataType["String"] = 1] = "String";
        /** Number. */
        DataType[DataType["Number"] = 2] = "Number";
        /** Boolean. */
        DataType[DataType["Boolean"] = 3] = "Boolean";
        /** Date (date and time). */
        DataType[DataType["Date"] = 4] = "Date";
        /** Array. */
        DataType[DataType["Array"] = 5] = "Array";
    })(wijmo.DataType || (wijmo.DataType = {}));
    var DataType = wijmo.DataType;
    /**
     * Casts a value to a type if possible.
     *
     * @param value Value to cast.
     * @param type Type or interface name to cast to.
     * @return The value passed in if the cast was successful, null otherwise.
     */
    function tryCast(value, type) {
        // null doesn't implement anything
        if (value == null) {
            return null;
        }
        // test for interface implementation (IQueryInterface)
        if (isString(type)) {
            return isFunction(value.implementsInterface) && value.implementsInterface(type) ? value : null;
        }
        // regular type test
        return value instanceof type ? value : null;
    }
    wijmo.tryCast = tryCast;
    /**
     * Determines whether an object is a primitive type (string, number, boolean, or date).
     *
     * @param value Value to test.
     */
    function isPrimitive(value) {
        return isString(value) || isNumber(value) || isBoolean(value) || isDate(value);
    }
    wijmo.isPrimitive = isPrimitive;
    /**
     * Determines whether an object is a string.
     *
     * @param value Value to test.
     */
    function isString(value) {
        return typeof (value) == 'string';
    }
    wijmo.isString = isString;
    /**
     * Determines whether a string is null, empty, or whitespace only.
     *
     * @param value Value to test.
     */
    function isNullOrWhiteSpace(value) {
        return value == null ? true : value.replace(/\s/g, '').length < 1;
    }
    wijmo.isNullOrWhiteSpace = isNullOrWhiteSpace;
    /**
     * Determines whether an object is a number.
     *
     * @param value Value to test.
     */
    function isNumber(value) {
        return typeof (value) == 'number';
    }
    wijmo.isNumber = isNumber;
    /**
     * Determines whether an object is an integer.
     *
     * @param value Value to test.
     */
    function isInt(value) {
        return isNumber(value) && value == Math.round(value);
    }
    wijmo.isInt = isInt;
    /**
     * Determines whether an object is a Boolean.
     *
     * @param value Value to test.
     */
    function isBoolean(value) {
        return typeof (value) == 'boolean';
    }
    wijmo.isBoolean = isBoolean;
    /**
     * Determines whether an object is a function.
     *
     * @param value Value to test.
     */
    function isFunction(value) {
        return typeof (value) == 'function';
    }
    wijmo.isFunction = isFunction;
    /**
     * Determines whether an object is undefined.
     *
     * @param value Value to test.
     */
    function isUndefined(value) {
        return typeof value == 'undefined';
    }
    wijmo.isUndefined = isUndefined;
    /**
     * Determines whether an object is a Date.
     *
     * @param value Value to test.
     */
    function isDate(value) {
        return value instanceof Date && !isNaN(value.getTime());
    }
    wijmo.isDate = isDate;
    /**
     * Determines whether an object is an Array.
     *
     * @param value Value to test.
     */
    function isArray(value) {
        return value instanceof Array;
    }
    wijmo.isArray = isArray;
    /**
     * Determines whether a value is an object
     * (as opposed to a value type, an array, or a Date).
     *
     * @param value Value to test.
     */
    function isObject(value) {
        return value != null && typeof value == 'object' && !isDate(value) && !isArray(value);
    }
    wijmo.isObject = isObject;
    /**
     * Converts mouse or touch event arguments into a @see:Point in page coordinates.
     */
    function mouseToPage(e) {
        // accept Point objects
        if (e instanceof Point) {
            return e;
        }
        // accept touch events
        if (e.touches && e.touches.length > 0) {
            e = e.touches[0];
        }
        // accept mouse events
        // NOTE: we should be able to use pageX/Y properties, but those may return
        // wrong values (e.g. Android with zoomed screens); so we get the client 
        // coordinates and apply the page offset ourselves instead...
        if (isNumber(e.clientX) && isNumber(e.clientY)) {
            return new Point(e.clientX + pageXOffset, e.clientY + pageYOffset);
        }
        // wrong parameter type...
        throw 'Mouse or touch event expected.';
    }
    wijmo.mouseToPage = mouseToPage;
    /**
     * Gets the type of a value.
     *
     * @param value Value to test.
     * @return A @see:DataType value representing the type of the value passed in.
     */
    function getType(value) {
        if (isNumber(value))
            return DataType.Number;
        if (isBoolean(value))
            return DataType.Boolean;
        if (isDate(value))
            return DataType.Date;
        if (isString(value))
            return DataType.String;
        if (isArray(value))
            return DataType.Array;
        return DataType.Object;
    }
    wijmo.getType = getType;
    /**
     * Changes the type of a value.
     *
     * If the conversion fails, the original value is returned. To check if a
     * conversion succeeded, you should check the type of the returned value.
     *
     * @param value Value to convert.
     * @param type @see:DataType to convert the value to.
     * @param format Format to use when converting to or from strings.
     * @return The converted value, or the original value if a conversion was not possible.
     */
    function changeType(value, type, format) {
        if (value != null) {
            // convert strings to numbers, dates, or booleans
            if (isString(value)) {
                switch (type) {
                    case DataType.Number:
                        var num = wijmo.Globalize.parseFloat(value, format);
                        return isNaN(num) ? value : num;
                    case DataType.Date:
                        var date = wijmo.Globalize.parseDate(value, format);
                        if (!date && !format && value) {
                            date = new Date(value); // fallback on JavaScript parser
                        }
                        return date && isFinite(date.getTime()) ? date : value;
                    case DataType.Boolean:
                        switch (value.toLowerCase()) {
                            case 'true': return true;
                            case 'false': return false;
                        }
                        return value; // TFS 125067
                }
            }
            // convert anything to string
            if (type == DataType.String) {
                return wijmo.Globalize.format(value, format);
            }
        }
        // did not convert...
        //console.log('did not convert "' + value + '" to type ' + DataType[type]);
        return value;
    }
    wijmo.changeType = changeType;
    /**
     * Rounds or truncates a number to a specified precision.
     *
     * @param value Value to round or truncate.
     * @param prec Number of decimal digits for the result.
     * @param truncate Whether to truncate or round the original value.
     */
    function toFixed(value, prec, truncate) {
        if (truncate) {
            var s = value.toString(), decPos = s.indexOf('.');
            if (decPos > -1) {
                s = s.substr(0, decPos + 1 + prec);
                value = parseFloat(s);
            }
        }
        else {
            var s = value.toFixed(prec);
            value = parseFloat(s);
        }
        return value;
    }
    wijmo.toFixed = toFixed;
    /**
     * Replaces each format item in a specified string with the text equivalent of an
     * object's value.
     *
     * The function works by replacing parts of the <b>formatString</b> with the pattern
     * '{name:format}' with properties of the <b>data</b> parameter. For example:
     *
     * <pre>
     * var data = { name: 'Joe', amount: 123456 };
     * var msg = wijmo.format('Hello {name}, you won {amount:n2}!', data);
     * </pre>
     *
     * The optional <b>formatFunction</b> allows you to customize the content by providing
     * context-sensitive formatting. If provided, the format function gets called for each
     * format element and gets passed the data object, the parameter name, the format,
     * and the value; it should return an output string. For example:
     *
     * <pre>
     * var data = { name: 'Joe', amount: 123456 };
     * var msg = wijmo.format('Hello {name}, you won {amount:n2}!', data,
     *             function (data, name, fmt, val) {
     *               if (wijmo.isString(data[name])) {
     *                   val = wijmo.escapeHtml(data[name]);
     *               }
     *               return val;
     *             });
     * </pre>
     *
     * @param format A composite format string.
     * @param data The data object used to build the string.
     * @param formatFunction An optional function used to format items in context.
     * @return The formatted string.
     */
    function format(format, data, formatFunction) {
        format = asString(format);
        return format.replace(/\{(.*?)(:(.*?))?\}/g, function (match, name, x, fmt) {
            var val = match;
            if (name && name[0] != '{' && data) {
                // get the value
                val = data[name];
                // apply static format
                if (fmt) {
                    val = wijmo.Globalize.format(val, fmt);
                }
                // apply format function
                if (formatFunction) {
                    val = formatFunction(data, name, fmt, val);
                }
            }
            return val == null ? '' : val;
        });
    }
    wijmo.format = format;
    /**
     * Clamps a value between a minimum and a maximum.
     *
     * @param value Original value.
     * @param min Minimum allowed value.
     * @param max Maximum allowed value.
     */
    function clamp(value, min, max) {
        if (value != null) {
            if (max != null && value > max)
                value = max;
            if (min != null && value < min)
                value = min;
        }
        return value;
    }
    wijmo.clamp = clamp;
    /**
     * Copies properties from an object to another.
     *
     * This method is typically used to initialize controls and other Wijmo objects
     * by setting their properties and assigning event handlers.
     *
     * The destination object must define all the properties defined in the source,
     * or an error will be thrown.
     *
     * @param dst The destination object.
     * @param src The source object.
     */
    function copy(dst, src) {
        if (src) {
            for (var key in src) {
                if (key[0] != '_') {
                    assert(key in dst, 'Unknown property "' + key + '".');
                    var value = src[key];
                    if (!dst._copy || !dst._copy(key, value)) {
                        if (dst[key] instanceof wijmo.Event && isFunction(value)) {
                            dst[key].addHandler(value); // add event handler
                        }
                        else if (isObject(value) && dst[key] && key != 'itemsSource') {
                            copy(dst[key], value); // copy sub-objects
                        }
                        else {
                            dst[key] = value; // assign values
                        }
                    }
                }
            }
        }
    }
    wijmo.copy = copy;
    /**
     * Throws an exception if a condition is false.
     *
     * @param condition Condition expected to be true.
     * @param msg Message of the exception if the condition is not true.
     */
    function assert(condition, msg) {
        if (!condition) {
            throw '** Assertion failed in Wijmo: ' + msg;
        }
    }
    wijmo.assert = assert;
    /**
     * Outputs a message to indicate a member has been deprecated.
     *
     * @param oldMember Member that has been deprecated.
     * @param newMember Member that replaces the one that has been deprecated.
     */
    function _deprecated(oldMember, newMember) {
        console.error('** WARNING: "' + oldMember + '" has been deprecated; please use "' + newMember + '" instead.');
    }
    wijmo._deprecated = _deprecated;
    /**
     * Asserts that a value is a string.
     *
     * @param value Value supposed to be a string.
     * @param nullOK Whether null values are acceptable.
     * @return The string passed in.
     */
    function asString(value, nullOK) {
        if (nullOK === void 0) { nullOK = true; }
        assert((nullOK && value == null) || isString(value), 'String expected.');
        return value;
    }
    wijmo.asString = asString;
    /**
     * Asserts that a value is a number.
     *
     * @param value Value supposed to be numeric.
     * @param nullOK Whether null values are acceptable.
     * @param positive Whether to accept only positive numeric values.
     * @return The number passed in.
     */
    function asNumber(value, nullOK, positive) {
        if (nullOK === void 0) { nullOK = false; }
        if (positive === void 0) { positive = false; }
        assert((nullOK && value == null) || isNumber(value), 'Number expected.');
        if (positive && value && value < 0)
            throw 'Positive number expected.';
        return value;
    }
    wijmo.asNumber = asNumber;
    /**
     * Asserts that a value is an integer.
     *
     * @param value Value supposed to be an integer.
     * @param nullOK Whether null values are acceptable.
     * @param positive Whether to accept only positive integers.
     * @return The number passed in.
     */
    function asInt(value, nullOK, positive) {
        if (nullOK === void 0) { nullOK = false; }
        if (positive === void 0) { positive = false; }
        assert((nullOK && value == null) || isInt(value), 'Integer expected.');
        if (positive && value && value < 0)
            throw 'Positive integer expected.';
        return value;
    }
    wijmo.asInt = asInt;
    /**
     * Asserts that a value is a Boolean.
     *
     * @param value Value supposed to be Boolean.
     * @param nullOK Whether null values are acceptable.
     * @return The Boolean passed in.
     */
    function asBoolean(value, nullOK) {
        if (nullOK === void 0) { nullOK = false; }
        assert((nullOK && value == null) || isBoolean(value), 'Boolean expected.');
        return value;
    }
    wijmo.asBoolean = asBoolean;
    /**
     * Asserts that a value is a Date.
     *
     * @param value Value supposed to be a Date.
     * @param nullOK Whether null values are acceptable.
     * @return The Date passed in.
     */
    function asDate(value, nullOK) {
        if (nullOK === void 0) { nullOK = false; }
        // parse strings into dates using RFC 3339 pattern ([yyyy-MM-dd] [hh:mm[:ss]])
        if (isString(value)) {
            var dt = changeType(value, DataType.Date, 'r');
            if (isDate(dt)) {
                value = dt;
            }
        }
        assert((nullOK && value == null) || isDate(value), 'Date expected.');
        return value;
    }
    wijmo.asDate = asDate;
    /**
     * Asserts that a value is a function.
     *
     * @param value Value supposed to be a function.
     * @param nullOK Whether null values are acceptable.
     * @return The function passed in.
     */
    function asFunction(value, nullOK) {
        if (nullOK === void 0) { nullOK = true; }
        assert((nullOK && value == null) || isFunction(value), 'Function expected.');
        return value;
    }
    wijmo.asFunction = asFunction;
    /**
     * Asserts that a value is an array.
     *
     * @param value Value supposed to be an array.
     * @param nullOK Whether null values are acceptable.
     * @return The array passed in.
     */
    function asArray(value, nullOK) {
        if (nullOK === void 0) { nullOK = true; }
        assert((nullOK && value == null) || isArray(value), 'Array expected.');
        return value;
    }
    wijmo.asArray = asArray;
    /**
     * Asserts that a value is an instance of a given type.
     *
     * @param value Value to be checked.
     * @param type Type of value expected.
     * @param nullOK Whether null values are acceptable.
     * @return The value passed in.
     */
    function asType(value, type, nullOK) {
        if (nullOK === void 0) { nullOK = false; }
        value = tryCast(value, type);
        assert(nullOK || value != null, type + ' expected.');
        return value;
    }
    wijmo.asType = asType;
    /**
     * Asserts that a value is a valid setting for an enumeration.
     *
     * @param value Value supposed to be a member of the enumeration.
     * @param enumType Enumeration to test for.
     * @param nullOK Whether null values are acceptable.
     * @return The value passed in.
     */
    function asEnum(value, enumType, nullOK) {
        if (nullOK === void 0) { nullOK = false; }
        if (value == null && nullOK)
            return null;
        var e = enumType[value];
        assert(e != null, 'Invalid enum value.');
        return isNumber(e) ? e : value;
    }
    wijmo.asEnum = asEnum;
    /**
     * Asserts that a value is an @see:ICollectionView or an Array.
     *
     * @param value Array or @see:ICollectionView.
     * @param nullOK Whether null values are acceptable.
     * @return The @see:ICollectionView that was passed in or a @see:CollectionView
     * created from the array that was passed in.
     */
    function asCollectionView(value, nullOK) {
        if (nullOK === void 0) { nullOK = true; }
        if (value == null && nullOK) {
            return null;
        }
        var cv = tryCast(value, 'ICollectionView');
        if (cv != null) {
            return cv;
        }
        if (!isArray(value)) {
            assert(false, 'Array or ICollectionView expected.');
        }
        return new wijmo.collections.CollectionView(value);
    }
    wijmo.asCollectionView = asCollectionView;
    /**
     * Checks whether an @see:ICollectionView is defined and not empty.
     *
     * @param value @see:ICollectionView to check.
     */
    function hasItems(value) {
        return value && value.items && value.items.length;
    }
    wijmo.hasItems = hasItems;
    /**
     * Converts a camel-cased string into a header-type string by capitalizing the first letter
     * and adding spaces before uppercase characters preceded by lower-case characters.
     *
     * For example, 'somePropertyName' becomes 'Some Property Name'.
     *
     * @param text String to convert to header case.
     */
    function toHeaderCase(text) {
        return text && text.length
            ? text[0].toUpperCase() + text.substr(1).replace(/([a-z])([A-Z])/g, '$1 $2')
            : '';
    }
    wijmo.toHeaderCase = toHeaderCase;
    /**
     * Escapes a string by replacing HTML characters as text entities.
     *
     * Strings entered by uses should always be escaped before they are displayed
     * in HTML pages. This ensures page integrity and prevents HTML/javascript
     * injection attacks.
     *
     * @param text Text to escape.
     * @return An HTML-escaped version of the original string.
     */
    function escapeHtml(text) {
        if (isString(text)) {
            text = text.replace(/[&<>"'\/]/g, function (s) {
                return _ENTITYMAP[s];
            });
        }
        return text;
    }
    wijmo.escapeHtml = escapeHtml;
    var _ENTITYMAP = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;',
        '/': '&#x2F;'
    };
    /**
     * Checks whether an element has a class.
     *
     * @param e Element to check.
     * @param className Class to check for.
     */
    function hasClass(e, className) {
        // NOTE: using e.getAttribute('class') instead of e.classNames
        // so this works with SVG as well as regular HTML elements.
        // NOTE: don't use word boundaries because class names may have 
        // hyphens and other non-word boundary characters
        if (e && e.getAttribute) {
            var rx = new RegExp('(\\s|^)' + className + '(\\s|$)');
            return e && rx.test(e.getAttribute('class'));
        }
        return false;
    }
    wijmo.hasClass = hasClass;
    /**
     * Removes a class from an element.
     *
     * @param e Element that will have the class removed.
     * @param className Class to remove form the element.
     */
    function removeClass(e, className) {
        if (e && className && e.setAttribute && hasClass(e, className)) {
            var rx = new RegExp('((\\s|^)' + className + '(\\s|$))', 'g'), cn = e.getAttribute('class');
            cn = cn.replace(rx, ' ').replace(/ +/g, ' ').trim();
            e.setAttribute('class', cn);
        }
    }
    wijmo.removeClass = removeClass;
    /**
     * Adds a class to an element.
     *
     * @param e Element that will have the class added.
     * @param className Class to add to the element.
     */
    function addClass(e, className) {
        if (e && className && e.setAttribute && !hasClass(e, className)) {
            var cn = e.getAttribute('class');
            e.setAttribute('class', cn ? cn + ' ' + className : className);
        }
    }
    wijmo.addClass = addClass;
    /**
     * Adds or removes a class to or from an element.
     *
     * @param e Element that will have the class added.
     * @param className Class to add or remove.
     * @param addOrRemove Whether to add or remove the class.
     * Use true to add class to element and false to remove class from element.
     */
    function toggleClass(e, className, addOrRemove) {
        if (addOrRemove) {
            addClass(e, className);
        }
        else {
            removeClass(e, className);
        }
    }
    wijmo.toggleClass = toggleClass;
    /**
     * Sets the start and end positions of a selection in a text field.
     *
     * This method is similar to the native @see:setSelectionRange method
     * in HTMLInputElement objects, except it checks for conditions that
     * may cause exceptions (element not in the DOM, disabled, or hidden).
     *
     * @param start Offset into the text field for the start of the selection.
     * @param end Offset into the text field for the end of the selection.
     */
    function setSelectionRange(e, start, end) {
        if (end === void 0) { end = start; }
        e = asType(e, HTMLInputElement);
        if (contains(document.body, e) && !e.disabled && e.style.display != 'none') {
            try {
                e.focus(); // needed in Chrome (TFS 124102, 142672)
                e.setSelectionRange(asNumber(start), asNumber(end));
            }
            catch (x) { }
        }
    }
    wijmo.setSelectionRange = setSelectionRange;
    /**
     * Gets a reference to the element that contains the focus,
     * accounting for shadow document fragments.
     */
    function getActiveElement() {
        var ae = document.activeElement;
        if (ae) {
            // account for shadowRoot: https://github.com/w3c/webcomponents/issues/358)
            var shadowRoot = ae['shadowRoot'];
            if (shadowRoot && shadowRoot.activeElement) {
                ae = shadowRoot.activeElement;
            }
        }
        return ae;
    }
    wijmo.getActiveElement = getActiveElement;
    /**
     * Moves the focus to the next/previous/first focusable child element.
     *
     * @param parent Parent element.
     * @param offset Offset to use when moving the focus (use zero to focus on the first focusable child).
     */
    function moveFocus(parent, offset) {
        // build array of focusable elements (including divs and spans)
        var inputs = parent.querySelectorAll('input,select,textarea,button,a,div,span'), focusable = [];
        for (var i = 0; i < inputs.length; i++) {
            var el = inputs[i];
            if (!el.disabled && el.tabIndex > -1 && el.clientHeight > 0 &&
                !closest(el, '[disabled],.wj-state-disabled')) {
                focusable.push(el);
            }
        }
        // calculate focus index
        var index = 0;
        if (offset) {
            var i = focusable.indexOf(getActiveElement());
            if (i > -1) {
                index = (i + offset + focusable.length) % focusable.length; // TFS 152269, 152163
            }
        }
        // move focus to element at the focus index
        if (index < focusable.length) {
            var el = focusable[index];
            el.focus();
            if (el instanceof HTMLInputElement) {
                el.select(); // TFS 190336
            }
        }
    }
    wijmo.moveFocus = moveFocus;
    // ** jQuery replacement methods
    /**
     * Gets an element from a jQuery-style selector.
     *
     * @param selector An element, a query selector string, or a jQuery object.
     */
    function getElement(selector) {
        if (selector instanceof Element)
            return selector;
        if (isString(selector))
            return document.querySelector(selector);
        if (selector && selector.jquery)
            return selector[0];
        return null;
    }
    wijmo.getElement = getElement;
    /**
     * Creates an element from an HTML string.
     *
     * @param html HTML fragment to convert into an HTMLElement.
     * @param appendTo Optional HTMLElement to append the new element to.
     * @return The new element.
     */
    function createElement(html, appendTo) {
        var div = document.createElement('div');
        div.innerHTML = html;
        var e = div.removeChild(div.firstChild);
        return (appendTo instanceof HTMLElement)
            ? appendTo.appendChild(e)
            : e;
    }
    wijmo.createElement = createElement;
    /**
     * Sets the text content of an element.
     *
     * @param e Element that will have its content updated.
     * @param text Plain text to be assigned to the element.
     */
    function setText(e, text) {
        // clear
        if (text == null) {
            if (e.hasChildNodes()) {
                // this causes serious/weird problems in IE, so DON'T DO IT!!! 
                //e.innerHTML = '';
                // this works, but seems inefficient/convoluted
                //var dr = document.createRange();
                //dr.setStart(e, 0);
                //dr.setEnd(e, e.childNodes.length);
                //dr.deleteContents();
                // seems like the best option (simple and works)
                e.textContent = '';
            }
            return;
        }
        // set text
        var fc = e.firstChild;
        if (e.childNodes.length == 1 && fc.nodeType == 3) {
            if (fc.nodeValue != text) {
                fc.nodeValue = text; // update text directly in the text node
            }
        }
        else if (fc || text) {
            e.textContent = text; // something else, set the textContent
        }
    }
    wijmo.setText = setText;
    /**
     * Checks whether an HTML element contains another.
     *
     * @param parent Parent element.
     * @param child Child element.
     * @return True if the parent element contains the child element.
     */
    function contains(parent, child) {
        for (var e = child; e && parent;) {
            if (e === parent)
                return true; // found!
            e = e.parentNode || e['host']; // move up to parent node or host (shadow DOM)
        }
        return false;
    }
    wijmo.contains = contains;
    /**
     * Finds the closest ancestor that satisfies a selector.
     *
     * @param e Element where the search should start.
     * @param selector A string containing a selector expression to match elements against.
     * @return The closest ancestor that satisfies the selector (including the original element), or null if not found.
     */
    function closest(e, selector) {
        var matches = e ? (e.matches || e.webkitMatchesSelector || e.mozMatchesSelector || e.msMatchesSelector) : null;
        if (matches) {
            for (; e; e = e.parentNode) {
                if (e instanceof Element && matches.call(e, selector)) {
                    return e;
                }
            }
        }
        return null;
    }
    wijmo.closest = closest;
    /**
     * Enables or disables an element.
     *
     * @param e Element to enable or disable.
     * @param value Whether to enable or disable the element.
     */
    function enable(e, value) {
        // update wj-state-disabled class and disabled attribute on the element
        toggleClass(e, 'wj-state-disabled', !value);
        if (value) {
            e.removeAttribute('disabled');
        }
        else {
            e.setAttribute('disabled', 'true');
        }
        // update disabled attribute on inner input elements (TFS 190939)
        var inputs = e.querySelectorAll('input');
        for (var i = 0; i < inputs.length; i++) {
            var input = inputs[i];
            if (value) {
                input.removeAttribute('disabled');
            }
            else {
                input.setAttribute('disabled', 'true');
            }
        }
    }
    wijmo.enable = enable;
    /**
     * Gets the bounding rectangle of an element in page coordinates.
     *
     * This is similar to the <b>getBoundingClientRect</b> function,
     * except that uses window coordinates, which change when the
     * document scrolls.
     */
    function getElementRect(e) {
        var rc = e.getBoundingClientRect();
        return new Rect(rc.left + pageXOffset, rc.top + pageYOffset, rc.width, rc.height);
    }
    wijmo.getElementRect = getElementRect;
    /**
     * Modifies the style of an element by applying the properties specified in an object.
     *
     * @param e Element or array of elements whose style will be modified.
     * @param css Object containing the style properties to apply to the element.
     */
    function setCss(e, css) {
        // apply to arrays
        if (isArray(e)) {
            e.forEach(function (item) {
                setCss(item, css);
            });
            return;
        }
        // apply to elements
        var s = e.style;
        for (var p in css) {
            // add pixel units to numeric geometric properties
            var val = css[p];
            if (typeof (val) == 'number' &&
                p.match(/width|height|left|top|right|bottom|size|padding|margin'/i)) {
                val += 'px';
            }
            // set the attribute if it changed
            if (s[p] != val) {
                s[p] = val.toString();
            }
        }
    }
    wijmo.setCss = setCss;
    /**
     * Calls a function on a timer with a parameter varying between zero and one.
     *
     * Use this function to create animations by modifying document properties
     * or styles on a timer.
     *
     * For example, the code below changes the opacity of an element from zero
     * to one in one second:
     * <pre>var element = document.getElementById('someElement');
     * animate(function(pct) {
     *   element.style.opacity = pct;
     * }, 1000);</pre>
     *
     * The function returns an interval ID that you can use to stop the
     * animation. This is typically done when you are starting a new animation
     * and wish to suspend other on-going animations on the same element.
     * For example, the code below keeps track of the interval ID and clears
     * if before starting a new animation:
     * <pre>var element = document.getElementById('someElement');
     * if (this._animInterval) {
     *   clearInterval(this._animInterval);
     * }
     * var self = this;
     * self._animInterval = animate(function(pct) {
     *   element.style.opacity = pct;
     *   if (pct == 1) {
     *     self._animInterval = null;
     *   }
     * }, 1000);</pre>
     *
     * @param apply Callback function that modifies the document.
     * The function takes a single parameter that represents a percentage.
     * @param duration The duration of the animation, in milliseconds.
     * @param step The interval between animation frames, in milliseconds.
     * @return An interval id that you can use to suspend the animation.
     */
    function animate(apply, duration, step) {
        if (duration === void 0) { duration = 400; }
        if (step === void 0) { step = 35; }
        asFunction(apply);
        asNumber(duration, false, true);
        asNumber(step, false, true);
        var start = Date.now();
        var timer = setInterval(function () {
            var pct = Math.min(1, (Date.now() - start) / duration); // linear
            pct = Math.sin(pct * Math.PI / 2); // easeOutSin
            pct *= pct; // swing
            requestAnimationFrame(function () {
                apply(pct);
            });
            if (pct >= 1) {
                clearInterval(timer);
            }
        }, step);
        return timer;
    }
    wijmo.animate = animate;
    // ** utility classes
    /**
     * Class that represents a point (with x and y coordinates).
     */
    var Point = (function () {
        /**
         * Initializes a new instance of the @see:Point class.
         *
         * @param x X coordinate of the new Point.
         * @param y Y coordinate of the new Point.
         */
        function Point(x, y) {
            if (x === void 0) { x = 0; }
            if (y === void 0) { y = 0; }
            this.x = asNumber(x);
            this.y = asNumber(y);
        }
        /**
         * Returns true if a @see:Point has the same coordinates as this @see:Point.
         *
         * @param pt @see:Point to compare to this @see:Point.
         */
        Point.prototype.equals = function (pt) {
            return (pt instanceof Point) && this.x == pt.x && this.y == pt.y;
        };
        /**
         * Creates a copy of this @see:Point.
         */
        Point.prototype.clone = function () {
            return new Point(this.x, this.y);
        };
        return Point;
    }());
    wijmo.Point = Point;
    /**
     * Class that represents a size (with width and height).
     */
    var Size = (function () {
        /**
         * Initializes a new instance of the @see:Size class.
         *
         * @param width Width of the new @see:Size.
         * @param height Height of the new @see:Size.
         */
        function Size(width, height) {
            if (width === void 0) { width = 0; }
            if (height === void 0) { height = 0; }
            this.width = asNumber(width);
            this.height = asNumber(height);
        }
        /**
         * Returns true if a @see:Size has the same dimensions as this @see:Size.
         *
         * @param sz @see:Size to compare to this @see:Size.
         */
        Size.prototype.equals = function (sz) {
            return (sz instanceof Size) && this.width == sz.width && this.height == sz.height;
        };
        /**
         * Creates a copy of this @see:Size.
         */
        Size.prototype.clone = function () {
            return new Size(this.width, this.height);
        };
        return Size;
    }());
    wijmo.Size = Size;
    /**
     * Class that represents a rectangle (with left, top, width, and height).
     */
    var Rect = (function () {
        /**
         * Initializes a new instance of the @see:Rect class.
         *
         * @param left Left coordinate of the new @see:Rect.
         * @param top Top coordinate of the new @see:Rect.
         * @param width Width of the new @see:Rect.
         * @param height Height of the new @see:Rect.
         */
        function Rect(left, top, width, height) {
            this.left = asNumber(left);
            this.top = asNumber(top);
            this.width = asNumber(width);
            this.height = asNumber(height);
        }
        Object.defineProperty(Rect.prototype, "right", {
            /**
             * Gets the right coordinate of this @see:Rect.
             */
            get: function () {
                return this.left + this.width;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Rect.prototype, "bottom", {
            /**
             * Gets the bottom coordinate of this @see:Rect.
             */
            get: function () {
                return this.top + this.height;
            },
            enumerable: true,
            configurable: true
        });
        /**
         * Returns true if a @see:Rect has the same coordinates and dimensions
         * as this @see:Rect.
         *
         * @param rc @see:Rect to compare to this @see:Rect.
         */
        Rect.prototype.equals = function (rc) {
            return (rc instanceof Rect) && this.left == rc.left && this.top == rc.top && this.width == rc.width && this.height == rc.height;
        };
        /**
         * Creates a copy of this @see:Rect.
         */
        Rect.prototype.clone = function () {
            return new Rect(this.left, this.top, this.width, this.height);
        };
        /**
         * Creates a @see:Rect from <b>ClientRect</b> or <b>SVGRect</b> objects.
         *
         * @param rc Rectangle obtained by a call to the DOM's <b>getBoundingClientRect</b>
         * or <b>GetBoundingBox</b> methods.
         */
        Rect.fromBoundingRect = function (rc) {
            if (rc.left != null) {
                return new Rect(rc.left, rc.top, rc.width, rc.height);
            }
            else if (rc.x != null) {
                return new Rect(rc.x, rc.y, rc.width, rc.height);
            }
            else {
                assert(false, 'Invalid source rectangle.');
            }
        };
        /**
         * Gets a rectangle that represents the union of two rectangles.
         *
         * @param rc1 First rectangle.
         * @param rc2 Second rectangle.
         */
        Rect.union = function (rc1, rc2) {
            var x = Math.min(rc1.left, rc2.left), y = Math.min(rc1.top, rc2.top), right = Math.max(rc1.right, rc2.right), bottom = Math.max(rc1.bottom, rc2.bottom);
            return new Rect(x, y, right - x, bottom - y);
        };
        /**
         * Gets a rectangle that represents the intersection of two rectangles.
         *
         * @param rc1 First rectangle.
         * @param rc2 Second rectangle.
         */
        Rect.intersection = function (rc1, rc2) {
            var x = Math.max(rc1.left, rc2.left), y = Math.max(rc1.top, rc2.top), right = Math.min(rc1.right, rc2.right), bottom = Math.min(rc1.bottom, rc2.bottom);
            return new Rect(x, y, right - x, bottom - y);
        };
        /**
         * Determines whether the rectangle contains a given point or rectangle.
         *
         * @param pt The @see:Point or @see:Rect to ckeck.
         */
        Rect.prototype.contains = function (pt) {
            if (pt instanceof Point) {
                return pt.x >= this.left && pt.x <= this.right &&
                    pt.y >= this.top && pt.y <= this.bottom;
            }
            else if (pt instanceof Rect) {
                var rc2 = pt;
                return rc2.left >= this.left && rc2.right <= this.right &&
                    rc2.top >= this.top && rc2.bottom <= this.bottom;
            }
            else {
                assert(false, 'Point or Rect expected.');
            }
        };
        /**
         * Creates a rectangle that results from expanding or shrinking a rectangle by the specified amounts.
         *
         * @param dx The amount by which to expand or shrink the left and right sides of the rectangle.
         * @param dy The amount by which to expand or shrink the top and bottom sides of the rectangle.
         */
        Rect.prototype.inflate = function (dx, dy) {
            return new Rect(this.left - dx, this.top - dy, this.width + 2 * dx, this.height + 2 * dy);
        };
        return Rect;
    }());
    wijmo.Rect = Rect;
    /**
     * Provides date and time utilities.
     */
    var DateTime = (function () {
        function DateTime() {
        }
        /**
         * Gets a new Date that adds the specified number of days to a given Date.
         *
         * @param value Original date.
         * @param days Number of days to add to the given date.
         */
        DateTime.addDays = function (value, days) {
            return new Date(value.getFullYear(), value.getMonth(), value.getDate() + days);
        };
        /**
         * Gets a new Date that adds the specified number of months to a given Date.
         *
         * @param value Original date.
         * @param months Number of months to add to the given date.
         */
        DateTime.addMonths = function (value, months) {
            return new Date(value.getFullYear(), value.getMonth() + months, value.getDate());
        };
        /**
         * Gets a new Date that adds the specified number of years to a given Date.
         *
         * @param value Original date.
         * @param years Number of years to add to the given date.
         */
        DateTime.addYears = function (value, years) {
            return new Date(value.getFullYear() + years, value.getMonth(), value.getDate());
        };
        /**
         * Gets a new Date that adds the specified number of hours to a given Date.
         *
         * @param value Original date.
         * @param hours Number of hours to add to the given date.
         */
        DateTime.addHours = function (value, hours) {
            return new Date(value.getFullYear(), value.getMonth(), value.getDate(), value.getHours() + hours);
        };
        /**
         * Gets a new Date that adds the specified number of minutes to a given Date.
         *
         * @param value Original date.
         * @param minutes Number of minutes to add to the given date.
         */
        DateTime.addMinutes = function (value, minutes) {
            return new Date(value.getFullYear(), value.getMonth(), value.getDate(), value.getHours(), value.getMinutes() + minutes);
        };
        /**
         * Gets a new Date that adds the specified number of seconds to a given Date.
         *
         * @param value Original date.
         * @param seconds Number of seconds to add to the given date.
         */
        DateTime.addSeconds = function (value, seconds) {
            return new Date(value.getFullYear(), value.getMonth(), value.getDate(), value.getHours(), value.getMinutes(), value.getSeconds() + seconds);
        };
        /**
         * Returns true if two Date objects refer to the same date (ignoring time).
         *
         * @param d1 First date.
         * @param d2 Second date.
         */
        DateTime.sameDate = function (d1, d2) {
            return isDate(d1) && isDate(d2) &&
                d1.getFullYear() == d2.getFullYear() &&
                d1.getMonth() == d2.getMonth() &&
                d1.getDate() == d2.getDate();
        };
        /**
         * Returns true if two Date objects refer to the same time (ignoring date).
         *
         * @param d1 First date.
         * @param d2 Second date.
         */
        DateTime.sameTime = function (d1, d2) {
            return isDate(d1) && isDate(d2) &&
                d1.getHours() == d2.getHours() &&
                d1.getMinutes() == d2.getMinutes() &&
                d1.getSeconds() == d2.getSeconds();
        };
        /**
         * Returns true if two Date objects refer to the same date and time.
         *
         * @param d1 First date.
         * @param d2 Second date.
         */
        DateTime.equals = function (d1, d2) {
            return isDate(d1) && isDate(d2) && d1.getTime() == d2.getTime();
        };
        /**
         * Gets a Date object with the date and time set on two Date objects.
         *
         * @param date Date object that contains the date (day/month/year).
         * @param time Date object that contains the time (hour:minute:second).
         */
        DateTime.fromDateTime = function (date, time) {
            if (!date && !time)
                return null;
            if (!date)
                date = time;
            if (!time)
                time = date;
            return new Date(date.getFullYear(), date.getMonth(), date.getDate(), time.getHours(), time.getMinutes(), time.getSeconds());
        };
        /**
         * Converts a calendar date to a fiscal date using the current culture.
         *
         * @param date Calendar date.
         * @param govt Whether to use the government or corporate fiscal year.
         */
        DateTime.toFiscal = function (date, govt) {
            var cal = wijmo.culture.Globalize.calendar;
            return isArray(cal.fiscalYearOffsets)
                ? DateTime.addMonths(date, -cal.fiscalYearOffsets[govt ? 0 : 1])
                : date;
        };
        /**
         * Converts a fiscal year date to a calendar date using the current culture.
         *
         * @param date Fiscal year date.
         * @param govt Whether to use the government or corporate fiscal year.
         */
        DateTime.fromFiscal = function (date, govt) {
            var cal = wijmo.culture.Globalize.calendar;
            return isArray(cal.fiscalYearOffsets)
                ? DateTime.addMonths(date, +cal.fiscalYearOffsets[govt ? 0 : 1])
                : date;
        };
        /**
         * Gets a new Date object set to today's date and no time portion.
         */
        DateTime.newDate = function () {
            var dt = new Date();
            return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
        };
        /**
         * Creates a copy of a given Date object.
         *
         * @param date Date object to copy.
         */
        DateTime.clone = function (date) {
            return DateTime.fromDateTime(date, date);
        };
        return DateTime;
    }());
    wijmo.DateTime = DateTime;
    /**
     * Performs HTTP requests.
     *
     * @param url String containing the URL to which the request is sent.
     * @param settings An optional object used to configure the request.
     *
     * The <b>settings</b> object may contain the following:
     *
     * <table>
     * <tr>
     *   <td><b>method</b></td>
     *   <td>The HTTP method to use for the request (e.g. "POST", "GET", "PUT").
     *       The default is "GET".</td>
     * </tr>
     * <tr>
     *   <td><b>data</b></td>
     *   <td>Data to be sent to the server. It is appended to the url for GET requests,
     *       and converted to a string for other requests.</td>
     * </tr>
     * <tr>
     *   <td><b>async</b></td>
     *   <td>By default, all requests are sent asynchronously (i.e. this is set to true by default).
     *       If you need synchronous requests, set this option to false.</td>
     * </tr>
     * <tr>
     *   <td><b>success</b></td>
     *   <td>A function to be called if the request succeeds.
     *       The function gets passed a single parameter of type <b>XMLHttpRequest</b>.</td>
     * </tr>
     * <tr>
     *   <td><b>error</b></td>
     *   <td>A function to be called if the request fails.
     *       The function gets passed a single parameter of type <b>XMLHttpRequest</b>.</td>
     * </tr>
     * <tr>
     *   <td><b>complete</b></td>
     *   <td>A function to be called when the request finishes (after success and error callbacks are executed).
     *       The function gets passed a single parameter of type <b>XMLHttpRequest</b>.</td>
     * </tr>
     * <tr>
     *   <td><b>beforeSend</b></td>
     *   <td>A function to be called immediately before the request us sent.
     *       The function gets passed a single parameter of type <b>XMLHttpRequest</b>.</td>
     * </tr>
     * <tr>
     *   <td><b>requestHeaders</b></td>
     *   <td>A JavaScript object containing key/value pairs to be added to the request
     *       headers.</td>
     * </tr>
     * <tr>
     *   <td><b>user</b></td>
     *   <td>A username to be used with <b>XMLHttpRequest</b> in response to an HTTP access
     *       authentication request.</td>
     * </tr>
     * <tr>
     *   <td><b>password</b></td>
     *   <td>A password to be used with <b>XMLHttpRequest</b> in response to an HTTP access
     *       authentication request.</td>
     * </tr>
     * </table>
     *
     * Use the <b>success</b> to obtain the result of the request which is provided in
     * the callback's <b>XMLHttpRequest</b> parameter. For example, the code below uses
     * the @see:httpRequest method to retrieve a list of customers from an OData service:
     *
     * <pre>wijmo.httpRequest('http://services.odata.org/Northwind/Northwind.svc/Customers?$format=json', {
     *   success: function (xhr) {
     *     var response = JSON.parse(xhr.response),
     *         customers = response.value;
     *     // do something with the customers...
     *   }
     * });</pre>
     *
     * @return The <b>XMLHttpRequest</b> object used to perform the request.
     */
    function httpRequest(url, settings) {
        if (!settings)
            settings = {};
        // select method and basic options
        var method = settings.method ? asString(settings.method).toUpperCase() : 'GET', async = settings.async != null ? asBoolean(settings.async) : true, data = settings.data;
        // convert data to url parameters for GET requests
        if (data != null && method == 'GET') {
            var s = [];
            for (var k in data) {
                s.push(k + '=' + data[k]);
            }
            if (s.length) {
                var sep = url.indexOf('?') < 0 ? '?' : '&';
                url += sep + s.join('&');
            }
            data = null;
        }
        // create the request
        var xhr = new XMLHttpRequest();
        xhr['URL_DEBUG'] = url; // add some debug info
        // if the data is not a string, stringify it
        var isJson = false;
        if (data != null && !isString(data)) {
            isJson = isObject(data);
            data = JSON.stringify(data);
        }
        // callbacks
        xhr.onload = function () {
            if (xhr.readyState == 4) {
                if (xhr.status < 300) {
                    if (settings.success) {
                        asFunction(settings.success)(xhr);
                    }
                }
                else if (settings.error) {
                    asFunction(settings.error)(xhr);
                }
                if (settings.complete) {
                    asFunction(settings.complete)(xhr);
                }
            }
        };
        xhr.onerror = function () {
            if (settings.error) {
                asFunction(settings.error)(xhr);
            }
            else {
                throw 'HttpRequest Error: ' + xhr.status + ' ' + xhr.statusText;
            }
        };
        // send the request
        xhr.open(method, url, async, settings.user, settings.password);
        if (settings.user && settings.password) {
            xhr.setRequestHeader('Authorization', 'Basic ' + btoa(settings.user + ':' + settings.password));
        }
        if (isJson) {
            xhr.setRequestHeader('Content-Type', 'application/json');
        }
        if (settings.requestHeaders) {
            for (var key in settings.requestHeaders) {
                xhr.setRequestHeader(key, settings.requestHeaders[key]);
            }
        }
        if (isFunction(settings.beforeSend)) {
            settings.beforeSend(xhr);
        }
        xhr.send(data);
        // return the request
        return xhr;
    }
    wijmo.httpRequest = httpRequest;
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Util.js.map
var wijmo;
(function (wijmo) {
    'use strict';
    /**
     * Gets or sets an object that contains all localizable strings in the Wijmo library.
     *
     * The culture selector is a two-letter string that represents an
     * <a href='http://en.wikipedia.org/wiki/List_of_ISO_639-1_codes'>ISO 639 culture</a>.
     */
    wijmo.culture = {
        Globalize: {
            numberFormat: {
                '.': '.',
                ',': ',',
                percent: { pattern: ['-n %', 'n %'] },
                currency: { decimals: 2, symbol: '$', pattern: ['($n)', '$n'] }
            },
            calendar: {
                '/': '/',
                ':': ':',
                firstDay: 0,
                days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
                daysAbbr: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
                months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
                monthsAbbr: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
                am: ['AM', 'A'],
                pm: ['PM', 'P'],
                eras: ['A.D.', 'B.C.'],
                patterns: {
                    d: 'M/d/yyyy', D: 'dddd, MMMM dd, yyyy',
                    f: 'dddd, MMMM dd, yyyy h:mm tt', F: 'dddd, MMMM dd, yyyy h:mm:ss tt',
                    t: 'h:mm tt', T: 'h:mm:ss tt',
                    M: 'MMMM d', m: 'MMMM d',
                    Y: 'MMMM, yyyy', y: 'MMMM, yyyy',
                    g: 'M/d/yyyy h:mm tt', G: 'M/d/yyyy h:mm:ss tt',
                    s: 'yyyy"-"MM"-"dd"T"HH":"mm":"ss',
                    o: 'yyyy"-"MM"-"dd"T"HH":"mm":"ss"."fffffffK',
                    O: 'yyyy"-"MM"-"dd"T"HH":"mm":"ss"."fffffffK',
                    U: 'dddd, MMMM dd, yyyy h:mm:ss tt'
                },
                fiscalYearOffsets: [-3, -3]
            }
        }
    };
    /**
     * Class that implements formatting and parsing of numbers and Dates.
     *
     * By default, @see:Globalize uses the American English culture.
     * To switch cultures, include the appropriate <b>wijmo.culture.*.js</b>
     * file after the wijmo files.
     */
    var Globalize = (function () {
        function Globalize() {
        }
        /**
         * Formats a number or a date.
         *
         * The format strings used with the @see:format function are similar to
         * the ones used by <b>Globalize.js</b> and by the .NET Globalization
         * library. The tables below contains links that describe the formats
         * available:
         *
         * <ul>
         * <li><a href="http://msdn.microsoft.com/en-us/library/dwhawy9k(v=vs.110).aspx">
         *      Standard Numeric Format Strings</a></li>
         * <li><a href="http://msdn.microsoft.com/en-us/library/az4se3k1(v=vs.110).aspx">
         *      Standard Date and Time Format Strings</a></li>
         * <li><a href="http://msdn.microsoft.com/en-us/library/8kb3ddd4(v=vs.110).aspx">
         *      Custom Date and Time Format Strings</a></li>
         * </ul>
         *
         * @param value Number or Date to format (all other types are converted to strings).
         * @param format Format string to use when formatting numbers or dates.
         * @param trim Whether to remove trailing zeros from numeric results.
         * @param truncate Whether to truncate the numeric values rather than round them.
         * @return A string representation of the given value.
         */
        Globalize.format = function (value, format, trim, truncate) {
            // format numbers and dates, convert others to string
            if (wijmo.isString(value)) {
                return value;
            }
            else if (wijmo.isNumber(value)) {
                format = format || (value == Math.round(value) ? 'n0' : 'n2');
                return Globalize.formatNumber(value, format, trim, truncate);
            }
            else if (wijmo.isDate(value)) {
                format = format || 'd';
                return Globalize.formatDate(value, format);
            }
            else {
                return value != null ? value.toString() : '';
            }
        };
        /**
         * Formats a number using the current culture.
         *
         * The @see:formatNumber method accepts most .NET-style
         * <a href="http://msdn.microsoft.com/en-us/library/dwhawy9k(v=vs.110).aspx">
         * Standard Numeric Format Strings</a>, except for the 'e' and 'x' formats
         * (scientific notation and hexadecimal) which are not supported.
         *
         * Numeric format strings takes the form <i>Axxccss</i>, where:
         * <ul>
         * <li>
         *  <i>A</i> is a single case-insensitive alphabetic character called the
         *  format specifier.</li>
         * <li>
         *  <i>xx</i> is an optional integer called the precision specifier.
         *  The precision specifier affects the number of digits in the result.</li>
         * <li>
         *  <i>cc</i> is an optional string used to override the currency symbol
         *  when formatting currency values. This is useful when formatting
         *  currency values for cultures different than the current default
         *  (for example, when formatting Euro or Yen values in applications
         *  that use the English culture).</li>
         * <li>
         *  <i>ss</i> is an optional string used to scale the number. If provided,
         *  it must consist of commas. The number is divided by 1000 for each comma
         *  specified.</li>
         * </ul>
         *
         * The following table describes the standard numeric format specifiers and
         * displays sample output produced by each format specifier for the default
         * culture.
         *
         * <b>n</b> Number: <code>formatNumber(1234.5, 'n2') => '1,234.50'</code><br/>
         * <b>f</b> Fixed-point: <code>formatNumber(1234.5, 'f2') => '1234.50'</code><br/>
         * <b>g</b> General (no trailing zeros): <code>formatNumber(1234.5, 'g2') => '1234.5'</code><br/>
         * <b>d</b> Decimal (integers): <code>formatNumber(-1234, 'd6') => '-001234'</code><br/>
         * <b>x</b> Hexadecimal (integers): <code>formatNumber(1234, 'x6') => '0004d2'</code><br/>
         * <b>c</b> Currency: <code>formatNumber(1234, 'c') => '$ 1,234.00'</code><br/>
         * <b>p</b> Percent: <code>formatNumber(0.1234, 'p2') => '12.34 %'</code>
         *
         * The scaling specifier is especially useful when charting large values. For
         * example, the markup below creates a chart that plots population versus GDP.
         * The raw data expresses the population is units and the GDP in millions.
         * The scaling specified in the axes formats causes the chart to show population
         * in millions and GDP in trillions:
         *
         * <pre>&lt;wj-flex-chart
         *   items-source="countriesGDP" binding-x="pop" chart-type="Scatter"&gt;
         *   &lt;wj-flex-chart-series
         *     name="GDP" binding="gdp"&gt;&lt;/wj-flex-chart-series&gt;
         *   &lt;wj-flex-chart-axis
         *     wj-property="axisX" title="Population (millions)"
         *     format="n0,,"&gt;
         *   &lt;/wj-flex-chart-axis&gt;
         *   &lt;wj-flex-chart-axis
         *     wj-property="axisY" title="GDP (US$ trillions)"
         *     format="c0,,"&gt;
         *   &lt;/wj-flex-chart-axis&gt;
         * &lt;/wj-flex-chart&gt;</pre>
         *
         * @param value Number to format.
         * @param format .NET-style standard numeric format string (e.g. 'n2', 'c4', 'p0', 'g2', 'd2').
         * @param trim Whether to remove trailing zeros from the result.
         * @param truncate Whether to truncate the value rather than round it.
         * @return A string representation of the given number.
         */
        Globalize.formatNumber = function (value, format, trim, truncate) {
            wijmo.asNumber(value);
            wijmo.asString(format);
            var result, m = format ? format.match(/([a-z])(\d*)(,*)(.*)/i) : null, nf = wijmo.culture.Globalize.numberFormat, f1 = m ? m[1].toLowerCase() : 'n', prec = (m && m[2]) ? parseInt(m[2]) : (f1 == 'c') ? nf.currency.decimals : value == Math.round(value) ? 0 : 2, scale = (m && m[3]) ? 3 * m[3].length : 0, curr = (m && m[4]) ? m[4] : nf.currency.symbol, dp = nf['.'], ts = nf[','];
            // scale (,:thousands ,,:millions ,,,:billions)
            if (scale) {
                value /= Math.pow(10, scale);
            }
            // d, x: integers/hexadecimal
            if (f1 == 'd' || f1 == 'x') {
                result = Math.round(Math.abs(value)).toString(f1 == 'd' ? 10 : 16);
                while (result.length < prec) {
                    result = '0' + result;
                }
                if (value < 0) {
                    result = '-' + result;
                }
                if (format && format[0] == 'X') {
                    result = result.toUpperCase();
                }
                return result;
            }
            // p: percentage
            if (f1 == 'p') {
                value = wijmo.toFixed(value * 100, prec, truncate); // TFS 147585
            }
            // truncate value
            if (truncate) {
                value = wijmo.toFixed(value, prec, true);
            }
            // get result
            result = (f1 == 'c' || f1 == 'p')
                ? Math.abs(value).toFixed(prec)
                : value.toFixed(prec);
            // g: remove trailing zeros
            if ((trim || f1 == 'g') && result.indexOf('.') > -1) {
                result = result.replace(/(\.[0-9]*?)0+$/g, '$1');
                result = result.replace(/\.$/, '');
            }
            // replace decimal point
            if (dp != '.') {
                result = result.replace('.', dp);
            }
            // n, c, p: thousand separators
            if (ts && (f1 == 'n' || f1 == 'c' || f1 == 'p')) {
                var idx = result.indexOf(dp), rx = /\B(?=(\d\d\d)+(?!\d))/g;
                result = idx > -1 ? result.substr(0, idx).replace(rx, ts) + result.substr(idx) : result.replace(rx, ts);
            }
            // c: currency pattern
            if (f1 == 'c') {
                var pat = nf.currency.pattern[value < 0 ? 0 : 1];
                result = pat.replace('n', result).replace('$', curr);
            }
            // p: percentage pattern
            if (f1 == 'p') {
                var pat = nf.percent.pattern[value < 0 ? 0 : 1];
                result = pat.replace('n', result);
            }
            // done
            return result;
        };
        /**
         * Formats a date using the current culture.
         *
         * The @see:format parameter contains a .NET-style
         * <a href="http://msdn.microsoft.com/en-us/library/8kb3ddd4(v=vs.110).aspx">Date format string</a>
         * with the following additions:
         * <ul>
         *  <li><i>Q, q</i> Calendar quarter.</li>
         *  <li><i>U</i> Fiscal quarter (government).</li>
         *  <li><i>u</i> Fiscal quarter (private sector).</li>
         *  <li><i>EEEE, EEE, EE, E</i> Fiscal year (government).</li>
         *  <li><i>eeee, eee, ee, e</i> Fiscal year (private sector).</li>
         * </ul>
         *
         * For example:
         *
         * <pre>
         * var d = new Date(2015, 9, 1); // Oct 1, 2015
         * console.log(wijmo.Globalize.format(d, '"FY"EEEE"Q"U') + ' (US culture)');
         * &gt; FY2016Q1 (US culture)
         * </pre>
         *
         * @param value Number or Date to format.
         * @param format .NET-style Date format string.
         * @return A string representation of the given date.
         */
        Globalize.formatDate = function (value, format) {
            value = wijmo.asDate(value);
            // culture-invariant formats
            switch (format) {
                case 'r':
                case 'R':
                    return value.toUTCString();
                case 'u':
                    return value.toISOString().replace(/\.\d{3}/, '');
            }
            // expand pre-defined formats
            format = Globalize._expandFormat(format);
            // parse the format string and build return value
            var parts = Globalize._parseDateFormat(format), str = '';
            for (var i = 0; i < parts.length; i++) {
                str += Globalize._formatDatePart(value, format, parts[i]);
            }
            // all done
            return str;
        };
        /**
         * Parses a string into an integer.
         *
         * @param value String to convert to an integer.
         * @param format Format to use when parsing the number.
         * @return The integer represented by the given string,
         * or <b>NaN</b> if the string cannot be parsed into an integer.
         */
        Globalize.parseInt = function (value, format) {
            return Math.round(Globalize.parseFloat(value, format));
        };
        /**
         * Parses a string into a floating point number.
         *
         * @param value String to convert to a number.
         * @param format Format to use when parsing the number.
         * @return The floating point number represented by the given string,
         * or <b>NaN</b> if the string cannot be parsed into a floating point number.
         */
        Globalize.parseFloat = function (value, format) {
            var neg = value.indexOf('-') > -1 || (value.indexOf('(') > -1 && value.indexOf(')') > -1) ? -1 : +1, pct = value.indexOf('%') > -1 ? .01 : 1, m = format ? format.match(/,+/) : null, scale = m ? m[0].length * 3 : 0;
            // hex
            if (format && (format[0] == 'x' || format[0] == 'X')) {
                value = value.replace(/[^0-9a-f]+.*$/gi, ''); // truncate at first invalid char
                return parseInt(value, 16) * neg * pct * Math.pow(10, scale);
            }
            // decimal
            var dp = wijmo.culture.Globalize.numberFormat['.'], rx = new RegExp('[^\\d\\' + dp + ']', 'g'), val = value.replace(rx, '').replace(dp, '.'); // remove non-digits, replace decimal point
            return parseFloat(val) * neg * pct * Math.pow(10, scale);
        };
        /**
         * Parses a string into a Date.
         *
         * @param value String to convert to a Date.
         * @param format Format string used to parse the date.
         * @return The date represented by the given string, or null if the string
         * cannot be parsed into a Date.
         */
        Globalize.parseDate = function (value, format) {
            // make sure we have a value
            value = wijmo.asString(value);
            if (!value) {
                return null;
            }
            // culture-invariant formats
            if (format == 'u') {
                return new Date(value);
            }
            // parse using RFC 3339 pattern ([yyyy-MM-dd] [hh:mm[:ss]])
            var d;
            if (format == 'R' || format == 'r') {
                var rx = /(([0-9]+)\-([0-9]+)\-([0-9]+))?\s?(([0-9]+):([0-9]+)(:([0-9]+))?)?/, match = value.match(rx);
                if (match[1] || match[5]) {
                    var d = match[1] // parse date
                        ? new Date(parseInt(match[2]), parseInt(match[3]) - 1, parseInt(match[4]))
                        : new Date();
                    if (match[5]) {
                        d.setHours(parseInt(match[6]));
                        d.setMinutes(parseInt(match[7]));
                        d.setSeconds(match[8] ? parseInt(match[9]) : 0);
                    }
                }
                else {
                    d = new Date(value);
                }
                return !isNaN(d.getTime()) ? d : null;
            }
            // expand the format
            format = Globalize._expandFormat(format ? format : 'd');
            // get format parts and data parts
            //
            // cjk: chars, http://www.rikai.com/library/kanjitables/kanji_codes.unicode.shtml
            // rxf: format (no dots in strings: 'mm.dd.yyyy' => ['mm', 'dd', 'yyyy']).
            // rxv: value (dots OK in strings: 'A.D' => 'A.D', but not by themselves)
            var cal = wijmo.culture.Globalize.calendar, cjk = Globalize._CJK, rxv = new RegExp('(\\' + cal['/'] + ')|(\\' + cal[':'] + ')|' +
                '(\\d+)|' +
                '([' + cjk + '\\.]{2,})|' +
                '([' + cjk + ']+)', // strings with no dots
            'gi'), vparts = value.match(rxv), fparts = Globalize._parseDateFormat(format), offset = 0, year = -1, month = 0, day = 1, hour = 0, min = 0, tzm = 0, sec = 0, ms = 0, era = -1, hasDayName, hasDay, hasQuarter, hasMonth, fiscalFmt;
            // basic validation (TFS 81465, 128359)
            if (!vparts || !vparts.length || !fparts || !fparts.length) {
                return null;
            }
            // parse each element
            for (var i = 0; i < fparts.length && vparts; i++) {
                var vpi = i - offset, pval = (vpi > -1 && vpi < vparts.length) ? vparts[vpi] : '', plen = fparts[i].length;
                switch (fparts[i]) {
                    // ** year
                    case 'EEEE':
                    case 'EEE':
                    case 'EE':
                    case 'E': // fiscal (govt)
                    case 'eeee':
                    case 'eee':
                    case 'ee':
                    case 'e':
                        fiscalFmt = fparts[i];
                    // ** fall through **
                    case 'yyyy':
                    case 'yyy':
                    case 'yy':
                    case 'y':
                        if (plen > 1 && pval.length > plen) {
                            vparts[vpi] = pval.substr(plen);
                            pval = pval.substr(0, plen);
                            offset++;
                        }
                        year = parseInt(pval);
                        break;
                    // ** month
                    case 'MMMM':
                    case 'MMM':
                        hasMonth = true;
                        var monthName = pval.toLowerCase();
                        month = -1;
                        for (var j = 0; j < 12; j++) {
                            if (cal.months[j].toLowerCase().indexOf(monthName) == 0) {
                                month = j;
                                break;
                            }
                        }
                        if (month > -1) {
                            break;
                        }
                    // FALL THROUGH: 
                    // and try parsing month as a number
                    // so users can type "1/2" instead of "1/February"
                    case 'MM':
                    case 'M':
                        hasMonth = true;
                        if (plen > 1 && pval.length > plen) {
                            vparts[vpi] = pval.substr(plen);
                            pval = pval.substr(0, plen);
                            offset++;
                        }
                        month = parseInt(pval) - 1;
                        break;
                    // ** day
                    case 'dddd':
                    case 'ddd':
                        hasDayName = true;
                        break; // skip day names
                    case 'dd':
                    case 'd':
                        if (plen > 1 && pval.length > plen) {
                            vparts[vpi] = pval.substr(plen);
                            pval = pval.substr(0, plen);
                            offset++;
                        }
                        day = parseInt(pval);
                        hasDay = true;
                        break;
                    // ** hour
                    case 'hh':
                    case 'h':
                        if (plen > 1 && pval.length > plen) {
                            vparts[vpi] = pval.substr(plen);
                            pval = pval.substr(0, plen);
                            offset++;
                        }
                        hour = parseInt(pval);
                        hour = hour == 12 ? 0 : hour; // 0-12, 12 == midnight
                        break;
                    case 'HH':
                        if (plen > 1 && pval.length > plen) {
                            vparts[vpi] = pval.substr(plen);
                            pval = pval.substr(0, plen);
                            offset++;
                        }
                        hour = parseInt(pval); // 0-24
                        break;
                    case 'H':
                        hour = parseInt(pval); // 0-24
                        break;
                    // ** minute
                    case 'mm':
                    case 'm':
                        if (plen > 1 && pval.length > plen) {
                            vparts[vpi] = pval.substr(plen);
                            pval = pval.substr(0, plen);
                            offset++;
                        }
                        min = parseInt(pval);
                        break;
                    // ** second
                    case 'ss':
                    case 's':
                        if (plen > 1 && pval.length > plen) {
                            vparts[vpi] = pval.substr(plen);
                            pval = pval.substr(0, plen);
                            offset++;
                        }
                        sec = parseInt(pval);
                        break;
                    // ** millisecond
                    case 'fffffff':
                    case 'FFFFFFF':
                    case 'ffffff':
                    case 'FFFFFF':
                    case 'fffff':
                    case 'FFFFF':
                    case 'ffff':
                    case 'FFFF':
                    case 'fff':
                    case 'FFF':
                    case 'ff':
                    case 'FF':
                    case 'f':
                    case 'F':
                        ms = parseInt(pval) / Math.pow(10, plen - 3);
                        break;
                    // ** am/pm
                    case 'tt':
                    case 't':
                        pval = pval.toUpperCase();
                        if ((cal.pm[0] && pval == cal.pm[0] && hour < 12) ||
                            (cal.pm[1] && pval == cal.pm[1] && hour < 12)) {
                            hour += 12;
                        }
                        break;
                    // ** quarter
                    case 'q':
                    case 'Q':
                    case 'u':
                    case 'U':
                        hasQuarter = true;
                        break;
                    // ** era
                    case 'ggg':
                    case 'gg':
                    case 'g':
                        era = cal.eras.length > 1 ? Globalize._getEra(pval, cal) : -1;
                        break;
                    // ** localized separators (TFS 131320)
                    case cal['/']:
                    case cal[':']:
                        if (pval && pval != fparts[i]) {
                            return null; // present and wrong separator
                        }
                        break;
                    // ** time zone (skip )
                    case 'K':
                        break;
                    // ** all else: if not a match, keep using the same pval
                    default:
                        if (Globalize._unquote(fparts[i]) != pval) {
                            offset++;
                        }
                        break;
                }
            }
            // allow dates with no times even if the format requires times
            if (hasMonth && hasDay) {
                if (isNaN(hour))
                    hour = 0;
                if (isNaN(min))
                    min = 0;
                if (isNaN(sec))
                    sec = 0;
            }
            // basic validation
            if (month < 0 || month > 11 || isNaN(month) ||
                day < 0 || day > 31 || isNaN(day) ||
                hour < 0 || hour > 24 || isNaN(hour) ||
                min < 0 || min > 60 || isNaN(min) ||
                sec < 0 || sec > 60 || isNaN(sec)) {
                return null;
            }
            // convert fiscal year/month to calendar
            if (fiscalFmt) {
                if (!hasMonth) {
                    return null;
                }
                d = new Date(year, month);
                d = wijmo.DateTime.fromFiscal(d, fiscalFmt[0] == 'E');
                year = d.getFullYear();
                month = d.getMonth();
            }
            // if the day name was specified but the day wasn't, the result is meaningless
            if (hasDayName && !hasDay) {
                return null;
            }
            // if the quarter was specified but the month wasn't, the result is meaningless
            if (hasQuarter && !hasMonth) {
                return null;
            }
            // if year not found, use current (as Globalize.js)
            if (year < 0) {
                year = new Date().getFullYear();
            }
            // apply era offset if any, or adjust for two-digit years (see Calendar.TwoDigitYearMax)
            if (era > -1) {
                year = year + cal.eras[era].start.getFullYear() - 1;
            }
            else if (year < 100) {
                year += year >= 30 ? 1900 : 2000;
            }
            // return result
            d = new Date(year, month, day, hour, min + tzm, sec, ms);
            return isNaN(d.getTime()) ? null : d;
        };
        /**
         * Gets the first day of the week according to the current culture.
         *
         * The value returned is between zero (Sunday) and six (Saturday).
         */
        Globalize.getFirstDayOfWeek = function () {
            var fdw = wijmo.culture.Globalize.calendar.firstDay;
            return fdw ? fdw : 0;
        };
        /**
         * Gets the symbol used as a decimal separator in numbers.
         */
        Globalize.getNumberDecimalSeparator = function () {
            var ndc = wijmo.culture.Globalize.numberFormat['.'];
            return ndc ? ndc : '.';
        };
        // ** implementation
        // unquotes a string
        Globalize._unquote = function (s) {
            if (s.length > 1 && s[0] == s[s.length - 1]) {
                if (s[0] == '\'' || s[0] == '\"') {
                    return s.substr(1, s.length - 2);
                }
            }
            return s;
        };
        Globalize._parseDateFormat = function (format) {
            // use cache whenever possible
            if (format in Globalize._dateFomatParts) {
                return Globalize._dateFomatParts[format];
            }
            // parse the format
            var parts = [], str = '', start, end;
            for (start = 0; start > -1 && start < format.length; start++) {
                var c = format[start];
                if (c == '\'' || c == '"') {
                    end = format.indexOf(c, start + 1); // keep quotes to distinguish from regular date parts
                    if (end > -1) {
                        parts.push(format.substring(start, end + 1));
                        start = end;
                        continue;
                    }
                }
                end = start + 1;
                for (; end < format.length; end++) {
                    if (format[end] != c)
                        break;
                }
                parts.push(format.substring(start, end));
                start = end - 1;
            }
            // cache and return
            Globalize._dateFomatParts[format] = parts;
            return parts;
        };
        // format a date part into a string
        Globalize._formatDatePart = function (d, format, part) {
            var cal = wijmo.culture.Globalize.calendar, era = 0, year = 0, ff = 0, fd, plen = part.length;
            switch (part) {
                // ** year
                case 'yyyy':
                case 'yyy':
                case 'yy':
                case 'y': // calendar year
                case 'EEEE':
                case 'EEE':
                case 'EE':
                case 'E': // fiscal year (govt)
                case 'eeee':
                case 'eee':
                case 'ee':
                case 'e':
                    // get the year (calendar or fiscal)
                    fd = part[0] == 'E' ? wijmo.DateTime.toFiscal(d, true) :
                        part[0] == 'e' ? wijmo.DateTime.toFiscal(d, false) :
                            d;
                    year = fd.getFullYear();
                    // if the calendar has multiple eras and the format specifies an era,
                    // then adjust the year to count from the start of the era.
                    // if the format has no era, then use the regular (Western) year.
                    if (cal.eras.length > 1 && format.indexOf('g') > -1) {
                        era = Globalize._getEra(d, cal);
                        if (era > -1) {
                            year = year - cal.eras[era].start.getFullYear() + 1;
                        }
                    }
                    // adjust number of digits
                    return Globalize._zeroPad(year, 4).substr(4 - part.length);
                // ** month
                case 'MMMM':
                    return cal.months[d.getMonth()];
                case 'MMM':
                    return cal.monthsAbbr[d.getMonth()];
                case 'MM':
                case 'M':
                    return Globalize._zeroPad(d.getMonth() + 1, plen);
                // ** day
                case 'dddd':
                    return cal.days[d.getDay()];
                case 'ddd':
                    return cal.daysAbbr[d.getDay()];
                case 'dd':
                    return Globalize._zeroPad(d.getDate(), 2);
                case 'd':
                    return d.getDate().toString();
                // ** hour
                case 'hh':
                case 'h':
                    return Globalize._zeroPad(Globalize._h12(d), plen);
                case 'HH':
                case 'H':
                    return Globalize._zeroPad(d.getHours(), plen);
                // ** minute
                case 'mm':
                case 'm':
                    return Globalize._zeroPad(d.getMinutes(), plen);
                // ** second
                case 'ss':
                case 's':
                    return Globalize._zeroPad(d.getSeconds(), plen);
                // ** millisecond
                case 'fffffff':
                case 'FFFFFFF':
                case 'ffffff':
                case 'FFFFFF':
                case 'fffff':
                case 'FFFFF':
                case 'ffff':
                case 'FFFF':
                case 'fff':
                case 'FFF':
                case 'ff':
                case 'FF':
                case 'f':
                case 'F':
                    ff = d.getMilliseconds() * Math.pow(10, plen - 3);
                    return part[0] == 'f' ? Globalize._zeroPad(ff, plen) : ff.toFixed(0);
                // ** am/pm
                case 'tt':
                    return d.getHours() < 12 ? cal.am[0] : cal.pm[0];
                case 't':
                    return d.getHours() < 12 ? cal.am[1] : cal.pm[1];
                // ** quarter
                case 'q':
                case 'Q':
                    return (Math.floor(d.getMonth() / 3) + 1).toString();
                case 'u':
                case 'U':
                    fd = wijmo.DateTime.toFiscal(d, part == 'U');
                    return (Math.floor(fd.getMonth() / 3) + 1).toString();
                // ** era
                case 'ggg':
                case 'gg':
                case 'g':
                    if (cal.eras.length > 1) {
                        era = Globalize._getEra(d, cal);
                        if (era > -1) {
                            return part == 'ggg' ? cal.eras[era].name : part == 'gg' ? cal.eras[era].name[0] : cal.eras[era].symbol;
                        }
                    }
                    return cal.eras[0];
                // ** localized separators
                case ':':
                case '/':
                    return cal[part];
                // ** time zone
                case 'K':
                    var tz = d.toString().match(/(\+|\-)(\d{2})(\d{2})/);
                    return tz ? tz[1] + tz[2] + tz[3] : '';
            }
            // unquote part
            if (plen > 1 && part[0] == part[plen - 1]) {
                if (part[0] == '\"' || part[0] == '\'') {
                    return part.substr(1, plen - 2);
                }
            }
            // return part
            return part;
        };
        // get a date's era (used only in Japanese locales)
        Globalize._getEra = function (d, cal) {
            if (wijmo.isDate(d)) {
                for (var i = 0; i < cal.eras.length; i++) {
                    if (d >= cal.eras[i].start) {
                        return i;
                    }
                }
            }
            else if (wijmo.isString(d)) {
                for (var i = 0; i < cal.eras.length; i++) {
                    if (cal.eras[i].name) {
                        if (cal.eras[i].name.indexOf(d) == 0 || cal.eras[i].symbol.indexOf(d) == 0) {
                            return i;
                        }
                    }
                }
            }
            return -1; // not found
        };
        // expand date pattern into full date format
        Globalize._expandFormat = function (format) {
            var fmt = wijmo.culture.Globalize.calendar.patterns[format];
            return fmt ? fmt : format;
        };
        // format a number with leading zeros
        Globalize._zeroPad = function (num, places) {
            var n = num.toFixed(0), zero = places - n.length + 1;
            return zero > 0 ? Array(zero).join('0') + n : n;
        };
        // format an hour to 12 or 24 hour base depending on the calendar
        Globalize._h12 = function (d) {
            var cal = wijmo.culture.Globalize.calendar, h = d.getHours();
            if (cal.am && cal.am[0]) {
                h = h % 12;
                if (h == 0)
                    h = 12;
            }
            return h;
        };
        // Chinese/Japanese/Korean characters
        // http://www.rikai.com/library/kanjitables/kanji_codes.unicode.shtml
        // http://www.programminginkorean.com/programming/hangul-in-unicode/
        // NOTE: using 'replace' to keep minifier from switching the escaped Unicode chars into real Unicode.
        // NOTE: will have to expand for other cultures? (Arabic/Hebrew/Hindi etc?)
        Globalize._CJK = 'a-z' +
            'u00c0-u017fu3000-u30ffu4e00-u9faf'.replace(/u/g, '\\u') +
            'u1100-u11ffu3130-u318fua960-ua97fuac00-ud7afud7b0-ud7ff'.replace(/u/g, '\\u'); // k
        // parse a date format string into its parts
        Globalize._dateFomatParts = {};
        return Globalize;
    }());
    wijmo.Globalize = Globalize;
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Globalize.js.map
var wijmo;
(function (wijmo) {
    'use strict';
    /**
     * Provides binding to complex properties (e.g. 'customer.address.city')
     */
    var Binding = (function () {
        /**
         * Initializes a new instance of the @see:Binding class.
         *
         * @param path Name of the property to bind to.
         */
        function Binding(path) {
            this.path = path;
        }
        Object.defineProperty(Binding.prototype, "path", {
            /**
             * Gets or sets the path for the binding.
             *
             * In the simplest case, the path is the name of the property of the source
             * object to use for the binding (e.g. 'street').
             *
             * Sub-properties of a property can be specified by a syntax similar to that
             * used in JavaScript (e.g. 'address.street').
             */
            get: function () {
                return this._path;
            },
            set: function (value) {
                this._path = value;
                this._parts = value ? value.split('.') : []; // e.g. 'customer.balance'
                for (var i = 0; i < this._parts.length; i++) {
                    var part = this._parts[i], ib = part.indexOf('['); // e.g. 'customer.balance[0]'
                    if (ib > -1) {
                        this._parts[i] = part.substr(0, ib);
                        this._parts.splice(++i, 0, parseInt(part.substr(ib + 1)));
                    }
                }
                this._key = this._parts.length == 1 ? this._parts[0] : null;
            },
            enumerable: true,
            configurable: true
        });
        /**
         * Gets the binding value for a given object.
         *
         * If the object does not contain the property specified by the
         * binding @see:path, the method returns null.
         *
         * @param object The object that contains the data to be retrieved.
         */
        Binding.prototype.getValue = function (object) {
            if (object) {
                // optimize common case
                if (this._key) {
                    return object[this._key];
                }
                // handle case where property name has a decimal point (TFS 139176)
                if (this._path in object) {
                    return object[this._path];
                }
                // traverse path for complex properties
                for (var i = 0; i < this._parts.length && object; i++) {
                    object = object[this._parts[i]];
                }
            }
            return object;
        };
        /**
         * Sets the binding value on a given object.
         *
         * If the object does not contain the property specified by the
         * binding @see:path, the value is not set.
         *
         * @param object The object that contains the data to be set.
         * @param value Data value to set.
         */
        Binding.prototype.setValue = function (object, value) {
            if (object) {
                // handle simple cases (and cases where the property name has a decimal point)
                if (this._path in object) {
                    object[this._path] = value;
                    return;
                }
                // traverse parts for complex properties
                for (var i = 0; i < this._parts.length - 1; i++) {
                    object = object[this._parts[i]];
                    if (object == null) {
                        return;
                    }
                }
                // make the assignment
                object[this._parts[this._parts.length - 1]] = value;
            }
        };
        return Binding;
    }());
    wijmo.Binding = Binding;
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Binding.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    'use strict';
    /*
     * Represents an event handler (private class)
     */
    var EventHandler = (function () {
        function EventHandler(handler, self) {
            this.handler = handler;
            this.self = self;
        }
        return EventHandler;
    }());
    /**
     * Represents an event.
     *
     * Wijmo events are similar to .NET events. Any class may define events by
     * declaring them as fields. Any class may subscribe to events using the
     * event's @see:addHandler method and unsubscribe using the @see:removeHandler
     * method.
     *
     * Wijmo event handlers take two parameters: <i>sender</i> and <i>args</i>.
     * The first is the object that raised the event, and the second is an object
     * that contains the event parameters.
     *
     * Classes that define events follow the .NET pattern where for every event
     * there is an <i>on[EVENTNAME]</i> method that raises the event. This pattern
     * allows derived classes to override the <i>on[EVENTNAME]</i> method and
     * handle the event before and/or after the base class raises the event.
     * Derived classes may even suppress the event by not calling the base class
     * implementation.
     *
     * For example, the TypeScript code below overrides the <b>onValueChanged</b>
     * event for a control to perform some processing before and after the
     * <b>valueChanged</b> event fires:
     *
     * <pre>// override base class
     * onValueChanged(e: EventArgs) {
     *   // execute some code before the event fires
     *   console.log('about to fire valueChanged');
     *   // optionally, call base class to fire the event
     *   super.onValueChanged(e);
     *   // execute some code after the event fired
     *   console.log('valueChanged event just fired');
     * }</pre>
     */
    var Event = (function () {
        function Event() {
            this._handlers = [];
        }
        /**
         * Adds a handler to this event.
         *
         * @param handler Function invoked when the event is raised.
         * @param self Object that defines the event handler
         * (accessible as 'this' from the handler code).
         */
        Event.prototype.addHandler = function (handler, self) {
            wijmo.asFunction(handler);
            this._handlers.push(new EventHandler(handler, self));
        };
        /**
         * Removes a handler from this event.
         *
         * @param handler Function invoked when the event is raised.
         * @param self Object that defines the event handler (accessible as 'this' from the handler code).
         */
        Event.prototype.removeHandler = function (handler, self) {
            wijmo.asFunction(handler);
            for (var i = 0; i < this._handlers.length; i++) {
                var l = this._handlers[i];
                if (l.handler == handler || handler == null) {
                    if (l.self == self || self == null) {
                        this._handlers.splice(i, 1);
                        if (handler && self) {
                            break;
                        }
                    }
                }
            }
        };
        /**
         * Removes all handlers associated with this event.
         */
        Event.prototype.removeAllHandlers = function () {
            this._handlers.length = 0;
        };
        /**
         * Raises this event, causing all associated handlers to be invoked.
         *
         * @param sender Source object.
         * @param args Event parameters.
         */
        Event.prototype.raise = function (sender, args) {
            if (args === void 0) { args = EventArgs.empty; }
            for (var i = 0; i < this._handlers.length; i++) {
                var l = this._handlers[i];
                l.handler.call(l.self, sender, args);
            }
        };
        Object.defineProperty(Event.prototype, "hasHandlers", {
            /**
             * Gets a value that indicates whether this event has any handlers.
             */
            get: function () {
                return this._handlers.length > 0;
            },
            enumerable: true,
            configurable: true
        });
        return Event;
    }());
    wijmo.Event = Event;
    /**
     * Base class for event arguments.
     */
    var EventArgs = (function () {
        function EventArgs() {
        }
        /**
         * Provides a value to use with events that do not have event data.
         */
        EventArgs.empty = new EventArgs();
        return EventArgs;
    }());
    wijmo.EventArgs = EventArgs;
    /**
     * Provides arguments for cancellable events.
     */
    var CancelEventArgs = (function (_super) {
        __extends(CancelEventArgs, _super);
        function CancelEventArgs() {
            _super.apply(this, arguments);
            /**
             * Gets or sets a value that indicates whether the event should be canceled.
             */
            this.cancel = false;
        }
        return CancelEventArgs;
    }(EventArgs));
    wijmo.CancelEventArgs = CancelEventArgs;
    /**
     * Provides arguments for property change events.
     */
    var PropertyChangedEventArgs = (function (_super) {
        __extends(PropertyChangedEventArgs, _super);
        /**
         * Initializes a new instance of the @see:PropertyChangedEventArgs class.
         *
         * @param propertyName The name of the property whose value changed.
         * @param oldValue The old value of the property.
         * @param newValue The new value of the property.
         */
        function PropertyChangedEventArgs(propertyName, oldValue, newValue) {
            _super.call(this);
            this._name = propertyName;
            this._oldVal = oldValue;
            this._newVal = newValue;
        }
        Object.defineProperty(PropertyChangedEventArgs.prototype, "propertyName", {
            /**
             * Gets the name of the property whose value changed.
             */
            get: function () {
                return this._name;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PropertyChangedEventArgs.prototype, "oldValue", {
            /**
             * Gets the old value of the property.
             */
            get: function () {
                return this._oldVal;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PropertyChangedEventArgs.prototype, "newValue", {
            /**
             * Gets the new value of the property.
             */
            get: function () {
                return this._newVal;
            },
            enumerable: true,
            configurable: true
        });
        return PropertyChangedEventArgs;
    }(EventArgs));
    wijmo.PropertyChangedEventArgs = PropertyChangedEventArgs;
    /**
     * Provides arguments for @see:XMLHttpRequest error events.
     */
    var RequestErrorEventArgs = (function (_super) {
        __extends(RequestErrorEventArgs, _super);
        /**
         * Initializes a new instance of the @see:RequestErrorEventArgs class.
         *
         * @param xhr The @see:XMLHttpRequest that detected the error.
         * The status and statusText properties of the request object
         * contain details about the error.
         */
        function RequestErrorEventArgs(xhr) {
            _super.call(this);
            this._xhr = xhr;
        }
        Object.defineProperty(RequestErrorEventArgs.prototype, "request", {
            /**
             * Gets a reference to the @see:XMLHttpRequest that detected the error.
             *
             * The status and statusText properties of the request object contain
             * details about the error.
             */
            get: function () {
                return this._xhr;
            },
            enumerable: true,
            configurable: true
        });
        return RequestErrorEventArgs;
    }(CancelEventArgs));
    wijmo.RequestErrorEventArgs = RequestErrorEventArgs;
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Event.js.map
var wijmo;
(function (wijmo) {
    'use strict';
    /**
     * Base class for all Wijmo controls.
     *
     * The @see:Control class handles the association between DOM elements and the
     * actual control. Use the @see:hostElement property to get the DOM element
     * that is hosting a control, or the @see:getControl method to get the control
     * hosted in a given DOM element.
     *
     * The @see:Control class also provides a common pattern for invalidating and
     * refreshing controls, for updating the control layout when its size changes,
     * and for handling the HTML templates that define the control structure.
     */
    var Control = (function () {
        /**
         * Initializes a new instance of the @see:Control class and attaches it to a DOM element.
         *
         * @param element The DOM element that will host the control, or a selector for the host element (e.g. '#theCtrl').
         * @param options JavaScript object containing initialization data for the control.
         * @param invalidateOnResize Whether the control should be invalidated when it is resized.
         */
        function Control(element, options, invalidateOnResize) {
            var _this = this;
            if (options === void 0) { options = null; }
            if (invalidateOnResize === void 0) { invalidateOnResize = false; }
            this._focus = false; // whether the control currently contains the focus
            this._updating = 0; // update count (no refreshes while > 0)
            this._fullUpdate = false; // in case there are multiple calls to invalidate(x)
            /**
             * Occurs when the control gets the focus.
             */
            this.gotFocus = new wijmo.Event();
            /**
             * Occurs when the control loses the focus.
             */
            this.lostFocus = new wijmo.Event();
            // check that the element is not in use
            wijmo.assert(Control.getControl(element) == null, 'Element is already hosting a control.');
            // get the host element
            var host = wijmo.getElement(element);
            wijmo.assert(host != null, 'Cannot find the host element.');
            // save host and original content (to restore on dispose)
            this._orgOuter = host.outerHTML;
            this._orgInner = host.innerHTML;
            // save host attributes, replace <input> and <select> elements with <div>
            this._orgTag = host.tagName;
            if (host.tagName == 'INPUT' || host.tagName == 'SELECT') {
                host = this._replaceWithDiv(host);
            }
            // save host element and store control instance in element
            // (to retrieve with Control.getControl(element))
            this._e = host;
            host[Control._DATA_KEY] = this;
            // update layout when user resizes the browser
            if (invalidateOnResize == true) {
                this._szCtl = new wijmo.Size(host.offsetWidth, host.offsetHeight);
                var hr = this._handleResize.bind(this);
                this.addEventListener(window, 'resize', hr);
            }
            // fire events for got/lost focus
            // use timeOuts since Chrome and FF sometimes move the focus to the body
            // before moving it to the new focused element
            this.addEventListener(host, 'focus', function () {
                setTimeout(function () {
                    _this._updateFocusState();
                });
            }, true);
            this.addEventListener(host, 'blur', function () {
                setTimeout(function () {
                    _this._updateFocusState();
                }, 20); // >0, <200: TFS 100250, 112599, 115816, 195150
            }, true);
            // handle disabled controls 
            // (pointer-events requires IE11, doesn't prevent wheel at all)
            var hd = this._handleDisabled.bind(this);
            this.addEventListener(host, 'mousedown', hd, true);
            this.addEventListener(host, 'mouseup', hd, true);
            this.addEventListener(host, 'click', hd, true);
            this.addEventListener(host, 'dblclick', hd, true);
            this.addEventListener(host, 'keydown', hd, true);
            this.addEventListener(host, 'wheel', hd, true);
            // keep track of touch actions at the document level
            // (no need to add/remove event handlers to every Wijmo control)
            if (Control._touching == null) {
                Control._touching = false;
                if ('ontouchstart' in window || 'onpointerdown' in window) {
                    var b = document.body, ts = this._handleTouchStart, te = this._handleTouchEnd;
                    if ('ontouchstart' in window) {
                        b.addEventListener('touchstart', ts);
                        b.addEventListener('touchend', te);
                        b.addEventListener('touchcancel', te);
                        b.addEventListener('touchleave', te);
                    }
                    else if ('onpointerdown' in window) {
                        b.addEventListener('pointerdown', ts);
                        b.addEventListener('pointerup', te);
                        b.addEventListener('pointerout', te);
                        b.addEventListener('pointercancel', te);
                        b.addEventListener('pointerleave', te);
                    }
                }
            }
        }
        /**
         * Gets the HTML template used to create instances of the control.
         *
         * This method traverses up the class hierarchy to find the nearest ancestor that
         * specifies a control template. For example, if you specify a prototype for the
         * @see:ComboBox control, it will override the template defined by the @see:DropDown
         * base class.
         */
        Control.prototype.getTemplate = function () {
            for (var p = Object.getPrototypeOf(this); p; p = Object.getPrototypeOf(p)) {
                var tpl = p.constructor.controlTemplate;
                if (tpl) {
                    return tpl;
                }
            }
            return null;
        };
        /**
         * Applies the template to a new instance of a control, and returns the root element.
         *
         * This method should be called by constructors of templated controls.
         * It is responsible for binding the template parts to the
         * corresponding control members.
         *
         * For example, the code below applies a template to an instance
         * of an @see:InputNumber control. The template must contain elements
         * with the 'wj-part' attribute set to 'input', 'btn-inc', and 'btn-dec'.
         * The control members '_tbx', '_btnUp', and '_btnDn' will be assigned
         * references to these elements.
         *
         * <pre>this.applyTemplate('wj-control wj-inputnumber', template, {
         *   _tbx: 'input',
         *   _btnUp: 'btn-inc',
         *   _btnDn: 'btn-dec'
         * }, 'input');</pre>
         *
         * @param classNames Names of classes to add to the control's host element.
         * @param template An HTML string that defines the control template.
         * @param parts A dictionary of part variables and their names.
         * @param namePart Name of the part to be named after the host element. This
         * determines how the control submits data when used in forms.
         */
        Control.prototype.applyTemplate = function (classNames, template, parts, namePart) {
            var host = this._e;
            // apply standard classes to host element
            if (classNames) {
                wijmo.addClass(host, classNames);
            }
            // convert string into HTML template and append to host
            var tpl = null;
            if (template) {
                tpl = wijmo.createElement(template, host);
            }
            // copy key attributes from the host element (name, validation) to inner input 
            // element
            // NOTE 1: do this only if there is a single input element in the template
            // NOTE 2: do not copy 'type' since it causes issues in Chrome: TFS 84900, 84901
            var inputs = host.querySelectorAll('input'), input = (inputs.length == 1 ? inputs[0] : null), atts = host.attributes;
            if (input && atts) {
                for (var i = 0; i < atts.length; i++) {
                    if (atts[i].name.match(/name|autofocus|autocomplete|minlength|maxlength|pattern/i)) {
                        input.setAttribute(atts[i].name, atts[i].value);
                    }
                }
            }
            // change 'for' attribute of labels targeting the host element 
            // to target the inner input element instead (needed for Chrome and FF)
            if (input && host.id) {
                var label = document.querySelector('label[for="' + host.id + '"]'), baseId = host.id + '_input', newId = baseId;
                if (label instanceof HTMLLabelElement) {
                    for (var i = 0; document.getElementById(newId) != null; i++) {
                        newId = baseId + i; // new unique id
                    }
                    input.id = newId; // set id of inner input element
                    label.htmlFor = newId; // change 'for' attribute to match new id
                }
            }
            // fire 'change' events on behalf of inner input elements (TFS 190946)
            if (input) {
                var evtChange = document.createEvent('HTMLEvents'), orgVal = input.value;
                evtChange.initEvent('change', true, false);
                this.addEventListener(input, 'input', function () {
                    orgVal = input.value;
                }, true);
                this.gotFocus.addHandler(function () {
                    orgVal = input.value;
                });
                this.lostFocus.addHandler(function () {
                    if (orgVal != input.value) {
                        input.dispatchEvent(evtChange);
                    }
                });
            }
            // make sure the control can get the focus
            // this is a little tricky:
            // - Chrome won't give divs the focus unless we set tabIndex to something > -1
            // - But if we do set it and the control contains input elements, the back-tab key won't work
            // so we set the tabIndex to -1 or zero depending on whether the control contains input elements.
            // http://wijmo.com/topic/shift-tab-not-working-for-input-controls-in-ff-and-chrome/, TFS 123457
            if (!host.getAttribute('tabindex')) {
                host.tabIndex = (input != null) ? -1 : 0;
            }
            // initialize state (empty/invalid)
            this._updateState();
            // bind control variables to template parts
            if (parts) {
                for (var part in parts) {
                    var wjPart = parts[part];
                    this[part] = tpl.querySelector('[wj-part="' + wjPart + '"]');
                    // look in the root as well (querySelector doesn't...)
                    if (this[part] == null && tpl.getAttribute('wj-part') == wjPart) {
                        this[part] = tpl;
                    }
                    // make sure we found the part
                    if (this[part] == null) {
                        throw 'Missing template part: "' + wjPart + '"';
                    }
                    // copy/move attributes from host to input element
                    if (wjPart == namePart) {
                        // copy parent element's name attribute to the namePart element
                        // (to send data when submitting forms).
                        var key = 'name', att = host.attributes[key];
                        if (att && att.value) {
                            this[part].setAttribute(key, att.value);
                        }
                        // transfer access key
                        key = 'accesskey';
                        att = host.attributes[key];
                        if (att && att.value) {
                            this[part].setAttribute(key, att.value);
                            host.removeAttribute(key);
                        }
                    }
                }
            }
            // return template
            return tpl;
        };
        /**
         * Disposes of the control by removing its association with the host element.
         *
         * The @see:dispose method automatically removes any event listeners added
         * with the @see:addEventListener method.
         *
         * Calling the @see:dispose method is important in applications that create
         * and remove controls dynamically. Failing to dispose of the controls may
         * cause memory leaks.
         */
        Control.prototype.dispose = function () {
            // dispose of any child controls
            var cc = this._e.querySelectorAll('.wj-control');
            for (var i = 0; i < cc.length; i++) {
                var ctl = Control.getControl(cc[i]);
                if (ctl) {
                    ctl.dispose();
                }
            }
            // cancel any pending refreshes
            if (this._toInv) {
                clearTimeout(this._toInv);
            }
            // remove all HTML event listeners
            this.removeEventListener();
            // remove all Wijmo event listeners 
            // (without getting the value for all properties)
            for (var prop in this) {
                if (prop.length > 2 && prop.indexOf('on') == 0) {
                    var evt = this[prop[2].toLowerCase() + prop.substr(3)];
                    if (evt instanceof wijmo.Event) {
                        evt.removeAllHandlers();
                    }
                }
            }
            // if the control has a collectionView property, remove handlers to stop receiving notifications
            // REVIEW: perhaps optimize by caching the CollectionView properties?
            var cv = this['collectionView'];
            if (cv instanceof wijmo.collections.CollectionView) {
                for (var prop in cv) {
                    var evt = cv[prop];
                    if (evt instanceof wijmo.Event) {
                        evt.removeHandler(null, this);
                    }
                }
            }
            // restore original content
            if (this._e.parentNode) {
                this._e.outerHTML = this._orgOuter;
            }
            // done
            this._e[Control._DATA_KEY] = null;
            this._e = this._orgOuter = this._orgInner = this._orgTag = null;
        };
        /**
         * Gets the control that is hosted in a given DOM element.
         *
         * @param element The DOM element that is hosting the control, or a selector for the host element (e.g. '#theCtrl').
         */
        Control.getControl = function (element) {
            var e = wijmo.getElement(element);
            return e ? wijmo.asType(e[Control._DATA_KEY], Control, true) : null;
        };
        Object.defineProperty(Control.prototype, "hostElement", {
            /**
             * Gets the DOM element that is hosting the control.
             */
            get: function () {
                return this._e;
            },
            enumerable: true,
            configurable: true
        });
        /**
         * Sets the focus to this control.
         */
        Control.prototype.focus = function () {
            if (this._e) {
                this._e.focus();
            }
        };
        /**
         * Checks whether this control contains the focused element.
         */
        Control.prototype.containsFocus = function () {
            // test for disposed controls
            if (!this._e) {
                return false;
            }
            // scan child controls (they may have popups, TFS 112676)
            var c = this._e.querySelectorAll('.wj-control');
            for (var i = 0; i < c.length; i++) {
                var ctl = Control.getControl(c[i]);
                if (ctl && ctl != this && ctl.containsFocus()) {
                    return true;
                }
            }
            // see if active element is in a popup control
            // and its owner is contained in this control
            var ae = wijmo.getActiveElement(), popup = Control.getControl(wijmo.closest(ae, '.wj-control.wj-popup'));
            if (popup && popup.owner && wijmo.contains(this._e, popup.owner)) {
                return true;
            }
            // check for actual HTML containment
            return wijmo.contains(this._e, ae);
        };
        /**
         * Invalidates the control causing an asynchronous refresh.
         *
         * @param fullUpdate Whether to update the control layout as well as the content.
         */
        Control.prototype.invalidate = function (fullUpdate) {
            var _this = this;
            if (fullUpdate === void 0) { fullUpdate = true; }
            this._fullUpdate = this._fullUpdate || fullUpdate;
            if (this._toInv) {
                clearTimeout(this._toInv);
                this._toInv = null;
            }
            if (!this.isUpdating) {
                this._toInv = setTimeout(function () {
                    _this.refresh(_this._fullUpdate);
                }, Control._REFRESH_INTERVAL);
            }
        };
        /**
         * Refreshes the control.
         *
         * @param fullUpdate Whether to update the control layout as well as the content.
         */
        Control.prototype.refresh = function (fullUpdate) {
            if (fullUpdate === void 0) { fullUpdate = true; }
            if (!this.isUpdating && this._toInv) {
                clearTimeout(this._toInv);
                this._toInv = null;
                this._fullUpdate = false;
            }
            // derived classes should override this...
        };
        /**
         * Invalidates all Wijmo controls contained in an HTML element.
         *
         * Use this method when your application has dynamic panels that change
         * the control's visibility or dimensions. For example, splitters, accordions,
         * and tab controls usually change the visibility of its content elements.
         * In this case, failing to notify the controls contained in the element
         * may cause them to stop working properly.
         *
         * If this happens, you must handle the appropriate event in the dynamic
         * container and call the @see:Control.invalidateAll method so the contained
         * Wijmo controls will update their layout information properly.
         *
         * @param e Container element. If set to null, all Wijmo controls
         * on the page will be invalidated.
         */
        Control.invalidateAll = function (e) {
            if (!e)
                e = document.body;
            if (e.children) {
                for (var i = 0; i < e.children.length; i++) {
                    Control.invalidateAll(e.children[i]);
                }
            }
            var ctl = Control.getControl(e);
            if (ctl) {
                ctl.invalidate();
            }
        };
        /**
         * Refreshes all Wijmo controls contained in an HTML element.
         *
         * This method is similar to @see:invalidateAll, except the controls
         * are updated immediately rather than after an interval.
         *
         * @param e Container element. If set to null, all Wijmo controls
         * on the page will be invalidated.
         */
        Control.refreshAll = function (e) {
            if (!e)
                e = document.body;
            if (e.children) {
                for (var i = 0; i < e.children.length; i++) {
                    Control.refreshAll(e.children[i]);
                }
            }
            var ctl = Control.getControl(e);
            if (ctl) {
                ctl.refresh();
            }
        };
        /**
         * Disposes of all Wijmo controls contained in an HTML element.
         *
         * @param e Container element.
         */
        Control.disposeAll = function (e) {
            var ctl = Control.getControl(e);
            if (ctl) {
                ctl.dispose();
            }
            else if (e.children) {
                for (var i = 0; i < e.children.length; i++) {
                    Control.disposeAll(e.children[i]);
                }
            }
        };
        /**
         * Suspends notifications until the next call to @see:endUpdate.
         */
        Control.prototype.beginUpdate = function () {
            this._updating++;
        };
        /**
         * Resumes notifications suspended by calls to @see:beginUpdate.
         */
        Control.prototype.endUpdate = function () {
            this._updating--;
            if (this._updating <= 0) {
                this.invalidate();
            }
        };
        Object.defineProperty(Control.prototype, "isUpdating", {
            /**
             * Gets a value that indicates whether the control is currently being updated.
             */
            get: function () {
                return this._updating > 0;
            },
            enumerable: true,
            configurable: true
        });
        /**
         * Executes a function within a @see:beginUpdate/@see:endUpdate block.
         *
         * The control will not be updated until the function has been executed.
         * This method ensures @see:endUpdate is called even if the function throws
         * an exception.
         *
         * @param fn Function to be executed.
         */
        Control.prototype.deferUpdate = function (fn) {
            try {
                this.beginUpdate();
                fn();
            }
            finally {
                this.endUpdate();
            }
        };
        Object.defineProperty(Control.prototype, "isTouching", {
            /**
             * Gets a value that indicates whether the control is currently handling a touch event.
             */
            get: function () {
                return Control._touching;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Control.prototype, "isDisabled", {
            /**
             * Gets or sets a value that determines whether the control is disabled.
             *
             * Disabled controls cannot get mouse or keyboard events.
             */
            get: function () {
                return this._e && this._e.getAttribute('disabled') != null;
            },
            set: function (value) {
                value = wijmo.asBoolean(value, true);
                if (value != this.isDisabled) {
                    wijmo.enable(this._e, !value);
                }
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Control.prototype, "disabled", {
            /*
             * Deprecated: use 'isDisabled' instead to avoid confusion with 'disabled' HTML attribute.
             */
            get: function () {
                wijmo._deprecated('disabled', 'isDisabled');
                return this.isDisabled;
            },
            set: function (value) {
                wijmo._deprecated('disabled', 'isDisabled');
                this.isDisabled = value;
            },
            enumerable: true,
            configurable: true
        });
        /**
         * Initializes the control by copying the properties from a given object.
         *
         * This method allows you to initialize controls using plain data objects
         * instead of setting the value of each property in code.
         *
         * For example:
         * <pre>
         * grid.initialize({
         *   itemsSource: myList,
         *   autoGenerateColumns: false,
         *   columns: [
         *     { binding: 'id', header: 'Code', width: 130 },
         *     { binding: 'name', header: 'Name', width: 60 }
         *   ]
         * });
         * // is equivalent to
         * grid.itemsSource = myList;
         * grid.autoGenerateColumns = false;
         * // etc.
         * </pre>
         *
         * The initialization data is type-checked as it is applied. If the
         * initialization object contains unknown property names or invalid
         * data types, this method will throw.
         *
         * @param options Object that contains the initialization data.
         */
        Control.prototype.initialize = function (options) {
            if (options) {
                this.beginUpdate();
                wijmo.copy(this, options);
                this.endUpdate();
            }
        };
        /**
         * Adds an event listener to an element owned by this @see:Control.
         *
         * The control keeps a list of attached listeners and their handlers,
         * making it easier to remove them when the control is disposed (see the
         * @see:dispose and @see:removeEventListener method).
         *
         * Failing to remove event listeners may cause memory leaks.
         *
         * @param target Target element for the event.
         * @param type String that specifies the event.
         * @param fn Function to execute when the event occurs.
         * @param capture Whether the listener is capturing.
         */
        Control.prototype.addEventListener = function (target, type, fn, capture) {
            if (capture === void 0) { capture = false; }
            if (target) {
                target.addEventListener(type, fn, capture);
                if (this._listeners == null) {
                    this._listeners = [];
                }
                this._listeners.push({ target: target, type: type, fn: fn, capture: capture });
            }
        };
        /**
         * Removes one or more event listeners attached to elements owned by this @see:Control.
         *
         * @param target Target element for the event. If null, removes listeners attached to all targets.
         * @param type String that specifies the event. If null, removes listeners attached to all events.
         * @param capture Whether the listener is capturing. If null, removes capturing and non-capturing listeners.
         * @return The number of listeners removed.
         */
        Control.prototype.removeEventListener = function (target, type, capture) {
            var cnt = 0;
            if (this._listeners) {
                for (var i = 0; i < this._listeners.length; i++) {
                    var l = this._listeners[i];
                    if (target == null || target == l.target) {
                        if (type == null || type == l.type) {
                            if (capture == null || capture == l.capture) {
                                l.target.removeEventListener(l.type, l.fn, l.capture);
                                this._listeners.splice(i, 1);
                                i--;
                                cnt++;
                            }
                        }
                    }
                }
            }
            return cnt;
        };
        /**
         * Raises the @see:gotFocus event.
         */
        Control.prototype.onGotFocus = function (e) {
            this.gotFocus.raise(this, e);
        };
        /**
         * Raises the @see:lostFocus event.
         */
        Control.prototype.onLostFocus = function (e) {
            this.lostFocus.raise(this, e);
        };
        // ** implementation
        // invalidates the control when its size changes
        Control.prototype._handleResize = function () {
            if (this._e.parentElement) {
                var sz = new wijmo.Size(this._e.offsetWidth, this._e.offsetHeight);
                if (!sz.equals(this._szCtl)) {
                    this._szCtl = sz;
                    this.invalidate();
                }
            }
        };
        // update focus state and raise got/lost focus events
        Control.prototype._updateFocusState = function () {
            var _this = this;
            // use a timeOut since Chrome and FF sometimes move the focus to the body
            // before moving it to the new focused element (CellTemplateIntro sample)
            setTimeout(function () {
                // update focus for this control and all ancestors as well
                for (var e = _this._e; e; e = e.parentElement) {
                    var ctl = Control.getControl(e);
                    if (ctl) {
                        var focus = ctl.containsFocus();
                        if (focus != ctl._focus) {
                            ctl._focus = focus;
                            if (focus) {
                                ctl.onGotFocus();
                            }
                            else {
                                ctl.onLostFocus();
                            }
                            ctl._updateState();
                        }
                    }
                }
            }); // no interval: TFS 146949
        };
        // update state attributes for this control (focused, empty, invalid)
        Control.prototype._updateState = function () {
            var host = this.hostElement, input = this.hostElement.querySelector('input');
            wijmo.toggleClass(host, 'wj-state-focused', this.containsFocus());
            if (input instanceof HTMLInputElement) {
                wijmo.toggleClass(host, 'wj-state-empty', input.value.length == 0);
                wijmo.toggleClass(host, 'wj-state-readonly', input.readOnly);
                var vm = input.validationMessage; // may be null in IE9 (TFS 204492)
                wijmo.toggleClass(host, 'wj-state-invalid', vm && vm.length > 0);
            }
        };
        // keep track of touch events
        Control.prototype._handleTouchStart = function (e) {
            if (e.pointerType == null || e.pointerType == 'touch') {
                Control._touching = true;
            }
        };
        Control.prototype._handleTouchEnd = function (e) {
            if (e.pointerType == null || e.pointerType == 'touch') {
                setTimeout(function () {
                    Control._touching = false;
                    //console.log('touching = false');
                }, 400); // 300ms click event delay on IOS, plus some safety
            }
        };
        // suppress mouse and keyboard events if the control is disabled
        // (pointer-events requires IE11, doesn't prevent wheel at all)
        Control.prototype._handleDisabled = function (e) {
            if (this.isDisabled) {
                e.preventDefault();
                e.stopPropagation();
                e.stopImmediatePropagation();
            }
        };
        // replaces an element with a div element, copying the child elements 
        // and the 'id' and 'style' attributes from the original element
        Control.prototype._replaceWithDiv = function (element) {
            // replace the element
            var div = document.createElement('div');
            element.parentElement.replaceChild(div, element);
            // copy children
            div.innerHTML = element.innerHTML;
            // copy original attributes
            var atts = element.attributes;
            for (var i = 0; i < atts.length; i++) {
                var name = atts[i].name;
                if (atts[i].name.match(/id|style|class/i)) {
                    div.setAttribute(name, atts[i].value);
                }
            }
            // return new div
            return div;
        };
        Control._DATA_KEY = 'wj-Control'; // key used to store control reference in host element
        Control._REFRESH_INTERVAL = 10; // interval between invalidation and refresh
        return Control;
    }());
    wijmo.Control = Control;
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Control.js.map
var wijmo;
(function (wijmo) {
    'use strict';
    /**
     * Specifies the type of aggregate to calculate over a group of values.
     */
    (function (Aggregate) {
        /**
         * No aggregate.
         */
        Aggregate[Aggregate["None"] = 0] = "None";
        /**
         * Returns the sum of the numeric values in the group.
         */
        Aggregate[Aggregate["Sum"] = 1] = "Sum";
        /**
         * Returns the count of non-null values in the group.
         */
        Aggregate[Aggregate["Cnt"] = 2] = "Cnt";
        /**
         * Returns the average value of the numeric values in the group.
         */
        Aggregate[Aggregate["Avg"] = 3] = "Avg";
        /**
         * Returns the maximum value in the group.
         */
        Aggregate[Aggregate["Max"] = 4] = "Max";
        /**
         * Returns the minimum value in the group.
         */
        Aggregate[Aggregate["Min"] = 5] = "Min";
        /**
         * Returns the difference between the maximum and minimum numeric values in the group.
         */
        Aggregate[Aggregate["Rng"] = 6] = "Rng";
        /**
         * Returns the sample standard deviation of the numeric values in the group
         * (uses the formula based on n-1).
         */
        Aggregate[Aggregate["Std"] = 7] = "Std";
        /**
         * Returns the sample variance of the numeric values in the group
         * (uses the formula based on n-1).
         */
        Aggregate[Aggregate["Var"] = 8] = "Var";
        /**
         * Returns the population standard deviation of the values in the group
         * (uses the formula based on n).
         */
        Aggregate[Aggregate["StdPop"] = 9] = "StdPop";
        /**
         * Returns the population variance of the values in the group
         * (uses the formula based on n).
         */
        Aggregate[Aggregate["VarPop"] = 10] = "VarPop";
        /**
         * Returns the count of all values in the group (including nulls).
         */
        Aggregate[Aggregate["CntAll"] = 11] = "CntAll";
    })(wijmo.Aggregate || (wijmo.Aggregate = {}));
    var Aggregate = wijmo.Aggregate;
    /**
     * Calculates an aggregate value from the values in an array.
     *
     * @param aggType Type of aggregate to calculate.
     * @param items Array with the items to aggregate.
     * @param binding Name of the property to aggregate on (in case the items are not simple values).
     */
    function getAggregate(aggType, items, binding) {
        var cnt = 0, cntn = 0, sum = 0, sum2 = 0, min = null, max = null, bnd = binding ? new wijmo.Binding(binding) : null;
        // special case: overall count (including nulls)
        aggType = wijmo.asEnum(aggType, Aggregate);
        if (aggType == Aggregate.CntAll) {
            return items.length;
        }
        // calculate aggregate
        for (var i = 0; i < items.length; i++) {
            // get item/value
            var val = items[i];
            if (bnd) {
                val = bnd.getValue(val);
            }
            // aggregate
            if (val != null) {
                cnt++;
                if (min == null || val < min) {
                    min = val;
                }
                if (max == null || val > max) {
                    max = val;
                }
                if (wijmo.isNumber(val) && !isNaN(val)) {
                    cntn++;
                    sum += val;
                    sum2 += val * val;
                }
                else if (wijmo.isBoolean(val)) {
                    cntn++;
                    if (val == true) {
                        sum++;
                        sum2++;
                    }
                }
            }
        }
        // return result
        var avg = cntn == 0 ? 0 : sum / cntn;
        switch (aggType) {
            case Aggregate.Avg:
                return avg;
            case Aggregate.Cnt:
                return cnt;
            case Aggregate.Max:
                return max;
            case Aggregate.Min:
                return min;
            case Aggregate.Rng:
                return max - min;
            case Aggregate.Sum:
                return sum;
            case Aggregate.VarPop:
                return cntn <= 1 ? 0 : sum2 / cntn - avg * avg;
            case Aggregate.StdPop:
                return cntn <= 1 ? 0 : Math.sqrt(sum2 / cntn - avg * avg);
            case Aggregate.Var:
                return cntn <= 1 ? 0 : (sum2 / cntn - avg * avg) * cntn / (cntn - 1);
            case Aggregate.Std:
                return cntn <= 1 ? 0 : Math.sqrt((sum2 / cntn - avg * avg) * cntn / (cntn - 1));
        }
        // should never get here...
        throw 'Invalid aggregate type.';
    }
    wijmo.getAggregate = getAggregate;
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Aggregate.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
/**
 * Defines interfaces and classes related to data, including the @see:ICollectionView
 * interface, @see:CollectionView and @see:ObservableArray classes.
 */
var wijmo;
(function (wijmo) {
    var collections;
    (function (collections) {
        'use strict';
        /**
         * Describes the action that caused the @see:INotifyCollectionChanged.collectionChanged
         * event to fire.
         */
        (function (NotifyCollectionChangedAction) {
            /** An item was added to the collection. */
            NotifyCollectionChangedAction[NotifyCollectionChangedAction["Add"] = 0] = "Add";
            /** An item was removed from the collection. */
            NotifyCollectionChangedAction[NotifyCollectionChangedAction["Remove"] = 1] = "Remove";
            /** An item was changed or replaced. */
            NotifyCollectionChangedAction[NotifyCollectionChangedAction["Change"] = 2] = "Change";
            /**
             * Several items changed simultaneously
             * (for example, the collection was sorted, filtered, or grouped).
             */
            NotifyCollectionChangedAction[NotifyCollectionChangedAction["Reset"] = 3] = "Reset";
        })(collections.NotifyCollectionChangedAction || (collections.NotifyCollectionChangedAction = {}));
        var NotifyCollectionChangedAction = collections.NotifyCollectionChangedAction;
        /**
         * Provides data for the @see:INotifyCollectionChanged.collectionChanged event.
         */
        var NotifyCollectionChangedEventArgs = (function (_super) {
            __extends(NotifyCollectionChangedEventArgs, _super);
            /**
             * Initializes a new instance of the @see:NotifyCollectionChangedEventArgs class.
             *
             * @param action Type of action that caused the event to fire.
             * @param item Item that was added or changed.
             * @param index Index of the item.
             */
            function NotifyCollectionChangedEventArgs(action, item, index) {
                if (action === void 0) { action = NotifyCollectionChangedAction.Reset; }
                if (item === void 0) { item = null; }
                if (index === void 0) { index = -1; }
                _super.call(this);
                this.action = action;
                this.item = item;
                this.index = index;
            }
            /**
             * Provides a reset notification.
             */
            NotifyCollectionChangedEventArgs.reset = new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset);
            return NotifyCollectionChangedEventArgs;
        }(wijmo.EventArgs));
        collections.NotifyCollectionChangedEventArgs = NotifyCollectionChangedEventArgs;
        /**
         * Describes a sorting criterion.
         */
        var SortDescription = (function () {
            /**
             * Initializes a new instance of the @see:SortDescription class.
             *
             * @param property Name of the property to sort on.
             * @param ascending Whether to sort in ascending order.
             */
            function SortDescription(property, ascending) {
                this._bnd = new wijmo.Binding(property);
                this._asc = ascending;
            }
            Object.defineProperty(SortDescription.prototype, "property", {
                /**
                 * Gets the name of the property used to sort.
                 */
                get: function () {
                    return this._bnd.path;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(SortDescription.prototype, "ascending", {
                /**
                 * Gets a value that determines whether to sort the values in ascending order.
                 */
                get: function () {
                    return this._asc;
                },
                enumerable: true,
                configurable: true
            });
            return SortDescription;
        }());
        collections.SortDescription = SortDescription;
        /**
         * Provides data for the @see:IPagedCollectionView.pageChanging event
         */
        var PageChangingEventArgs = (function (_super) {
            __extends(PageChangingEventArgs, _super);
            /**
             * Initializes a new instance of the @see:PageChangingEventArgs class.
             *
             * @param newIndex Index of the page that is about to become current.
             */
            function PageChangingEventArgs(newIndex) {
                _super.call(this);
                this.newPageIndex = newIndex;
            }
            return PageChangingEventArgs;
        }(wijmo.CancelEventArgs));
        collections.PageChangingEventArgs = PageChangingEventArgs;
        /**
         * Represents a base class for types defining grouping conditions.
         *
         * The concrete class which is commonly used for this purpose is
         * @see:PropertyGroupDescription.
         */
        var GroupDescription = (function () {
            function GroupDescription() {
            }
            /**
             * Returns the group name for the given item.
             *
             * @param item The item to get group name for.
             * @param level The zero-based group level index.
             * @return The name of the group the item belongs to.
             */
            GroupDescription.prototype.groupNameFromItem = function (item, level) {
                return '';
            };
            /**
             * Returns a value that indicates whether the group name and the item name
             * match (which implies that the item belongs to the group).
             *
             * @param groupName The name of the group.
             * @param itemName The name of the item.
             * @return True if the names match; otherwise, false.
             */
            GroupDescription.prototype.namesMatch = function (groupName, itemName) {
                return groupName === itemName;
            };
            return GroupDescription;
        }());
        collections.GroupDescription = GroupDescription;
        /**
         * Describes the grouping of items using a property name as the criterion.
         *
         * For example, the code below causes a @see:CollectionView to group items
         * by the value of their 'country' property:
         * <pre>
         * var cv = new wijmo.collections.CollectionView(items);
         * var gd = new wijmo.collections.PropertyGroupDescription('country');
         * cv.groupDescriptions.push(gd);
         * </pre>
         *
         * You may also specify a callback function that generates the group name.
         * For example, the code below causes a @see:CollectionView to group items
         * by the first letter of the value of their 'country' property:
         * <pre>
         * var cv = new wijmo.collections.CollectionView(items);
         * var gd = new wijmo.collections.PropertyGroupDescription('country',
         *   function(item, propName) {
         *     return item[propName][0]; // return country's initial
         * });
         * cv.groupDescriptions.push(gd);
         * </pre>
         */
        var PropertyGroupDescription = (function (_super) {
            __extends(PropertyGroupDescription, _super);
            /**
             * Initializes a new instance of the @see:PropertyGroupDescription class.
             *
             * @param property The name of the property that specifies
             * which group an item belongs to.
             * @param converter A callback function that takes an item and
             * a property name and returns the group name. If not specified,
             * the group name is the property value for the item.
             */
            function PropertyGroupDescription(property, converter) {
                _super.call(this);
                this._bnd = new wijmo.Binding(property);
                this._converter = converter;
            }
            Object.defineProperty(PropertyGroupDescription.prototype, "propertyName", {
                /*
                 * Gets the name of the property that is used to determine which
                 * group an item belongs to.
                 */
                get: function () {
                    return this._bnd.path;
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Returns the group name for the given item.
             *
             * @param item The item to get group name for.
             * @param level The zero-based group level index.
             * @return The name of the group the item belongs to.
             */
            PropertyGroupDescription.prototype.groupNameFromItem = function (item, level) {
                return this._converter
                    ? this._converter(item, this.propertyName)
                    : this._bnd.getValue(item);
            };
            /**
             * Returns a value that indicates whether the group name and the item name
             * match (which implies that the item belongs to the group).
             *
             * @param groupName The name of the group.
             * @param itemName The name of the item.
             * @return True if the names match; otherwise, false.
             */
            PropertyGroupDescription.prototype.namesMatch = function (groupName, itemName) {
                return groupName === itemName;
            };
            return PropertyGroupDescription;
        }(GroupDescription));
        collections.PropertyGroupDescription = PropertyGroupDescription;
    })(collections = wijmo.collections || (wijmo.collections = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=ICollectionView.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var collections;
    (function (collections) {
        'use strict';
        /**
         * Base class for Array classes with notifications.
         */
        var ArrayBase = (function () {
            /**
             * Initializes a new instance of the @see:ArrayBase class.
             */
            function ArrayBase() {
                this.length = 0;
                Array.apply(this, arguments);
            }
            // keep TypeScript happy (these will never be called, we changed the prototype)
            ArrayBase.prototype.pop = function () {
                return null;
            };
            ArrayBase.prototype.push = function () {
                var item = [];
                for (var _i = 0; _i < arguments.length; _i++) {
                    item[_i - 0] = arguments[_i];
                }
                return 0;
            };
            ArrayBase.prototype.splice = function (index, count, value) {
                return null;
            };
            ArrayBase.prototype.slice = function (begin, end) {
                return null;
            };
            ArrayBase.prototype.indexOf = function (searchElement, fromIndex) {
                return -1;
            };
            ArrayBase.prototype.sort = function (compareFn) {
                return null;
            };
            return ArrayBase;
        }());
        collections.ArrayBase = ArrayBase;
        // inheriting from Array
        // NOTE: set this in declaration rather than in constructor so the
        // the TypeScript inheritance mechanism works correctly with instanceof.
        ArrayBase.prototype = Array.prototype;
        /**
         * Array that sends notifications on changes.
         *
         * The class raises the @see:collectionChanged event when changes are made with
         * the push, pop, splice, insert, or remove methods.
         *
         * Warning: Changes made by assigning values directly to array members or to the
         * length of the array do not raise the @see:collectionChanged event.
         */
        var ObservableArray = (function (_super) {
            __extends(ObservableArray, _super);
            /**
             * Initializes a new instance of the @see:ObservableArray class.
             *
             * @param data Array containing items used to populate the @see:ObservableArray.
             */
            function ObservableArray(data) {
                _super.call(this);
                this._updating = 0;
                // ** INotifyCollectionChanged
                /**
                 * Occurs when the collection changes.
                 */
                this.collectionChanged = new wijmo.Event();
                // initialize the array
                if (data) {
                    data = wijmo.asArray(data);
                    this._updating++;
                    for (var i = 0; i < data.length; i++) {
                        this.push(data[i]);
                    }
                    this._updating--;
                }
            }
            /**
             * Adds one or more items to the end of the array.
             *
             * @param ...item One or more items to add to the array.
             * @return The new length of the array.
             */
            ObservableArray.prototype.push = function () {
                var item = [];
                for (var _i = 0; _i < arguments.length; _i++) {
                    item[_i - 0] = arguments[_i];
                }
                var length = this.length;
                for (var i = 0; item && i < item.length; i++) {
                    length = _super.prototype.push.call(this, item[i]);
                    if (!this._updating) {
                        this._raiseCollectionChanged(collections.NotifyCollectionChangedAction.Add, item[i], length - 1);
                    }
                }
                return length;
            };
            /*
             * Removes the last item from the array.
             *
             * @return The item that was removed from the array.
             */
            ObservableArray.prototype.pop = function () {
                var item = _super.prototype.pop.call(this);
                this._raiseCollectionChanged(collections.NotifyCollectionChangedAction.Remove, item, this.length);
                return item;
            };
            /**
             * Removes and/or adds items to the array.
             *
             * @param index Position where items will be added or removed.
             * @param count Number of items to remove from the array.
             * @param item Item to add to the array.
             * @return An array containing the removed elements.
             */
            ObservableArray.prototype.splice = function (index, count, item) {
                var rv;
                if (count && item) {
                    rv = _super.prototype.splice.call(this, index, count, item);
                    if (count == 1) {
                        this._raiseCollectionChanged(collections.NotifyCollectionChangedAction.Change, item, index);
                    }
                    else {
                        this._raiseCollectionChanged();
                    }
                    return rv;
                }
                else if (item) {
                    rv = _super.prototype.splice.call(this, index, 0, item);
                    this._raiseCollectionChanged(collections.NotifyCollectionChangedAction.Add, item, index);
                    return rv;
                }
                else {
                    rv = _super.prototype.splice.call(this, index, count);
                    if (count == 1) {
                        this._raiseCollectionChanged(collections.NotifyCollectionChangedAction.Remove, rv[0], index);
                    }
                    else {
                        this._raiseCollectionChanged();
                    }
                    return rv;
                }
            };
            /**
             * Creates a shallow copy of a portion of an array.
             *
             * @param begin Position where the copy starts.
             * @param end Position where the copy ends.
             * @return A shallow copy of a portion of an array.
             */
            ObservableArray.prototype.slice = function (begin, end) {
                return _super.prototype.slice.call(this, begin, end);
            };
            /**
             * Searches for an item in the array.
             *
             * @param searchElement Element to locate in the array.
             * @param fromIndex The index where the search should start.
             * @return The index of the item in the array, or -1 if the item was not found.
             */
            ObservableArray.prototype.indexOf = function (searchElement, fromIndex) {
                return _super.prototype.indexOf.call(this, searchElement, fromIndex);
            };
            /**
             * Sorts the elements of the array in place.
             *
             * @param compareFn Specifies a function that defines the sort order.
             * If specified, the function should take two arguments and should return
             * -1, +1, or 0 to indicate the first argument is smaller, greater than,
             * or equal to the second argument.
             * If omitted, the array is sorted in dictionary order according to the
             * string conversion of each element.
             * @return A copy of the sorted array.
             */
            ObservableArray.prototype.sort = function (compareFn) {
                var rv = _super.prototype.sort.call(this, compareFn);
                this._raiseCollectionChanged();
                return rv;
            };
            /**
             * Inserts an item at a specific position in the array.
             *
             * @param index Position where the item will be added.
             * @param item Item to add to the array.
             */
            ObservableArray.prototype.insert = function (index, item) {
                this.splice(index, 0, item);
            };
            /**
             * Removes an item from the array.
             *
             * @param item Item to remove.
             * @return True if the item was removed, false if it wasn't found in the array.
             */
            ObservableArray.prototype.remove = function (item) {
                var index = this.indexOf(item);
                if (index > -1) {
                    this.removeAt(index);
                    return true;
                }
                return false;
            };
            /**
             * Removes an item at a specific position in the array.
             *
             * @param index Position of the item to remove.
             */
            ObservableArray.prototype.removeAt = function (index) {
                this.splice(index, 1);
            };
            /**
             * Assigns an item at a specific position in the array.
             *
             * @param index Position where the item will be assigned.
             * @param item Item to assign to the array.
             */
            ObservableArray.prototype.setAt = function (index, item) {
                // make sure we have enough elements to set at the right index!
                if (index > this.length) {
                    this.length = index;
                }
                // go ahead and splice now
                this.splice(index, 1, item);
            };
            /**
             * Removes all items from the array.
             */
            ObservableArray.prototype.clear = function () {
                if (this.length !== 0) {
                    this.splice(0, this.length); // safer than setting length = 0
                    //this.length = 0; // fastest way to clear an array
                    this._raiseCollectionChanged();
                }
            };
            /**
             * Suspends notifications until the next call to @see:endUpdate.
             */
            ObservableArray.prototype.beginUpdate = function () {
                this._updating++;
            };
            /**
             * Resumes notifications suspended by a call to @see:beginUpdate.
             */
            ObservableArray.prototype.endUpdate = function () {
                if (this._updating > 0) {
                    this._updating--;
                    if (this._updating == 0) {
                        this._raiseCollectionChanged();
                    }
                }
            };
            Object.defineProperty(ObservableArray.prototype, "isUpdating", {
                /**
                 * Gets a value that indicates whether notifications are currently suspended
                 * (see @see:beginUpdate and @see:endUpdate).
                 */
                get: function () {
                    return this._updating > 0;
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Executes a function within a @see:beginUpdate/@see:endUpdate block.
             *
             * The collection will not be refreshed until the function finishes.
             * This method ensures @see:endUpdate is called even if the function throws
             * an exception.
             *
             * @param fn Function to be executed without updates.
             */
            ObservableArray.prototype.deferUpdate = function (fn) {
                try {
                    this.beginUpdate();
                    fn();
                }
                finally {
                    this.endUpdate();
                }
            };
            // ** IQueryInterface
            /**
             * Returns true if the caller queries for a supported interface.
             *
             * @param interfaceName Name of the interface to look for.
             * @return True if the caller queries for a supported interface.
             */
            ObservableArray.prototype.implementsInterface = function (interfaceName) {
                return interfaceName == 'INotifyCollectionChanged';
            };
            /**
             * Raises the @see:collectionChanged event.
             *
             * @param e Contains a description of the change.
             */
            ObservableArray.prototype.onCollectionChanged = function (e) {
                if (e === void 0) { e = collections.NotifyCollectionChangedEventArgs.reset; }
                if (!this.isUpdating) {
                    this.collectionChanged.raise(this, e);
                }
            };
            // creates event args and calls onCollectionChanged
            ObservableArray.prototype._raiseCollectionChanged = function (action, item, index) {
                if (action === void 0) { action = collections.NotifyCollectionChangedAction.Reset; }
                if (!this.isUpdating) {
                    var e = new collections.NotifyCollectionChangedEventArgs(action, item, index);
                    this.onCollectionChanged(e);
                }
            };
            return ObservableArray;
        }(ArrayBase));
        collections.ObservableArray = ObservableArray;
    })(collections = wijmo.collections || (wijmo.collections = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=ObservableArray.js.map
var wijmo;
(function (wijmo) {
    var collections;
    (function (collections) {
        'use strict';
        /**
         * Class that implements the @see:ICollectionView interface to expose data in
         * regular JavaScript arrays.
         *
         * The @see:CollectionView class implements the following interfaces:
         * <ul>
         *   <li>@see:ICollectionView: provides current record management,
         *       custom sorting, filtering, and grouping.</li>
         *   <li>@see:IEditableCollectionView: provides methods for editing,
         *       adding, and removing items.</li>
         *   <li>@see:IPagedCollectionView: provides paging.</li>
         * </ul>
         *
         * To use the @see:CollectionView class, start by declaring it and passing a
         * regular array as a data source. Then configure the view using the
         * @see:filter, @see:sortDescriptions, @see:groupDescriptions, and
         * @see:pageSize properties. Finally, access the view using the @see:items
         * property. For example:
         *
         * <pre>// create a new CollectionView
         * var cv = new wijmo.collections.CollectionView(myArray);
         *
         * // sort items by amount in descending order
         * var sd = new wijmo.collections.SortDescription('amount', false);
         * cv.sortDescriptions.push(sd);
         *
         * // show only items with amounts greater than 100
         * cv.filter = function(item) { return item.amount &gt; 100 };
         *
         * // show the sorted, filtered result on the console
         * for (var i = 0; i &lt; cv.items.length; i++) {
         *   var item = cv.items[i];
         *   console.log(i + ': ' + item.name + ' ' + item.amount);
         * }</pre>
         */
        var CollectionView = (function () {
            /**
             * Initializes a new instance of the @see:CollectionView class.
             *
             * @param sourceCollection Array that serves as a source for this
             * @see:CollectionView.
             * @param options JavaScript object containing initialization data for the control.
             */
            function CollectionView(sourceCollection, options) {
                var _this = this;
                this._idx = -1;
                this._srtDsc = new collections.ObservableArray();
                this._grpDesc = new collections.ObservableArray();
                this._newItem = null;
                this._edtItem = null;
                this._pgSz = 0;
                this._pgIdx = 0;
                this._updating = 0;
                this._canFilter = true;
                this._canGroup = true;
                this._canSort = true;
                this._canAddNew = true;
                this._canCancelEdit = true;
                this._canRemove = true;
                this._canChangePage = true;
                this._trackChanges = false;
                this._chgAdded = new collections.ObservableArray();
                this._chgRemoved = new collections.ObservableArray();
                this._chgEdited = new collections.ObservableArray();
                // ** INotifyCollectionChanged
                /**
                 * Occurs when the collection changes.
                 */
                this.collectionChanged = new wijmo.Event();
                /**
                 * Occurs before the value of the @see:sourceCollection property changes.
                 */
                this.sourceCollectionChanging = new wijmo.Event();
                /**
                 * Occurs after the value of the @see:sourceCollection property changes.
                 */
                this.sourceCollectionChanged = new wijmo.Event();
                /**
                 * Occurs after the current item changes.
                 */
                this.currentChanged = new wijmo.Event();
                /**
                 * Occurs before the current item changes.
                 */
                this.currentChanging = new wijmo.Event();
                /**
                * Occurs after the page index changes.
                */
                this.pageChanged = new wijmo.Event();
                /**
                 * Occurs before the page index changes.
                 */
                this.pageChanging = new wijmo.Event();
                // check that sortDescriptions contains SortDescriptions
                this._srtDsc.collectionChanged.addHandler(function () {
                    var arr = _this._srtDsc;
                    for (var i = 0; i < arr.length; i++) {
                        wijmo.assert(arr[i] instanceof collections.SortDescription, 'sortDescriptions array must contain SortDescription objects.');
                    }
                    if (_this.canSort) {
                        _this.refresh();
                    }
                });
                // check that groupDescriptions contains GroupDescriptions
                this._grpDesc.collectionChanged.addHandler(function () {
                    var arr = _this._grpDesc;
                    for (var i = 0; i < arr.length; i++) {
                        wijmo.assert(arr[i] instanceof collections.GroupDescription, 'groupDescriptions array must contain GroupDescription objects.');
                    }
                    if (_this.canGroup) {
                        _this.refresh();
                    }
                });
                // initialize the source collection
                this.sourceCollection = sourceCollection ? sourceCollection : new collections.ObservableArray();
                // apply options
                if (options) {
                    this.beginUpdate();
                    wijmo.copy(this, options);
                    this.endUpdate();
                }
            }
            // method used in JSON-style initialization
            CollectionView.prototype._copy = function (key, value) {
                if (key == 'sortDescriptions') {
                    this.sortDescriptions.clear();
                    var arr = wijmo.asArray(value);
                    for (var i = 0; i < arr.length; i++) {
                        var val = arr[i];
                        if (wijmo.isString(val)) {
                            val = new collections.SortDescription(val, true);
                        }
                        this.sortDescriptions.push(val);
                    }
                    return true;
                }
                if (key == 'groupDescriptions') {
                    this.groupDescriptions.clear();
                    var arr = wijmo.asArray(value);
                    for (var i = 0; i < arr.length; i++) {
                        var val = arr[i];
                        if (wijmo.isString(val)) {
                            val = new collections.PropertyGroupDescription(val);
                        }
                        this.groupDescriptions.push(val);
                    }
                    return true;
                }
                return false;
            };
            Object.defineProperty(CollectionView.prototype, "newItemCreator", {
                /**
                 * Gets or sets a function that creates new items for the collection.
                 *
                 * If the creator function is not supplied, the @see:CollectionView
                 * will try to create an uninitialized item of the appropriate type.
                 *
                 * If the creator function is supplied, it should be a function that
                 * takes no parameters and returns an initialized object of the proper
                 * type for the collection.
                 */
                get: function () {
                    return this._itemCreator;
                },
                set: function (value) {
                    this._itemCreator = wijmo.asFunction(value);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "sortConverter", {
                /**
                 * Gets or sets a function used to convert values when sorting.
                 *
                 * If provided, the function should take as parameters a
                 * @see:SortDescription, a data item, and a value to convert,
                 * and should return the converted value.
                 *
                 * This property provides a way to customize sorting. For example,
                 * the @see:FlexGrid control uses it to sort mapped columns by
                 * display value instead of by raw value.
                 *
                 * For example, the code below causes a @see:CollectionView to
                 * sort the 'country' property, which contains country code integers,
                 * using the corresponding country names:
                 *
                 * <pre>var countries = 'US,Germany,UK,Japan,Italy,Greece'.split(',');
                 * collectionView.sortConverter = function (sd, item, value) {
                 *   if (sd.property == 'countryMapped') {
                 *     value = countries[value]; // convert country id into name
                 *   }
                 *   return value;
                 * }</pre>
                 */
                get: function () {
                    return this._srtCvt;
                },
                set: function (value) {
                    if (value != this._srtCvt) {
                        this._srtCvt = wijmo.asFunction(value, true);
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "sortComparer", {
                /**
                 * Gets or sets a function used to compare values when sorting.
                 *
                 * If provided, the sort comparer function should take as parameters
                 * two values of any type, and should return -1, 0, or +1 to indicate
                 * whether the first value is smaller than, equal to, or greater than
                 * the second. If the sort comparer returns null, the standard built-in
                 * comparer is used.
                 *
                 * This @see:sortComparer property allows you to use custom comparison
                 * algorithms that in some cases result in sorting sequences that are
                 * more consistent with user's expectations than plain string comparisons.
                 *
                 * For example, see <a href="http://www.davekoelle.com/alphanum.html">Dave Koele's Alphanum algorithm</a>.
                 * It breaks up strings into chunks composed of strings or numbers, then
                 * sorts number chunks in value order and string chunks in ASCII order.
                 * Dave calls the result a "natural sorting order".
                 *
                 * The example below shows a typical use for the @see:sortComparer property:
                 * <pre>// create a CollectionView with a custom sort comparer
                 * var dataCustomSort = new wijmo.collections.CollectionView(data, {
                 *   sortComparer: function (a, b) {
                 *     return wijmo.isString(a) && wijmo.isString(b)
                 *       ? alphanum(a, b) // custom comparer used for strings
                 *       : null; // use default comparer used for everything else
                 *   }
                 * });</pre>
                 */
                get: function () {
                    return this._srtCmp;
                },
                set: function (value) {
                    if (value != this._srtCmp) {
                        this._srtCmp = wijmo.asFunction(value, true);
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "useStableSort", {
                /**
                 * Gets or sets whether to use a stable sort algorithm.
                 *
                 * Stable sorting algorithms maintain the relative order of records with equal keys.
                 * For example, consider a collection of objects with an "Amount" field.
                 * If you sort the collection by "Amount", a stable sort will keep the original
                 * order of records with the same Amount value.
                 *
                 * This property is false by default, which causes the @see:CollectionView to use
                 * JavaScript's built-in sort method, which is very fast but not stable. Setting
                 * the @see:useStableSort property to true increases sort times by 30% to 50%, which
                 * can be significant for large collections.
                 */
                get: function () {
                    return this._stableSort;
                },
                set: function (value) {
                    this._stableSort = wijmo.asBoolean(value);
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Calculates an aggregate value for the items in this collection.
             *
             * @param aggType Type of aggregate to calculate.
             * @param binding Property to aggregate on.
             * @param currentPage Whether to include only items on the current page.
             * @return The aggregate value.
             */
            CollectionView.prototype.getAggregate = function (aggType, binding, currentPage) {
                var items = currentPage ? this._pgView : this._view;
                return wijmo.getAggregate(aggType, items, binding);
            };
            // ** IQueryInterface
            /**
             * Returns true if the caller queries for a supported interface.
             *
             * @param interfaceName Name of the interface to look for.
             */
            CollectionView.prototype.implementsInterface = function (interfaceName) {
                switch (interfaceName) {
                    case 'ICollectionView':
                    case 'IEditableCollectionView':
                    case 'IPagedCollectionView':
                    case 'INotifyCollectionChanged':
                        return true;
                }
                return false;
            };
            Object.defineProperty(CollectionView.prototype, "trackChanges", {
                /**
                 * Gets or sets a value that determines whether the control should
                 * track changes to the data.
                 *
                 * If @see:trackChanges is set to true, the @see:CollectionView keeps
                 * track of changes to the data and exposes them through the
                 * @see:itemsAdded, @see:itemsRemoved, and @see:itemsEdited collections.
                 *
                 * Tracking changes is useful in situations where you need to update
                 * the server after the user has confirmed that the modifications are
                 * valid.
                 *
                 * After committing or cancelling changes, use the @see:clearChanges method
                 * to clear the @see:itemsAdded, @see:itemsRemoved, and @see:itemsEdited
                 * collections.
                 *
                 * The @see:CollectionView only tracks changes made when the proper
                 * @see:CollectionView methods are used (@see:editItem/@see:commitEdit,
                 * @see:addNew/@see:commitNew, and @see:remove).
                 * Changes made directly to the data are not tracked.
                 */
                get: function () {
                    return this._trackChanges;
                },
                set: function (value) {
                    this._trackChanges = wijmo.asBoolean(value);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "itemsAdded", {
                /**
                 * Gets an @see:ObservableArray containing the records that were added to
                 * the collection since @see:trackChanges was enabled.
                 */
                get: function () {
                    return this._chgAdded;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "itemsRemoved", {
                /**
                 * Gets an @see:ObservableArray containing the records that were removed from
                 * the collection since @see:trackChanges was enabled.
                 */
                get: function () {
                    return this._chgRemoved;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "itemsEdited", {
                /**
                 * Gets an @see:ObservableArray containing the records that were edited in
                 * the collection since @see:trackChanges was enabled.
                 */
                get: function () {
                    return this._chgEdited;
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Clears all changes by removing all items in the @see:itemsAdded,
             * @see:itemsRemoved, and @see:itemsEdited collections.
             *
             * Call this method after committing changes to the server or
             * after refreshing the data from the server.
             */
            CollectionView.prototype.clearChanges = function () {
                this._chgAdded.clear();
                this._chgRemoved.clear();
                this._chgEdited.clear();
            };
            /**
             * Raises the @see:collectionChanged event.
             *
             * @param e Contains a description of the change.
             */
            CollectionView.prototype.onCollectionChanged = function (e) {
                if (e === void 0) { e = collections.NotifyCollectionChangedEventArgs.reset; }
                // track changes applied to items outside of editItem/commitEdit blocks (TFS 204805)
                if (!this._committing && !this._canceling &&
                    e.action == collections.NotifyCollectionChangedAction.Change &&
                    e.item != this.currentEditItem) {
                    this._trackItemChanged(e.item);
                }
                // raise the event as usual
                this.collectionChanged.raise(this, e);
            };
            // creates event args and calls onCollectionChanged
            CollectionView.prototype._raiseCollectionChanged = function (action, item, index) {
                if (action === void 0) { action = collections.NotifyCollectionChangedAction.Reset; }
                //console.log('** collection changed: ' + NotifyCollectionChangedAction[action] + ' **');
                var e = new collections.NotifyCollectionChangedEventArgs(action, item, index);
                this.onCollectionChanged(e);
            };
            /**
             * Raises the @see:sourceCollectionChanging event.
             *
             * @param e @see:CancelEventArgs that contains the event data.
             */
            CollectionView.prototype.onSourceCollectionChanging = function (e) {
                this.sourceCollectionChanging.raise(this, e);
                return !e.cancel;
            };
            /**
             * Raises the @see:sourceCollectionChanged event.
             */
            CollectionView.prototype.onSourceCollectionChanged = function (e) {
                this.sourceCollectionChanged.raise(this, e);
            };
            Object.defineProperty(CollectionView.prototype, "canFilter", {
                // ** ICollectionView
                /**
                 * Gets a value that indicates whether this view supports filtering via the
                 * @see:filter property.
                 */
                get: function () {
                    return this._canFilter;
                },
                set: function (value) {
                    this._canFilter = wijmo.asBoolean(value);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "canGroup", {
                /**
                 * Gets a value that indicates whether this view supports grouping via the
                 * @see:groupDescriptions property.
                 */
                get: function () {
                    return this._canGroup;
                },
                set: function (value) {
                    this._canGroup = wijmo.asBoolean(value);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "canSort", {
                /**
                 * Gets a value that indicates whether this view supports sorting via the
                 * @see:sortDescriptions property.
                 */
                get: function () {
                    return this._canSort;
                },
                set: function (value) {
                    this._canSort = wijmo.asBoolean(value);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "currentItem", {
                /**
                 * Gets or sets the current item in the view.
                 */
                get: function () {
                    return this._pgView && this._idx > -1 && this._idx < this._pgView.length
                        ? this._pgView[this._idx]
                        : null;
                },
                set: function (value) {
                    this.moveCurrentTo(value);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "currentPosition", {
                /**
                 * Gets the ordinal position of the current item in the view.
                 */
                get: function () {
                    return this._idx;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "filter", {
                /**
                 * Gets or sets a callback used to determine if an item is suitable for
                 * inclusion in the view.
                 *
                 * The callback function should return true if the item passed in as a
                 * parameter should be included in the view.
                 *
                 * NOTE: If the filter function needs a scope (i.e. a meaningful 'this'
                 * value) remember to set the filter using the 'bind' function to specify
                 * the 'this' object. For example:
                 * <pre>
                 *   collectionView.filter = this._filter.bind(this);
                 * </pre>
                 */
                get: function () {
                    return this._filter;
                },
                set: function (value) {
                    if (this._filter != value) {
                        this._filter = wijmo.asFunction(value);
                        if (this.canFilter) {
                            this.refresh();
                        }
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "groupDescriptions", {
                /**
                 * Gets a collection of @see:GroupDescription objects that describe how the
                 * items in the collection are grouped in the view.
                 */
                get: function () {
                    return this._grpDesc;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "groups", {
                /**
                 * Gets an array of @see:CollectionViewGroup objects that represents the
                 * top-level groups.
                 */
                get: function () {
                    return this._groups;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "isEmpty", {
                /**
                 * Gets a value that indicates whether this view contains no items.
                 */
                get: function () {
                    return !this._pgView || !this._pgView.length;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "sortDescriptions", {
                /**
                 * Gets a collection of @see:SortDescription objects that describe how the items
                 * in the collection are sorted in the view.
                 */
                get: function () {
                    return this._srtDsc;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "sourceCollection", {
                /**
                 * Gets or sets the underlying (unfiltered and unsorted) collection.
                 */
                get: function () {
                    return this._src;
                },
                set: function (sourceCollection) {
                    if (sourceCollection != this._src) {
                        // raise changing event
                        if (!this.onSourceCollectionChanging(new wijmo.CancelEventArgs())) {
                            return;
                        }
                        // keep track of current index
                        var index = this.currentPosition;
                        // commit pending changes
                        this.commitEdit();
                        this.commitNew();
                        // disconnect old source
                        if (this._ncc != null) {
                            this._ncc.collectionChanged.removeHandler(this._sourceChanged);
                        }
                        // connect new source
                        this._src = wijmo.asArray(sourceCollection, false);
                        this._ncc = wijmo.tryCast(this._src, 'INotifyCollectionChanged');
                        if (this._ncc) {
                            this._ncc.collectionChanged.addHandler(this._sourceChanged, this);
                        }
                        // clear any changes
                        this.clearChanges();
                        // refresh view
                        this.refresh();
                        this.moveCurrentToFirst();
                        // raise changed event
                        this.onSourceCollectionChanged();
                        // if we have no items, notify listeners that the current index changed
                        if (this.currentPosition < 0 && index > -1) {
                            this.onCurrentChanged();
                        }
                    }
                },
                enumerable: true,
                configurable: true
            });
            // handle notifications from the source collection
            CollectionView.prototype._sourceChanged = function (s, e) {
                if (this._updating <= 0) {
                    this.refresh(); // TODO: optimize
                }
            };
            /**
             * Returns a value indicating whether a given item belongs to this view.
             *
             * @param item Item to seek.
             */
            CollectionView.prototype.contains = function (item) {
                return this._pgView.indexOf(item) > -1;
            };
            /**
             * Sets the specified item to be the current item in the view.
             *
             * @param item Item that will become current.
             */
            CollectionView.prototype.moveCurrentTo = function (item) {
                return this.moveCurrentToPosition(this._pgView.indexOf(item));
            };
            /**
             * Sets the first item in the view as the current item.
             */
            CollectionView.prototype.moveCurrentToFirst = function () {
                return this.moveCurrentToPosition(0);
            };
            /**
             * Sets the last item in the view as the current item.
             */
            CollectionView.prototype.moveCurrentToLast = function () {
                return this.moveCurrentToPosition(this._pgView.length - 1);
            };
            /**
             * Sets the item before the current item in the view as the current item.
             */
            CollectionView.prototype.moveCurrentToPrevious = function () {
                return this._idx > 0 ? this.moveCurrentToPosition(this._idx - 1) : false;
            };
            /**
             * Sets the item after the current item in the view as the current item.
             */
            CollectionView.prototype.moveCurrentToNext = function () {
                return this.moveCurrentToPosition(this._idx + 1);
            };
            /**
             * Sets the item at the specified index in the view as the current item.
             *
             * @param index Index of the item that will become current.
             */
            CollectionView.prototype.moveCurrentToPosition = function (index) {
                if (index >= -1 && index < this._pgView.length) {
                    var e = new wijmo.CancelEventArgs();
                    if (this._idx != index && this.onCurrentChanging(e)) {
                        // when moving away from current edit/new item, commit
                        if (this._edtItem && this._pgView[index] != this._edtItem) {
                            this.commitEdit();
                        }
                        if (this._newItem && this._pgView[index] != this._newItem) {
                            this.commitNew();
                        }
                        // update currency
                        this._idx = index;
                        this.onCurrentChanged();
                    }
                }
                return this._idx == index;
            };
            /**
             * Re-creates the view using the current sort, filter, and group parameters.
             */
            CollectionView.prototype.refresh = function () {
                // not while updating, adding, or editing
                if (this._updating > 0 || this._newItem || this._edtItem) {
                    return;
                }
                // perform the refresh
                this._performRefresh();
                // notify listeners
                this.onCollectionChanged();
            };
            // performs the refresh (without issuing notifications)
            CollectionView.prototype._performRefresh = function () {
                // benchmark
                //var start = new Date();
                // save current item
                var current = this.currentItem;
                // create filtered view
                if (!this._src) {
                    this._view = [];
                }
                else if (!this._filter || !this.canFilter) {
                    this._view = (this._srtDsc.length > 0 && this.canSort)
                        ? this._src.slice(0) // clone source array
                        : this._src; // don't waste time cloning
                }
                else {
                    this._view = this._performFilter(this._src);
                }
                // apply sort
                if (this.canSort && this._srtDsc.length > 0) {
                    this._performSort(this._view);
                }
                // apply grouping
                this._groups = this.canGroup ? this._createGroups(this._view) : null;
                this._fullGroups = this._groups;
                if (this._groups) {
                    this._view = this._mergeGroupItems(this._groups);
                }
                // apply paging to view
                this._pgIdx = wijmo.clamp(this._pgIdx, 0, this.pageCount - 1);
                this._pgView = this._getPageView();
                // update groups to take paging into account
                if (this._groups && this.pageCount > 1) {
                    this._groups = this._createGroups(this._pgView);
                    this._mergeGroupItems(this._groups);
                }
                // restore current item
                var index = this._pgView.indexOf(current);
                if (index < 0) {
                    index = Math.min(this._idx, this._pgView.length - 1);
                }
                this._idx = index;
                // save group digest to optimize updates (TFS 109119)
                this._digest = this._getGroupsDigest(this.groups);
                // raise currentChanged if needed
                if (this.currentItem !== current) {
                    this.onCurrentChanged();
                }
                //var now = new Date();
                //console.log('refreshed in ' + (now.getTime() - start.getTime()) / 1000 + ' seconds');
            };
            // sorts an array in-place using the current sort descriptions
            CollectionView.prototype._performSort = function (items) {
                if (this._stableSort) {
                    // stable sort (nice, but 30-50% slower)
                    // https://bugs.chromium.org/p/v8/issues/detail?id=90
                    var arrIndexed = items.map(function (item, index) { return { item: item, index: index }; }), compare = this._compareItems();
                    arrIndexed.sort(function (a, b) {
                        var r = compare(a.item, b.item);
                        return r == 0 ? a.index - b.index : r;
                    });
                    for (var i = 0; i < items.length; i++) {
                        items[i] = arrIndexed[i].item;
                    }
                }
                else {
                    // regular sort, not stable but very fast
                    items.sort(this._compareItems());
                }
            };
            // this function is used in some of our samples, so
            // if we remove it or change its name some things will break...
            CollectionView.prototype._compareItems = function () {
                var srtDsc = this._srtDsc, srtCvt = this._srtCvt, srtCmp = this._srtCmp, init = true, cmp = 0;
                return function (a, b) {
                    for (var i = 0; i < srtDsc.length; i++) {
                        // get values
                        var sd = srtDsc[i], v1 = sd._bnd.getValue(a), v2 = sd._bnd.getValue(b);
                        // custom converter function (before changing case! TFS 149638)
                        if (srtCvt) {
                            v1 = srtCvt(sd, a, v1, init);
                            v2 = srtCvt(sd, b, v2, false);
                            init = false;
                        }
                        // custom comparison function (TFS 151665)
                        if (srtCmp) {
                            cmp = srtCmp(v1, v2);
                            if (cmp != null) {
                                return sd.ascending ? +cmp : -cmp;
                            }
                        }
                        // check for NaN (isNaN returns true for NaN but also for non-numbers)
                        if (v1 !== v1)
                            v1 = null;
                        if (v2 !== v2)
                            v2 = null;
                        // ignore case when sorting unless the values are strings that only differ in case
                        // (to keep the sort consistent, TFS 131135)
                        if (typeof (v1) === 'string' && typeof (v2) === 'string') {
                            var lc1 = v1.toLowerCase(), lc2 = v2.toLowerCase();
                            if (lc1 != lc2) {
                                v1 = lc1;
                                v2 = lc2;
                            }
                        }
                        // nulls always at the bottom (like excel)
                        if (v1 != null && v2 == null)
                            return -1;
                        if (v1 == null && v2 != null)
                            return +1;
                        // compare the values (at last!)
                        cmp = (v1 < v2) ? -1 : (v1 > v2) ? +1 : 0;
                        if (cmp != 0) {
                            return sd.ascending ? +cmp : -cmp;
                        }
                    }
                    return 0;
                };
            };
            // returns an array filtered using the current filter definition
            CollectionView.prototype._performFilter = function (items) {
                return this.canFilter && this._filter
                    ? items.filter(this._filter, this)
                    : items;
            };
            /**
             * Raises the @see:currentChanged event.
             */
            CollectionView.prototype.onCurrentChanged = function (e) {
                if (e === void 0) { e = wijmo.EventArgs.empty; }
                this.currentChanged.raise(this, e);
            };
            /**
             * Raises the @see:currentChanging event.
             *
             * @param e @see:CancelEventArgs that contains the event data.
             */
            CollectionView.prototype.onCurrentChanging = function (e) {
                this.currentChanging.raise(this, e);
                return !e.cancel;
            };
            Object.defineProperty(CollectionView.prototype, "items", {
                /**
                 * Gets items in the view.
                 */
                get: function () {
                    return this._pgView;
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Suspend refreshes until the next call to @see:endUpdate.
             */
            CollectionView.prototype.beginUpdate = function () {
                this._updating++;
            };
            /**
             * Resume refreshes suspended by a call to @see:beginUpdate.
             */
            CollectionView.prototype.endUpdate = function () {
                this._updating--;
                if (this._updating <= 0) {
                    this.refresh();
                }
            };
            Object.defineProperty(CollectionView.prototype, "isUpdating", {
                /**
                 * Gets a value that indicates whether notifications are currently suspended
                 * (see @see:beginUpdate and @see:endUpdate).
                 */
                get: function () {
                    return this._updating > 0;
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Executes a function within a @see:beginUpdate/@see:endUpdate block.
             *
             * The collection will not be refreshed until the function finishes.
             * This method ensures @see:endUpdate is called even if the function throws
             * an exception.
             *
             * @param fn Function to be executed without updates.
             */
            CollectionView.prototype.deferUpdate = function (fn) {
                try {
                    this.beginUpdate();
                    fn();
                }
                finally {
                    this.endUpdate();
                }
            };
            Object.defineProperty(CollectionView.prototype, "canAddNew", {
                // ** IEditableCollectionView
                /**
                 * Gets a value that indicates whether a new item can be added to the collection.
                 */
                get: function () {
                    return this._canAddNew;
                },
                set: function (value) {
                    this._canAddNew = wijmo.asBoolean(value);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "canCancelEdit", {
                /**
                 * Gets a value that indicates whether the collection view can discard pending changes
                 * and restore the original values of an edited object.
                 */
                get: function () {
                    return this._canCancelEdit;
                },
                set: function (value) {
                    this._canCancelEdit = wijmo.asBoolean(value);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "canRemove", {
                /**
                 * Gets a value that indicates whether items can be removed from the collection.
                 */
                get: function () {
                    return this._canRemove;
                },
                set: function (value) {
                    this._canRemove = wijmo.asBoolean(value);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "currentAddItem", {
                /**
                 * Gets the item that is being added during the current add transaction.
                 */
                get: function () {
                    return this._newItem;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "currentEditItem", {
                /**
                 * Gets the item that is being edited during the current edit transaction.
                 */
                get: function () {
                    return this._edtItem;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "isAddingNew", {
                /**
                 * Gets a value that indicates whether an add transaction is in progress.
                 */
                get: function () {
                    return this._newItem != null;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "isEditingItem", {
                /**
                 * Gets a value that indicates whether an edit transaction is in progress.
                 */
                get: function () {
                    return this._edtItem != null;
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Creates a new item and adds it to the collection.
             *
             * This method takes no parameters. It creates a new item, adds it to the
             * collection, and prevents refresh operations until the new item is
             * committed using the @see:commitNew method or canceled using the
             * @see:cancelNew method.
             *
             * The code below shows how the @see:addNew method is typically used:
             *
             * <pre>
             * // create the new item, add it to the collection
             * var newItem = view.addNew();
             * // initialize the new item
             * newItem.id = getFreshId();
             * newItem.name = 'New Customer';
             * // commit the new item so the view can be refreshed
             * view.commitNew();
             * </pre>
             *
             * You can also add new items by pushing them into the @see:sourceCollection
             * and then calling the @see:refresh method. The main advantage of @see:addNew
             * is in user-interactive scenarios (like adding new items in a data grid),
             * because it gives users the ability to cancel the add operation. It also
             * prevents the new item from being sorted or filtered out of view until the
             * add operation is committed.
             *
             * @return The item that was added to the collection.
             */
            CollectionView.prototype.addNew = function () {
                // sanity
                if (arguments.length > 0) {
                    wijmo.assert(false, 'addNew does not take any parameters, it creates the new items.');
                }
                // commit pending changes
                this.commitEdit();
                this.commitNew();
                // honor canAddNew
                if (!this.canAddNew) {
                    wijmo.assert(false, 'cannot add items (canAddNew == false).');
                    return null;
                }
                // create new item
                var item = null, src = this.sourceCollection;
                if (this.newItemCreator) {
                    item = this.newItemCreator();
                }
                else if (src && src.length) {
                    item = new src[0].constructor();
                }
                else {
                    item = {};
                }
                if (item != null) {
                    // remember the new item
                    this._newItem = item;
                    // add the new item to the collection
                    this._updating++;
                    this._src.push(item); // **
                    this._updating--;
                    // add the new item to the bottom of the current view
                    if (this._pgView != this._src) {
                        this._pgView.push(item);
                    }
                    // add the new item to the last group and to the data items
                    if (this.groups && this.groups.length) {
                        var g = this.groups[this.groups.length - 1];
                        g.items.push(item);
                        while (g.groups && g.groups.length) {
                            g = g.groups[g.groups.length - 1];
                            g.items.push(item);
                        }
                    }
                    // notify listeners
                    this._raiseCollectionChanged(collections.NotifyCollectionChangedAction.Add, item, this._pgView.length - 1);
                    // select the new item
                    this.moveCurrentTo(item);
                }
                // done
                return this._newItem;
            };
            /**
             * Ends the current edit transaction and, if possible,
             * restores the original value to the item.
             */
            CollectionView.prototype.cancelEdit = function () {
                var item = this._edtItem;
                if (item != null) {
                    this._edtItem = null;
                    // honor canCancelEdit
                    if (!this.canCancelEdit) {
                        wijmo.assert(false, 'cannot cancel edits (canCancelEdit == false).');
                        return;
                    }
                    // check that we can do this (TFS 110168)
                    var index = this._src.indexOf(item);
                    if (index < 0 || !this._edtClone) {
                        return;
                    }
                    // restore original item value
                    this._extend(this._src[index], this._edtClone);
                    this._edtClone = null;
                    // notify listeners
                    this._canceling = true;
                    this._raiseCollectionChanged(collections.NotifyCollectionChangedAction.Change, item, index);
                    this._canceling = false;
                }
            };
            /**
             * Ends the current add transaction and discards the pending new item.
             */
            CollectionView.prototype.cancelNew = function () {
                var item = this._newItem;
                if (item != null) {
                    this.remove(item);
                }
            };
            /**
             * Ends the current edit transaction and saves the pending changes.
             */
            CollectionView.prototype.commitEdit = function () {
                var item = this._edtItem;
                if (item != null) {
                    // start committing
                    this._committing = true;
                    // check if anything really changed
                    var sameContent = this._sameContent(item, this._edtClone);
                    // clean up state
                    this._edtItem = null;
                    this._edtClone = null;
                    // refresh to update the edited item
                    var index = this._pgView.indexOf(item);
                    var digest = this._digest;
                    this._performRefresh();
                    // track changes (before notifying)
                    if (!sameContent) {
                        this._trackItemChanged(item);
                    }
                    // notify (single item change or full refresh)
                    if (this._pgView.indexOf(item) == index && digest == this._digest) {
                        this._raiseCollectionChanged(collections.NotifyCollectionChangedAction.Change, item, index);
                    }
                    else {
                        this._raiseCollectionChanged(); // full refresh
                    }
                    // done committing
                    this._committing = false;
                }
            };
            /**
             * Ends the current add transaction and saves the pending new item.
             */
            CollectionView.prototype.commitNew = function () {
                var item = this._newItem;
                if (item != null) {
                    // clean up state
                    this._newItem = null;
                    // refresh to update the new item
                    var index = this._pgView.indexOf(item);
                    var digest = this._digest;
                    this._performRefresh();
                    // track changes (before notifying)
                    if (this._trackChanges == true) {
                        var idx = this._chgEdited.indexOf(item);
                        if (idx > -1) {
                            this._chgEdited.removeAt(idx);
                        }
                        if (this._chgAdded.indexOf(item) < 0) {
                            this._chgAdded.push(item);
                        }
                    }
                    // notify (full refresh if the item moved)
                    if (this._pgView.indexOf(item) == index && digest == this._digest) {
                        this._raiseCollectionChanged(collections.NotifyCollectionChangedAction.Change, item, index);
                    }
                    else {
                        this._raiseCollectionChanged(); // full refresh
                    }
                }
            };
            /**
             * Begins an edit transaction of the specified item.
             *
             * @param item Item to be edited.
             */
            CollectionView.prototype.editItem = function (item) {
                // commit pending changes if not already editing/adding this item
                if (item != this._edtItem && this.moveCurrentTo(item)) {
                    this.commitEdit();
                    this._edtItem = item;
                    this._edtClone = {};
                    this._extend(this._edtClone, this._edtItem);
                }
            };
            /**
             * Removes the specified item from the collection.
             *
             * @param item Item to be removed from the collection.
             */
            CollectionView.prototype.remove = function (item) {
                // handle cases where the user is adding or editing items
                var pendingNew = (item == this._newItem);
                if (pendingNew) {
                    this._newItem = null;
                }
                if (item == this._edtItem) {
                    this.cancelEdit();
                }
                // honor canRemove
                if (!this.canRemove) {
                    wijmo.assert(false, 'cannot remove items (canRemove == false).');
                    return;
                }
                // find item
                var index = this._src.indexOf(item);
                if (index > -1) {
                    // get current item to notify later
                    var current = this.currentItem;
                    // remove item from source collection
                    this._updating++;
                    this._src.splice(index, 1); // **
                    this._updating--;
                    // refresh to update the removed item
                    //var index = this._pgView.indexOf(item);
                    var digest = this._digest;
                    this._performRefresh();
                    // track changes (before notifying)
                    if (this._trackChanges == true) {
                        // removing something that was added
                        var idxAdded = this._chgAdded.indexOf(item);
                        if (idxAdded > -1) {
                            this._chgAdded.removeAt(idxAdded);
                        }
                        // removing something that was edited
                        var idxEdited = this._chgEdited.indexOf(item);
                        if (idxEdited > -1) {
                            this._chgEdited.removeAt(idxEdited);
                        }
                        // add to removed list unless it was pending and not added in this session
                        var idxRemoved = this._chgRemoved.indexOf(item);
                        if (idxRemoved < 0 && !pendingNew && idxAdded < 0) {
                            this._chgRemoved.push(item);
                        }
                    }
                    // notify (item removed or full refresh) (TFS 85001)
                    var sorted = this.sortDescriptions.length > 0, // JavaScript sort is not stable...
                    paged = this.pageSize > 0 && this._pgIdx > -1;
                    if (sorted || paged || digest != this._getGroupsDigest(this.groups)) {
                        this._raiseCollectionChanged();
                    }
                    else {
                        this._raiseCollectionChanged(collections.NotifyCollectionChangedAction.Remove, item, index);
                    }
                    // raise currentChanged if needed
                    if (this.currentItem !== current) {
                        this.onCurrentChanged();
                    }
                }
            };
            /**
             * Removes the item at the specified index from the collection.
             *
             * @param index Index of the item to be removed from the collection.
             * The index is relative to the view, not to the source collection.
             */
            CollectionView.prototype.removeAt = function (index) {
                index = wijmo.asInt(index);
                this.remove(this._pgView[index]);
            };
            // track changes applied to an item (not necessarily the current edit item)
            CollectionView.prototype._trackItemChanged = function (item) {
                if (this._trackChanges) {
                    var idx = this._chgEdited.indexOf(item);
                    if (idx < 0 && this._chgAdded.indexOf(item) < 0) {
                        this._chgEdited.push(item);
                    }
                    else if (idx > -1) {
                        var e = new collections.NotifyCollectionChangedEventArgs(collections.NotifyCollectionChangedAction.Change, item, idx);
                        this._chgEdited.onCollectionChanged(e);
                    }
                    else {
                        idx = this._chgAdded.indexOf(item);
                        if (idx > -1) {
                            var e = new collections.NotifyCollectionChangedEventArgs(collections.NotifyCollectionChangedAction.Change, item, idx);
                            this._chgAdded.onCollectionChanged(e);
                        }
                    }
                }
            };
            // extends an object (shallow copy)
            CollectionView.prototype._extend = function (dst, src) {
                for (var key in src) {
                    dst[key] = src[key];
                }
            };
            // checks whether two objects have the same content
            CollectionView.prototype._sameContent = function (dst, src) {
                for (var key in src) {
                    if (!this._sameValue(dst[key], src[key])) {
                        return false;
                    }
                }
                for (var key in dst) {
                    if (!this._sameValue(dst[key], src[key])) {
                        return false;
                    }
                }
                return true;
            };
            // checks whether two values are the same
            CollectionView.prototype._sameValue = function (v1, v2) {
                return v1 == v2 || wijmo.DateTime.equals(v1, v2);
            };
            Object.defineProperty(CollectionView.prototype, "canChangePage", {
                // ** IPagedCollectionView
                /**
                 * Gets a value that indicates whether the @see:pageIndex value can change.
                 */
                get: function () {
                    return this._canChangePage;
                },
                set: function (value) {
                    this._canChangePage = wijmo.asBoolean(value);
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "isPageChanging", {
                /**
                 * Gets a value that indicates whether the page index is changing.
                 */
                get: function () {
                    return false;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "itemCount", {
                /**
                 * Gets the total number of items in the view taking paging into account.
                 */
                get: function () {
                    return this._pgView.length;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "pageIndex", {
                /**
                 * Gets the zero-based index of the current page.
                 */
                get: function () {
                    return this._pgIdx;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "pageSize", {
                /**
                 * Gets or sets the number of items to display on a page.
                 */
                get: function () {
                    return this._pgSz;
                },
                set: function (value) {
                    if (value != this._pgSz) {
                        this._pgSz = wijmo.asInt(value);
                        this.refresh();
                    }
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "totalItemCount", {
                /**
                 * Gets the total number of items in the view before paging is applied.
                 */
                get: function () {
                    return this._view.length;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionView.prototype, "pageCount", {
                /**
                 * Gets the total number of pages.
                 */
                get: function () {
                    return this.pageSize ? Math.ceil(this.totalItemCount / this.pageSize) : 1;
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Sets the first page as the current page.
             *
             * @return True if the page index was changed successfully.
             */
            CollectionView.prototype.moveToFirstPage = function () {
                return this.moveToPage(0);
            };
            /**
             * Sets the last page as the current page.
             *
             * @return True if the page index was changed successfully.
             */
            CollectionView.prototype.moveToLastPage = function () {
                return this.moveToPage(this.pageCount - 1);
            };
            /**
             * Moves to the page after the current page.
             *
             * @return True if the page index was changed successfully.
             */
            CollectionView.prototype.moveToNextPage = function () {
                return this.moveToPage(this.pageIndex + 1);
            };
            /**
             * Moves to the page at the specified index.
             *
             * @param index Index of the page to move to.
             * @return True if the page index was changed successfully.
             */
            CollectionView.prototype.moveToPage = function (index) {
                var newIndex = wijmo.clamp(index, 0, this.pageCount - 1);
                if (newIndex != this._pgIdx) {
                    // honor canChangePage
                    if (!this.canChangePage) {
                        wijmo.assert(false, 'cannot change pages (canChangePage == false).');
                    }
                    // raise pageChanging
                    var e = new collections.PageChangingEventArgs(newIndex);
                    if (this.onPageChanging(e)) {
                        // change the page
                        this._pgIdx = newIndex;
                        this._pgView = this._getPageView();
                        this._idx = 0;
                        // raise pageChanged and collectionChanged, or refresh if grouping
                        if (!this.groupDescriptions || this.groupDescriptions.length == 0) {
                            this.onPageChanged();
                            this.onCollectionChanged();
                        }
                        else {
                            this.refresh();
                        }
                    }
                }
                return this._pgIdx == index;
            };
            /**
             * Moves to the page before the current page.
             *
             * @return True if the page index was changed successfully.
             */
            CollectionView.prototype.moveToPreviousPage = function () {
                return this.moveToPage(this.pageIndex - 1);
            };
            /**
             * Raises the @see:pageChanged event.
             */
            CollectionView.prototype.onPageChanged = function (e) {
                if (e === void 0) { e = wijmo.EventArgs.empty; }
                this.pageChanged.raise(this, e);
            };
            /**
             * Raises the @see:pageChanging event.
             *
             * @param e @see:PageChangingEventArgs that contains the event data.
             */
            CollectionView.prototype.onPageChanging = function (e) {
                this.pageChanging.raise(this, e);
                return !e.cancel;
            };
            // gets the full group that corresponds to a paged group view
            CollectionView.prototype._getFullGroup = function (g) {
                // look for the group by level and name
                // this gets the full (unpaged) and updated group (TFS 109119)
                var fg = this._getGroupByPath(this._fullGroups, g.level, g._path);
                if (fg != null) {
                    g = fg;
                }
                // return the group
                return g;
            };
            // gets a group from a collection by path
            CollectionView.prototype._getGroupByPath = function (groups, level, path) {
                for (var i = 0; i < groups.length; i++) {
                    var g = groups[i];
                    if (g.level == level && g._path == path) {
                        return g;
                    }
                    if (g.level < level && path.indexOf(g._path) == 0) {
                        g = this._getGroupByPath(g.groups, level, path);
                        if (g != null) {
                            return g;
                        }
                    }
                }
                return null;
            };
            // gets the list that corresponds to the current page
            CollectionView.prototype._getPageView = function () {
                // not paging? return the whole view
                if (this.pageSize <= 0 || this._pgIdx < 0) {
                    return this._view;
                }
                // slice the current page out of the view
                var start = this._pgSz * this._pgIdx, end = Math.min(start + this._pgSz, this._view.length);
                return this._view.slice(start, end);
            };
            // creates a grouped view of the current page
            CollectionView.prototype._createGroups = function (items) {
                // not grouping? return null
                if (!this._grpDesc || !this._grpDesc.length) {
                    return null;
                }
                // build group tree
                var root = [], maps = {}, map = null;
                for (var i = 0; i < items.length; i++) {
                    // get the item
                    var item = items[i], groups = root, levels = this._grpDesc.length;
                    // add this item to the tree
                    var path = '';
                    for (var level = 0; level < levels; level++) {
                        // get the group name for this level
                        var gd = this._grpDesc[level], name = gd.groupNameFromItem(item, level), last = level == levels - 1;
                        // get the group map for this level (optimization)
                        map = maps[path];
                        if (!map && wijmo.isPrimitive(name)) {
                            map = {};
                            maps[path] = map;
                        }
                        // get or create the group
                        var group = this._getGroup(gd, groups, map, name, level, last);
                        // keep group path (all names in the hierarchy)
                        path += '/' + name;
                        group._path = path;
                        // add data items to last level groups
                        if (last) {
                            group.items.push(item);
                        }
                        // move on to the next group
                        groups = group.groups;
                    }
                }
                // done
                return root;
            };
            // gets a string digest of the current groups 
            // this is used to check whether changes require a full refresh
            CollectionView.prototype._getGroupsDigest = function (groups) {
                var digest = '';
                for (var i = 0; groups != null && i < groups.length; i++) {
                    var g = groups[i];
                    digest += '{' + g.name + ':' + (g.items ? g.items.length : '*');
                    if (g.groups.length > 0) {
                        digest += ',';
                        digest += this._getGroupsDigest(g.groups);
                    }
                    digest += '}';
                }
                return digest;
            };
            // gets an array that contains all the children for a list of groups
            // NOTE: use "push.apply" instead of "concat" for much better performance
            // NOTE2: use explicit loop for even better performance and to avoid stack overflows (TFS 15921)
            CollectionView.prototype._mergeGroupItems = function (groups) {
                var items = [];
                for (var i = 0; i < groups.length; i++) {
                    var g = groups[i];
                    if (!g._isBottomLevel) {
                        var groupItems = this._mergeGroupItems(g.groups);
                        //g._items.push.apply(g._items, groupItems);
                        for (var a = 0, len = groupItems.length; a < len; a++) {
                            g._items.push(groupItems[a]);
                        }
                    }
                    //items.push.apply(items, g._items);
                    for (var a = 0, len = g._items.length; a < len; a++) {
                        items.push(g._items[a]);
                    }
                }
                return items;
            };
            // finds or creates a group
            CollectionView.prototype._getGroup = function (gd, groups, map, name, level, isBottomLevel) {
                var g;
                // find existing group
                if (map && wijmo.isPrimitive(name)) {
                    g = map[name];
                    if (g) {
                        return g;
                    }
                }
                else {
                    for (var i = 0; i < groups.length; i++) {
                        if (gd.namesMatch(groups[i].name, name)) {
                            return groups[i];
                        }
                    }
                }
                // not found, create now
                var group = new CollectionViewGroup(gd, name, level, isBottomLevel);
                groups.push(group);
                // add group to map
                if (map) {
                    map[name] = group;
                }
                // done
                return group;
            };
            return CollectionView;
        }());
        collections.CollectionView = CollectionView;
        /**
         * Represents a group created by a @see:CollectionView object based on
         * its @see:CollectionView.groupDescriptions property.
         */
        var CollectionViewGroup = (function () {
            /**
             * Initializes a new instance of the @see:CollectionViewGroup class.
             *
             * @param groupDescription @see:GroupDescription that owns the new group.
             * @param name Name of the new group.
             * @param level Level of the new group.
             * @param isBottomLevel Whether this group has any subgroups.
             */
            function CollectionViewGroup(groupDescription, name, level, isBottomLevel) {
                this._gd = groupDescription;
                this._name = name;
                this._level = level;
                this._isBottomLevel = isBottomLevel;
                this._groups = [];
                this._items = [];
            }
            Object.defineProperty(CollectionViewGroup.prototype, "name", {
                /*
                 * Gets the name of this group.
                 */
                get: function () {
                    return this._name;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionViewGroup.prototype, "level", {
                /*
                 * Gets the level of this group.
                 */
                get: function () {
                    return this._level;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionViewGroup.prototype, "isBottomLevel", {
                /*
                 * Gets a value that indicates whether this group has any subgroups.
                 */
                get: function () {
                    return this._isBottomLevel;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionViewGroup.prototype, "items", {
                /*
                 * Gets an array containing the items included in this group (including all subgroups).
                 */
                get: function () {
                    return this._items;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionViewGroup.prototype, "groups", {
                /*
                 * Gets an array containing the this group's subgroups.
                 */
                get: function () {
                    return this._groups;
                },
                enumerable: true,
                configurable: true
            });
            Object.defineProperty(CollectionViewGroup.prototype, "groupDescription", {
                /*
                 * Gets the @see:GroupDescription that owns this group.
                 */
                get: function () {
                    return this._gd;
                },
                enumerable: true,
                configurable: true
            });
            /**
             * Calculates an aggregate value for the items in this group.
             *
             * @param aggType Type of aggregate to calculate.
             * @param binding Property to aggregate on.
             * @param view CollectionView that owns this group.
             * @return The aggregate value.
             */
            CollectionViewGroup.prototype.getAggregate = function (aggType, binding, view) {
                var cv = wijmo.tryCast(view, CollectionView), group = cv ? cv._getFullGroup(this) : this;
                return wijmo.getAggregate(aggType, group.items, binding);
            };
            return CollectionViewGroup;
        }());
        collections.CollectionViewGroup = CollectionViewGroup;
    })(collections = wijmo.collections || (wijmo.collections = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=CollectionView.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    'use strict';
    /**
     * Provides a pop-up window that displays additional information about elements on the page.
     *
     * The @see:Tooltip class can be used in two modes:
     *
     * <b>Automatic Mode:</b> Use the @see:setTooltip method to connect the @see:Tooltip to
     * one or more elements on the page. The @see:Tooltip will automatically monitor events
     * and display the tooltips when the user performs actions that trigger the tooltip.
     * For example:
     *
     * <pre>var tt = new wijmo.Tooltip();
     * tt.setTooltip('#menu', 'Select commands.');
     * tt.setTooltip('#tree', 'Explore the hierarchy.');
     * tt.setTooltip('#chart', '#idChartTooltip');</pre>
     *
     * <b>Manual Mode:</b> The caller is responsible for showing and hiding the tooltip
     * using the @see:show and @see:hide methods. For example:
     *
     * <pre>var tt = new wijmo.Tooltip();
     * element.addEventListener('click', function () {
     *   if (tt.isVisible) {
     *     tt.hide();
     *   } else {
     *     tt.show(element, 'This is an important element!');
     *   }
     * });</pre>
     */
    var Tooltip = (function () {
        /**
         * Initializes a new instance of the @see:Tooltip class.
         *
         * @param options JavaScript object containing initialization data for the @see:Tooltip.
         */
        function Tooltip(options) {
            this._showAutoTipBnd = this._showAutoTip.bind(this);
            this._hideAutoTipBnd = this._hideAutoTip.bind(this);
            // property storage
            this._html = true;
            this._gap = 6;
            this._showAtMouse = false;
            this._showDelay = 500; // http://msdn.microsoft.com/en-us/library/windows/desktop/bb760404(v=vs.85).aspx
            this._hideDelay = 0; // do not hide
            this._tips = [];
            /**
             * Occurs before the tooltip content is displayed.
             *
             * The event handler may customize the tooltip content or suppress the
             * tooltip display by changing the event parameters.
             */
            this.popup = new wijmo.Event();
            wijmo.copy(this, options);
        }
        // object model
        /**
         * Assigns tooltip content to a given element on the page.
         *
         * The same tooltip may be used to display information for any number
         * of elements on the page. To remove the tooltip from an element,
         * call @see:setTooltip and specify null for the content.
         *
         * @param element Element, element ID, or control that the tooltip explains.
         * @param content Tooltip content or ID of the element that contains the tooltip content.
         */
        Tooltip.prototype.setTooltip = function (element, content) {
            // get element and tooltip content
            element = wijmo.getElement(element);
            content = this._getContent(content);
            // remove old version from list
            var i = this._indexOf(element);
            if (i > -1) {
                this._detach(element);
                this._tips.splice(i, 1);
            }
            // add new version to list
            if (content) {
                this._attach(element);
                this._tips.push({ element: element, content: content });
            }
        };
        /**
         * Shows the tooltip with the specified content, next to the specified element.
         *
         * @param element Element, element ID, or control that the tooltip explains.
         * @param content Tooltip content or ID of the element that contains the tooltip content.
         * @param bounds Optional element that defines the bounds of the area that the tooltip
         * targets. If not provided, the bounds of the element are used (as reported by the
         * <b>getBoundingClientRect</b> method).
         */
        Tooltip.prototype.show = function (element, content, bounds) {
            // get element and tooltip content
            element = wijmo.getElement(element);
            content = this._getContent(content);
            if (!bounds) {
                bounds = wijmo.Rect.fromBoundingRect(element.getBoundingClientRect());
            }
            // create tooltip element if necessary
            var tip = Tooltip._eTip;
            if (!tip) {
                tip = Tooltip._eTip = document.createElement('div');
                wijmo.addClass(tip, 'wj-tooltip');
                tip.style.visibility = 'none';
                document.body.appendChild(tip);
            }
            // set tooltip content
            this._setContent(content);
            // fire event to allow customization
            var e = new TooltipEventArgs(content);
            this.onPopup(e);
            // if not canceled and content is present, show tooltip
            if (e.content && !e.cancel) {
                // update tooltip content with customize content, if any
                this._setContent(e.content);
                tip.style.minWidth = '';
                // apply gap and align to the center of the reference element
                bounds = new wijmo.Rect(bounds.left - (tip.offsetWidth - bounds.width) / 2, bounds.top - this.gap, tip.offsetWidth, bounds.height + 2 * this.gap);
                // show tooltip
                wijmo.showPopup(tip, bounds, true);
                // hide when the mouse goes down
                document.addEventListener('mousedown', this._hideAutoTipBnd);
            }
        };
        /**
         * Hides the tooltip if it is currently visible.
         */
        Tooltip.prototype.hide = function () {
            if (Tooltip._eTip) {
                Tooltip._eTip.style.visibility = 'hidden';
                Tooltip._eTip.innerHTML = '';
            }
            document.removeEventListener('mousedown', this._hideAutoTipBnd);
        };
        Object.defineProperty(Tooltip.prototype, "isVisible", {
            /**
             * Gets whether the tooltip is currently visible.
             */
            get: function () {
                return Tooltip._eTip && Tooltip._eTip.style.visibility != 'hidden';
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Tooltip.prototype, "isContentHtml", {
            /**
             * Gets or sets a value that determines whether the tooltip contents
             * should be displayed as plain text or as HTML.
             */
            get: function () {
                return this._html;
            },
            set: function (value) {
                this._html = wijmo.asBoolean(value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Tooltip.prototype, "gap", {
            /**
             * Gets or sets the distance between the tooltip and the target element.
             */
            get: function () {
                return this._gap;
            },
            set: function (value) {
                this._gap = wijmo.asNumber(value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Tooltip.prototype, "showAtMouse", {
            /**
             * Gets or sets a value that determines whether the tooltip should be
             * positioned with respect to the mouse position rather than the
             * target element.
             */
            get: function () {
                return this._showAtMouse;
            },
            set: function (value) {
                this._showAtMouse = wijmo.asBoolean(value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Tooltip.prototype, "showDelay", {
            /**
             * Gets or sets the delay, in milliseconds, before showing the tooltip after the
             * mouse enters the target element.
             */
            get: function () {
                return this._showDelay;
            },
            set: function (value) {
                this._showDelay = wijmo.asInt(value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Tooltip.prototype, "hideDelay", {
            /**
             * Gets or sets the delay, in milliseconds, before hiding the tooltip after the
             * mouse leaves the target element.
             */
            get: function () {
                return this._hideDelay;
            },
            set: function (value) {
                this._hideDelay = wijmo.asInt(value);
            },
            enumerable: true,
            configurable: true
        });
        /**
         * Raises the @see:popup event.
         *
         * @param e @see:TooltipEventArgs that contains the event data.
         */
        Tooltip.prototype.onPopup = function (e) {
            if (this.popup) {
                this.popup.raise(this, e);
            }
        };
        // implementation
        // finds an element in the auto-tooltip list
        Tooltip.prototype._indexOf = function (e) {
            for (var i = 0; i < this._tips.length; i++) {
                if (this._tips[i].element == e) {
                    return i;
                }
            }
            return -1;
        };
        // add event listeners to show and hide tooltips for an element
        Tooltip.prototype._attach = function (e) {
            e.addEventListener('mouseenter', this._showAutoTipBnd);
            e.addEventListener('mouseleave', this._hideAutoTipBnd);
            e.addEventListener('click', this._showAutoTipBnd);
        };
        // remove event listeners used to show and hide tooltips for an element
        Tooltip.prototype._detach = function (e) {
            e.removeEventListener('mouseenter', this._showAutoTipBnd);
            e.removeEventListener('mouseleave', this._hideAutoTipBnd);
            e.removeEventListener('click', this._showAutoTipBnd);
        };
        // shows an automatic tooltip
        Tooltip.prototype._showAutoTip = function (e) {
            var _this = this;
            var showDelay = e.type == 'mouseenter' ? this._showDelay : 0;
            this._clearTimeouts();
            this._toShow = setTimeout(function () {
                var i = _this._indexOf(e.target);
                if (i > -1) {
                    var tip = _this._tips[i], bounds = _this._showAtMouse ? new wijmo.Rect(e.clientX, e.clientY, 0, 0) : null;
                    _this.show(tip.element, tip.content, bounds);
                    if (_this._hideDelay > 0) {
                        _this._toHide = setTimeout(function () {
                            _this.hide();
                        }, _this._hideDelay);
                    }
                }
            }, showDelay);
        };
        // hides an automatic tooltip
        Tooltip.prototype._hideAutoTip = function () {
            this._clearTimeouts();
            this.hide();
        };
        // clears the timeouts used to show and hide automatic tooltips
        Tooltip.prototype._clearTimeouts = function () {
            if (this._toShow) {
                clearTimeout(this._toShow);
                this._toShow = null;
            }
            if (this._toHide) {
                clearTimeout(this._toHide);
                this._toHide = null;
            }
        };
        // gets content that may be a string or an element id
        Tooltip.prototype._getContent = function (content) {
            content = wijmo.asString(content);
            if (content && content[0] == '#') {
                var e = wijmo.getElement(content);
                if (e) {
                    content = e.innerHTML;
                }
            }
            return content;
        };
        // assigns content to the tooltip element
        Tooltip.prototype._setContent = function (content) {
            var tip = Tooltip._eTip;
            if (tip) {
                if (this.isContentHtml) {
                    tip.innerHTML = content;
                }
                else {
                    tip.textContent = content;
                }
            }
        };
        return Tooltip;
    }());
    wijmo.Tooltip = Tooltip;
    // helper class to hold element/tooltip information
    var ElementContent = (function () {
        function ElementContent() {
        }
        return ElementContent;
    }());
    /**
     * Provides arguments for the @see:Tooltip.popup event.
     */
    var TooltipEventArgs = (function (_super) {
        __extends(TooltipEventArgs, _super);
        /**
         * Initializes a new instance of the @see:TooltipEventArgs class.
         *
         * @param content String to show in the tooltip.
         */
        function TooltipEventArgs(content) {
            _super.call(this);
            this._content = wijmo.asString(content);
        }
        Object.defineProperty(TooltipEventArgs.prototype, "content", {
            /**
             * Gets or sets the content to show in the tooltip.
             *
             * This parameter can be used while handling the @see:Tooltip.popup
             * event to modify the content of the tooltip.
             */
            get: function () {
                return this._content;
            },
            set: function (value) {
                this._content = wijmo.asString(value);
            },
            enumerable: true,
            configurable: true
        });
        return TooltipEventArgs;
    }(wijmo.CancelEventArgs));
    wijmo.TooltipEventArgs = TooltipEventArgs;
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Tooltip.js.map
var wijmo;
(function (wijmo) {
    'use strict';
    /**
     * Represents a color.
     *
     * The @see:Color class parses colors specified as CSS strings and exposes
     * their red, green, blue, and alpha channels as read-write properties.
     *
     * The @see:Color class also provides @see:fromHsb and @see:fromHsl methods
     * for creating colors using the HSB and HSL color models instead of RGB,
     * as well as @see:getHsb and @see:getHsl methods for retrieving the color
     * components using those color models.
     *
     * Finally, the @see:Color class provides an @see:interpolate method that
     * creates colors by interpolating between two colors using the HSL model.
     * This method is especially useful for creating color animations with the
     * @see:animate method.
     */
    var Color = (function () {
        /**
         * Initializes a new @see:Color from a CSS color specification.
         *
         * @param color CSS color specification.
         */
        function Color(color) {
            // fields
            this._r = 0;
            this._g = 0;
            this._b = 0;
            this._a = 1;
            if (color) {
                this._parse(color);
            }
        }
        Object.defineProperty(Color.prototype, "r", {
            /**
             * Gets or sets the red component of this @see:Color,
             * in a range from 0 to 255.
             */
            get: function () {
                return this._r;
            },
            set: function (value) {
                this._r = wijmo.clamp(wijmo.asNumber(value), 0, 255);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Color.prototype, "g", {
            /**
             * Gets or sets the green component of this @see:Color,
             * in a range from 0 to 255.
             */
            get: function () {
                return this._g;
            },
            set: function (value) {
                this._g = wijmo.clamp(wijmo.asNumber(value), 0, 255);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Color.prototype, "b", {
            /**
             * Gets or sets the blue component of this @see:Color,
             * in a range from 0 to 255.
             */
            get: function () {
                return this._b;
            },
            set: function (value) {
                this._b = wijmo.clamp(wijmo.asNumber(value), 0, 255);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Color.prototype, "a", {
            /**
             * Gets or sets the alpha component of this @see:Color,
             * in a range from 0 to 1 (zero is transparent, one is solid).
             */
            get: function () {
                return this._a;
            },
            set: function (value) {
                this._a = wijmo.clamp(wijmo.asNumber(value), 0, 1);
            },
            enumerable: true,
            configurable: true
        });
        /**
         * Returns true if a @see:Color has the same value as this @see:Color.
         *
         * @param clr @see:Color to compare to this @see:Color.
         */
        Color.prototype.equals = function (clr) {
            return (clr instanceof Color) &&
                this.r == clr.r && this.g == clr.g && this.b == clr.b &&
                this.a == clr.a;
        };
        /**
         * Gets a string representation of this @see:Color.
         */
        Color.prototype.toString = function () {
            var a = Math.round(this.a * 100);
            return a > 99
                ? '#' + ((1 << 24) + (this.r << 16) + (this.g << 8) + this.b).toString(16).slice(1)
                : 'rgba(' + this.r + ',' + this.g + ',' + this.b + ',' + (a / 100) + ')';
        };
        /**
         * Creates a new @see:Color using the specified RGBA color channel values.
         *
         * @param r Value for the red channel, from 0 to 255.
         * @param g Value for the green channel, from 0 to 255.
         * @param b Value for the blue channel, from 0 to 255.
         * @param a Value for the alpha channel, from 0 to 1.
         */
        Color.fromRgba = function (r, g, b, a) {
            if (a === void 0) { a = 1; }
            var c = new Color(null);
            c.r = Math.round(wijmo.clamp(wijmo.asNumber(r), 0, 255));
            c.g = Math.round(wijmo.clamp(wijmo.asNumber(g), 0, 255));
            c.b = Math.round(wijmo.clamp(wijmo.asNumber(b), 0, 255));
            c.a = wijmo.clamp(wijmo.asNumber(a), 0, 1);
            return c;
        };
        /**
         * Creates a new @see:Color using the specified HSB values.
         *
         * @param h Hue value, from 0 to 1.
         * @param s Saturation value, from 0 to 1.
         * @param b Brightness value, from 0 to 1.
         * @param a Alpha value, from 0 to 1.
         */
        Color.fromHsb = function (h, s, b, a) {
            if (a === void 0) { a = 1; }
            var rgb = Color._hsbToRgb(wijmo.clamp(wijmo.asNumber(h), 0, 1), wijmo.clamp(wijmo.asNumber(s), 0, 1), wijmo.clamp(wijmo.asNumber(b), 0, 1));
            return Color.fromRgba(rgb[0], rgb[1], rgb[2], a);
        };
        /**
         * Creates a new @see:Color using the specified HSL values.
         *
         * @param h Hue value, from 0 to 1.
         * @param s Saturation value, from 0 to 1.
         * @param l Lightness value, from 0 to 1.
         * @param a Alpha value, from 0 to 1.
         */
        Color.fromHsl = function (h, s, l, a) {
            if (a === void 0) { a = 1; }
            var rgb = Color._hslToRgb(wijmo.clamp(wijmo.asNumber(h), 0, 1), wijmo.clamp(wijmo.asNumber(s), 0, 1), wijmo.clamp(wijmo.asNumber(l), 0, 1));
            return Color.fromRgba(rgb[0], rgb[1], rgb[2], a);
        };
        /**
         * Creates a new @see:Color from a CSS color string.
         *
         * @param value String containing a CSS color specification.
         * @return A new @see:Color, or null if the string cannot be parsed into a color.
         */
        Color.fromString = function (value) {
            var c = new Color(null);
            return c._parse(wijmo.asString(value)) ? c : null;
        };
        /**
         * Gets an array with this color's HSB components.
         */
        Color.prototype.getHsb = function () {
            return Color._rgbToHsb(this.r, this.g, this.b);
        };
        /**
         * Gets an array with this color's HSL components.
         */
        Color.prototype.getHsl = function () {
            return Color._rgbToHsl(this.r, this.g, this.b);
        };
        /**
         * Creates a @see:Color by interpolating between two colors.
         *
         * @param c1 First color.
         * @param c2 Second color.
         * @param pct Value between zero and one that determines how close the
         * interpolation should be to the second color.
         */
        Color.interpolate = function (c1, c2, pct) {
            // sanity
            pct = wijmo.clamp(wijmo.asNumber(pct), 0, 1);
            // convert rgb to hsl
            var h1 = Color._rgbToHsl(c1.r, c1.g, c1.b), h2 = Color._rgbToHsl(c2.r, c2.g, c2.b);
            // interpolate
            var qct = 1 - pct, alpha = c1.a * qct + c2.a * pct, h3 = [
                h1[0] * qct + h2[0] * pct,
                h1[1] * qct + h2[1] * pct,
                h1[2] * qct + h2[2] * pct
            ];
            // convert back to rgb
            var rgb = Color._hslToRgb(h3[0], h3[1], h3[2]);
            return Color.fromRgba(rgb[0], rgb[1], rgb[2], alpha);
        };
        /**
         * Gets the closest opaque color to a given color.
         *
         * @param c @see:Color to be converted to an opaque color
         * (the color may also be specified as a string).
         * @param bkg Background color to use when removing the transparency
         * (defaults to white).
         */
        Color.toOpaque = function (c, bkg) {
            // get color
            c = wijmo.isString(c) ? Color.fromString(c) : wijmo.asType(c, Color);
            // no alpha? no work
            if (c.a == 1)
                return c;
            // get background
            bkg = bkg == null ? Color.fromRgba(255, 255, 255, 1) :
                wijmo.isString(bkg) ? Color.fromString(bkg) :
                    wijmo.asType(bkg, Color);
            // interpolate in RGB space
            var p = c.a, q = 1 - p;
            return Color.fromRgba(c.r * p + bkg.r * q, c.g * p + bkg.g * q, c.b * p + bkg.b * q);
        };
        // ** implementation
        // parses a color string into r/b/g/a
        Color.prototype._parse = function (c) {
            // case-insensitive
            c = c.toLowerCase();
            // 'transparent' is a special case:
            // https://developer.mozilla.org/en-US/docs/Web/CSS/color_value
            if (c == 'transparent') {
                this._r = this._g = this._b = this._a = 0;
                return true;
            }
            // let browser parse stuff we don't handle
            if (c && c.indexOf('#') != 0 && c.indexOf('rgb') != 0 && c.indexOf('hsl') != 0) {
                var e = document.createElement('div');
                e.style.color = c;
                var cc = e.style.color;
                if (cc == c) {
                    cc = window.getComputedStyle(e).color; // then get computed style
                    if (!cc) {
                        document.body.appendChild(e); // then add element to document
                        cc = window.getComputedStyle(e).color; // and try again
                        document.body.removeChild(e);
                    }
                }
                c = cc.toLowerCase();
            }
            // parse #RGB/#RRGGBB
            if (c.indexOf('#') == 0) {
                if (c.length == 4) {
                    this.r = parseInt(c[1] + c[1], 16);
                    this.g = parseInt(c[2] + c[2], 16);
                    this.b = parseInt(c[3] + c[3], 16);
                    this.a = 1;
                    return true;
                }
                else if (c.length == 7) {
                    this.r = parseInt(c.substr(1, 2), 16);
                    this.g = parseInt(c.substr(3, 2), 16);
                    this.b = parseInt(c.substr(5, 2), 16);
                    this.a = 1;
                    return true;
                }
                return false;
            }
            // parse rgb/rgba
            if (c.indexOf('rgb') == 0) {
                var op = c.indexOf('('), ep = c.indexOf(')');
                if (op > -1 && ep > -1) {
                    var p = c.substr(op + 1, ep - (op + 1)).split(',');
                    if (p.length > 2) {
                        this.r = parseInt(p[0]);
                        this.g = parseInt(p[1]);
                        this.b = parseInt(p[2]);
                        this.a = p.length > 3 ? parseFloat(p[3]) : 1;
                        return true;
                    }
                }
            }
            // parse hsl/hsla
            if (c.indexOf('hsl') == 0) {
                var op = c.indexOf('('), ep = c.indexOf(')');
                if (op > -1 && ep > -1) {
                    var p = c.substr(op + 1, ep - (op + 1)).split(',');
                    if (p.length > 2) {
                        var h = parseInt(p[0]) / 360, s = parseInt(p[1]), l = parseInt(p[2]);
                        if (p[1].indexOf('%') > -1)
                            s /= 100;
                        if (p[2].indexOf('%') > -1)
                            l /= 100;
                        var rgb = Color._hslToRgb(h, s, l);
                        this.r = rgb[0];
                        this.g = rgb[1];
                        this.b = rgb[2];
                        this.a = p.length > 3 ? parseFloat(p[3]) : 1;
                        return true;
                    }
                }
            }
            // failed to parse
            return false;
        };
        /**
         * Converts an HSL color value to RGB.
         *
         * @param h The hue (between zero and one).
         * @param s The saturation (between zero and one).
         * @param l The lightness (between zero and one).
         * @return An array containing the R, G, and B values (between zero and 255).
         */
        Color._hslToRgb = function (h, s, l) {
            wijmo.assert(h >= 0 && h <= 1 && s >= 0 && s <= 1 && l >= 0 && l <= 1, 'bad HSL values');
            var r, g, b;
            if (s == 0) {
                r = g = b = l; // achromatic
            }
            else {
                var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
                var p = 2 * l - q;
                r = Color._hue2rgb(p, q, h + 1 / 3);
                g = Color._hue2rgb(p, q, h);
                b = Color._hue2rgb(p, q, h - 1 / 3);
            }
            return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
        };
        Color._hue2rgb = function (p, q, t) {
            if (t < 0)
                t += 1;
            if (t > 1)
                t -= 1;
            if (t < 1 / 6)
                return p + (q - p) * 6 * t;
            if (t < 1 / 2)
                return q;
            if (t < 2 / 3)
                return p + (q - p) * (2 / 3 - t) * 6;
            return p;
        };
        /**
         * Converts an RGB color value to HSL.
         *
         * @param r The value of the red channel (between zero and 255).
         * @param g The value of the green channel (between zero and 255).
         * @param b The value of the blue channel (between zero and 255).
         * @return An array containing the H, S, and L values (between zero and one).
         */
        Color._rgbToHsl = function (r, g, b) {
            wijmo.assert(r >= 0 && r <= 255 && g >= 0 && g <= 255 && b >= 0 && b <= 255, 'bad RGB values');
            r /= 255, g /= 255, b /= 255;
            var max = Math.max(r, g, b), min = Math.min(r, g, b), h, s, l = (max + min) / 2;
            if (max == min) {
                h = s = 0;
            }
            else {
                var d = max - min;
                s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
                switch (max) {
                    case r:
                        h = (g - b) / d + (g < b ? 6 : 0);
                        break;
                    case g:
                        h = (b - r) / d + 2;
                        break;
                    case b:
                        h = (r - g) / d + 4;
                        break;
                }
                h /= 6;
            }
            return [h, s, l];
        };
        /**
         * Converts an RGB color value to HSB.
         *
         * @param r The value of the red channel (between zero and 255).
         * @param g The value of the green channel (between zero and 255).
         * @param b The value of the blue channel (between zero and 255).
         * @return An array containing the H, S, and B values (between zero and one).
         */
        Color._rgbToHsb = function (r, g, b) {
            wijmo.assert(r >= 0 && r <= 255 && g >= 0 && g <= 255 && b >= 0 && b <= 255, 'bad RGB values');
            var hsl = Color._rgbToHsl(r, g, b);
            return Color._hslToHsb(hsl[0], hsl[1], hsl[2]);
        };
        /**
         * Converts an HSB color value to RGB.
         *
         * @param h The hue (between zero and one).
         * @param s The saturation (between zero and one).
         * @param b The brightness (between zero and one).
         * @return An array containing the R, G, and B values (between zero and 255).
         */
        Color._hsbToRgb = function (h, s, b) {
            //assert(h >= 0 && h <= 1 && s >= 0 && s <= 1 && b >= 0 && b <= 1, 'bad HSB values');
            var hsl = Color._hsbToHsl(h, s, b);
            return Color._hslToRgb(hsl[0], hsl[1], hsl[2]);
        };
        /**
         * Converts an HSB color value to HSL.
         *
         * @param h The hue (between zero and one).
         * @param s The saturation (between zero and one).
         * @param b The brightness (between zero and one).
         * @return An array containing the H, S, and L values (between zero and one).
         */
        Color._hsbToHsl = function (h, s, b) {
            wijmo.assert(h >= 0 && h <= 1 && s >= 0 && s <= 1 && b >= 0 && b <= 1, 'bad HSB values');
            var ll = wijmo.clamp(b * (2 - s) / 2, 0, 1), div = 1 - Math.abs(2 * ll - 1), ss = wijmo.clamp(div > 0 ? b * s / div : s /*0*/, 0, 1);
            wijmo.assert(!isNaN(ll) && !isNaN(ss), 'bad conversion to HSL');
            return [h, ss, ll];
        };
        /**
         * Converts an HSL color value to HSB.
         *
         * @param h The hue (between zero and one).
         * @param s The saturation (between zero and one).
         * @param l The lightness (between zero and one).
         * @return An array containing the H, S, and B values (between zero and one).
         */
        Color._hslToHsb = function (h, s, l) {
            wijmo.assert(h >= 0 && h <= 1 && s >= 0 && s <= 1 && l >= 0 && l <= 1, 'bad HSL values');
            var bb = wijmo.clamp(l == 1 ? 1 : (2 * l + s * (1 - Math.abs(2 * l - 1))) / 2, 0, 1);
            var ss = wijmo.clamp(bb > 0 ? 2 * (bb - l) / bb : s /*0*/, 0, 1);
            wijmo.assert(!isNaN(bb) && !isNaN(ss), 'bad conversion to HSB');
            return [h, ss, bb];
        };
        return Color;
    }());
    wijmo.Color = Color;
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Color.js.map
var wijmo;
(function (wijmo) {
    'use strict';
    /**
     * Static class that provides utility methods for clipboard operations.
     *
     * The @see:Clipboard class provides static @see:copy and @see:paste methods
     * that can be used by controls to customize the clipboard content during
     * clipboard operations.
     *
     * For example, the code below shows how a control could intercept the
     * clipboard shortcut keys and provide custom clipboard handling:
     *
     * <pre>
     * rootElement.addEventListener('keydown', function(e) {
     *   // copy: ctrl+c or ctrl+Insert
     *   if (e.ctrlKey && (e.keyCode == 67 || e.keyCode == 45)) {
     *     var text = this.getClipString();
     *     Clipboard.copy(text);
     *     return;
     *   }
     *   // paste: ctrl+v or shift+Insert
     *   if ((e.ctrlKey && e.keyCode == 86) || (e.shiftKey && e.keyCode == 45)) {
     *     Clipboard.paste(function (text) {
     *       this.setClipString(text);
     *     });
     *     return;
     *   }
     * });</pre>
     */
    var Clipboard = (function () {
        function Clipboard() {
        }
        /**
         * Copies a string to the clipboard.
         *
         * This method only works if invoked immediately after the user
         * pressed a clipboard copy command (such as ctrl+c).
         *
         * @param text Text to copy to the clipboard.
         */
        Clipboard.copy = function (text) {
            Clipboard._copyPasteInternal(text);
        };
        /**
         * Gets a string from the clipboard.
         *
         * This method only works if invoked immediately after the user
         * pressed a clipboard paste command (such as ctrl+v).
         *
         * @param callback Function called when the clipboard content
         * has been retrieved. The function receives the clipboard
         * content as a parameter.
         */
        Clipboard.paste = function (callback) {
            Clipboard._copyPasteInternal(callback);
        };
        // ** implementation
        Clipboard._copyPasteInternal = function (textOrCallback) {
            // get active element to restore later
            var activeElement = wijmo.getActiveElement();
            // find parent for temporary input element
            // (body may not be focusable when modal dialogs are used..., TFS 202992)
            // start with closest wijmo control
            var parent = wijmo.closest(activeElement, '.wj-control');
            // move on to non-control parent (since most controls handle the keyboard themselves)
            while (parent && wijmo.Control.getControl(parent)) {
                parent = parent.parentElement;
            }
            // fallback on body
            if (parent == null) {
                parent = document.body;
            }
            // create hidden input element, append it to parent
            if (parent) {
                var el = document.createElement('textarea');
                el.style.position = 'fixed';
                el.style.opacity = '0';
                parent.appendChild(el);
                // initialize text and give element the focus
                if (typeof (textOrCallback) == 'string') {
                    el.value = textOrCallback;
                }
                el.select();
                // prevent multiple ctrl+v's from getting to the hidden element (TFS 151939)
                el.onkeydown = function (e) {
                    if (el.value) {
                        e.preventDefault();
                    }
                };
                // when the clipboard operation is done, remove element, restore focus
                // and invoke the paste callback
                setTimeout(function () {
                    var text = el.value;
                    parent.removeChild(el);
                    activeElement.focus();
                    if (typeof (textOrCallback) == 'function') {
                        textOrCallback(text);
                    }
                }, 100); // Apple needs extra timeOut
            }
        };
        return Clipboard;
    }());
    wijmo.Clipboard = Clipboard;
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Clipboard.js.map
var wijmo;
(function (wijmo) {
    'use strict';
    /**
     * Shows an element as a popup.
     *
     * The popup element becomes a child of the body element,
     * and is positioned above or below a reference rectangle,
     * depending on how much room is available.
     *
     * The reference rectangle may be specified as one of the following:
     *
     * <dl class="dl-horizontal">
     *   <dt>HTMLElement</dt>
     *   <dd>The bounding rectangle of the element.</dd>
     *   <dt>MouseEvent</dt>
     *   <dd>The bounding rectangle of the event's target element.</dd>
     *   <dt>Rect</dt>
     *   <dd>The given rectangle.</dd>
     *   <dt>null</dt>
     *   <dd>No reference rectangle; the popup is centered on the window.</dd>
     * </dl>
     *
     * Call the @see:hidePopup method to hide the popup.
     *
     * @param popup Element to show as a popup.
     * @param ref Reference element or rectangle used to position the popup.
     * @param above Position popup above the reference rectangle if possible.
     * @param fadeIn Use a fade-in animation to make the popup appear gradually.
     * @param copyStyles Copy font and color styles from reference element.
     */
    function showPopup(popup, ref, above, fadeIn, copyStyles) {
        if (ref === void 0) { ref = null; }
        if (above === void 0) { above = false; }
        if (fadeIn === void 0) { fadeIn = false; }
        if (copyStyles === void 0) { copyStyles = true; }
        // the popup parent will be the document body (to avoid clipping issues)
        var parent = document.body;
        // if the reference parameter is an element
        if (ref instanceof HTMLElement) {
            // 1 - make sure the reference is in the DOM
            if (!wijmo.contains(document.body, ref)) {
                return;
            }
            // 2 - adjust the parent to account for ancestors with fixed position or scrollable content
            for (var e = ref.parentElement; e; e = e.parentElement) {
                if (getComputedStyle(e).position == 'fixed') {
                    parent = e;
                    break;
                }
            }
        }
        else {
            // 3 - no reference element, adjust the parent based on the current active element
            // (e.g.grid in dialog, jQuery panel)
            for (var e = wijmo.getActiveElement(); e; e = e.parentElement) {
                if (!wijmo.hasClass(e, 'wj-popup') &&
                    getComputedStyle(e).position == 'fixed') {
                    parent = e;
                    break;
                }
            }
        }
        // make sure popup is the last child of the parent element (TFS 149728)
        // but don't call this unless we have to: it affects the focus
        if (parent.lastChild != popup) {
            parent.appendChild(popup);
        }
        // copy style elements from ref element to popup
        // (since the popup is no longer a child of the ref element)
        if (ref instanceof HTMLElement && copyStyles) {
            var sref = getComputedStyle(ref), bkg = new wijmo.Color(sref.backgroundColor);
            if (bkg.a) {
                wijmo.setCss(popup, {
                    color: sref.color,
                    backgroundColor: sref.backgroundColor,
                    fontFamily: sref.fontFamily,
                    fontSize: sref.fontSize,
                    fontWeight: sref.fontWeight,
                    fontStyle: sref.fontStyle
                });
            }
        }
        // get popup's size, including margins
        wijmo.setCss(popup, {
            position: 'absolute',
            visibility: 'hidden',
            display: ''
        });
        // update layout for any Wijmo controls in the popup
        wijmo.Control.refreshAll(popup);
        // ready to compute margins, size
        var sp = getComputedStyle(popup), my = parseFloat(sp.marginTop) + parseFloat(sp.marginBottom), mx = parseFloat(sp.marginLeft) + parseFloat(sp.marginRight), sz = new wijmo.Size(popup.offsetWidth + mx, popup.offsetHeight + my);
        // ref can be a mouse event, a point, an element, or a rect
        var pos = new wijmo.Point(), rc = null;
        if (ref && ref.clientX != null && ref.clientY != null && ref.pageX != null && ref.pageY != null) {
            if (ref.clientX <= 0 && ref.clientY <= 0 && ref.target) {
                // this looks like a fake mouse event (e.g. context menu key),
                // so use the event target as a reference TFS 117115
                rc = ref.target.getBoundingClientRect();
            }
            else {
                // use pageX/Y-offsetX/Y instead of clientX/Y, which gives wrong results in Chrome/Android
                // REVIEW: this seems to work OK, although in some other places we do the opposite... sigh...
                pos.x = Math.max(0, ref.pageX - pageXOffset);
                pos.y = Math.max(0, ref.pageY - pageYOffset);
            }
        }
        else if (ref instanceof wijmo.Point) {
            pos = ref;
        }
        else if (ref instanceof HTMLElement) {
            rc = ref.getBoundingClientRect();
        }
        else if (ref && ref.top != null && ref.left != null) {
            rc = ref;
        }
        else if (ref == null) {
            pos.x = Math.max(0, (innerWidth - sz.width) / 2);
            pos.y = Math.max(0, Math.round((innerHeight - sz.height) / 2 * .7));
        }
        else {
            throw 'Invalid ref parameter.';
        }
        // calculate min width for the popup
        var minWidth = parseFloat(sp.minWidth);
        // if we have a rect, position popup above or below the rect
        if (rc) {
            var spcAbove = rc.top, spcBelow = innerHeight - rc.bottom, rtl = getComputedStyle(popup).direction == 'rtl';
            if (rtl) {
                pos.x = Math.max(0, rc.right - sz.width);
            }
            else {
                pos.x = Math.max(0, Math.min(rc.left, innerWidth - sz.width));
            }
            if (above) {
                pos.y = spcAbove > sz.height || spcAbove > spcBelow
                    ? Math.max(0, rc.top - sz.height)
                    : rc.bottom;
            }
            else {
                pos.y = spcBelow > sz.height || spcBelow > spcAbove
                    ? rc.bottom
                    : Math.max(0, rc.top - sz.height);
            }
            // make popup at least as wide as the element
            minWidth = Math.max(minWidth, rc.width);
        }
        // handle scroll offset (TFS 202906)
        var ptOffset = new wijmo.Point(0, 0);
        if (ref != null) {
            if (parent == document.body) {
                ptOffset = new wijmo.Point(-pageXOffset, -pageYOffset);
            }
            else if (parent) {
                var bcr = parent.getBoundingClientRect();
                ptOffset = new wijmo.Point(bcr.left - parent.scrollLeft, bcr.top - parent.scrollTop);
            }
        }
        // calculate popup position
        var css = {
            position: ref == null ? 'fixed' : 'absolute',
            left: pos.x - ptOffset.x,
            top: pos.y - ptOffset.y,
            minWidth: minWidth,
            display: '',
            visibility: '',
            zIndex: 1500 // to work in Bootstrap dialogs (zIndex 1050)
        };
        // apply fade in effect
        if (fadeIn) {
            popup.style.opacity = '0';
            wijmo.animate(function (pct) {
                popup.style.opacity = (pct == 1) ? '' : pct.toString();
            });
        }
        // show it
        wijmo.setCss(popup, css);
        // hide the popup if the user scrolls an ancestor other than the body
        var anchor = ref instanceof MouseEvent ? ref.target : ref;
        if (anchor instanceof HTMLElement) {
            var ctl; // listener to scroll events
            for (var se = anchor.parentElement; se && se != document.body; se = se.parentElement) {
                if (getComputedStyle(se).overflowY == 'auto' && se.scrollHeight > se.offsetHeight) {
                    if (!ctl) {
                        ctl = new wijmo.Control(document.createElement('div'));
                    }
                    ctl.addEventListener(se, 'scroll', function (e) {
                        _hidePopup(popup, true);
                        ctl.dispose();
                    });
                }
            }
        }
    }
    wijmo.showPopup = showPopup;
    /**
     * Hides a popup element previously displayed with the @see:showPopup
     * method.
     *
     * @param popup Popup element to hide.
     * @param remove Whether to remove the popup from the DOM or just
     * to hide it.
     * @param fadeOut Whether to use a fade-out animation to make the
     * popup disappear gradually.
     */
    function hidePopup(popup, remove, fadeOut) {
        if (remove === void 0) { remove = true; }
        if (fadeOut === void 0) { fadeOut = false; }
        if (fadeOut) {
            wijmo.animate(function (pct) {
                popup.style.opacity = (1 - pct).toString();
                if (pct == 1) {
                    _hidePopup(popup, remove);
                    popup.style.opacity = '';
                }
            });
        }
        else {
            _hidePopup(popup, remove);
        }
    }
    wijmo.hidePopup = hidePopup;
    function _hidePopup(popup, remove) {
        popup.style.display = 'none';
        if (remove && popup.parentElement) {
            popup.parentElement.removeChild(popup);
        }
    }
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Popup.js.map
var wijmo;
(function (wijmo) {
    'use strict';
    // Note: The Edge browser had an issue that prevented this component
    // from working correctly (see link below); that bug has been fixed
    // and the PrintDocument component now works fine with Edge.
    //  https://connect.microsoft.com/IE/Feedback/Details/1589775
    /**
     * Class that enables the creation of custom documents for printing.
     *
     * The @see:PrintDocument class makes it easy to create documents for printing or
     * exporting to PDF. Most browsers allow you to select the paper size, orientation,
     * margins, and whether to include page headers and footers.
     *
     * To use, instantiate a @see:PrintDocument, add content using the @see:append
     * method, and finish by calling the @see:print method.
     *
     * For example:
     * <pre>// create the document
     * var doc = new wijmo.PrintDocument({
     *   title: 'PrintDocument Test'
     * });
     * // add some simple text
     * doc.append('&lt;h1&gt;Printing Example&lt;/h1&gt;');
     * doc.append('&lt;p&gt;This document was created using the &lt;b&gt;PrintDocument&lt;/b&gt; class.&lt;/p&gt;');
     * // add some existing elements
     * doc.append(document.getElementById('gaugeControl'));
     * // print the document (or export it to PDF)
     * doc.print();</pre>
     */
    var PrintDocument = (function () {
        // ** ctor
        /**
         * Initializes a new instance of the @see:PrintDocument class.
         *
         * @param options JavaScript object containing initialization data for the @see:PrintDocument.
         */
        function PrintDocument(options) {
            this._copyCss = true;
            if (options != null) {
                wijmo.copy(this, options);
            }
        }
        Object.defineProperty(PrintDocument.prototype, "title", {
            // ** object model
            /**
             * Gets or sets the document title.
             */
            get: function () {
                return this._title;
            },
            set: function (value) {
                this._title = wijmo.asString(value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(PrintDocument.prototype, "copyCss", {
            /**
             * Gets or sets a value that determines whether the @see:PrintDocument should include the CSS
             * style sheets defined in the main document.
             */
            get: function () {
                return this._copyCss;
            },
            set: function (value) {
                this._copyCss = wijmo.asBoolean(value);
            },
            enumerable: true,
            configurable: true
        });
        /**
         * Adds a CSS style sheet to the document.
         *
         * @param href URL of the CSS file that should be added to the document.
         */
        PrintDocument.prototype.addCSS = function (href) {
            if (!this._css) {
                this._css = [];
            }
            this._css.push(href);
        };
        /**
         * Appends an HTML element or string to the document.
         *
         * @param child HTML element or string to append to the document.
         */
        PrintDocument.prototype.append = function (child) {
            var doc = this._getDocument();
            if (wijmo.isString(child)) {
                doc.write(child);
            }
            else if (child instanceof HTMLElement) {
                doc.write(child.outerHTML);
            }
            else {
                wijmo.assert(false, 'child should be an HTML element or a string.');
            }
        };
        /**
         * Prints the document.
         */
        PrintDocument.prototype.print = function () {
            var _this = this;
            if (this._iframe) {
                // close the document
                this._close();
                // give it some time before printing/disposing
                setTimeout(function () {
                    // print the document
                    var wnd = _this._iframe.contentWindow;
                    wnd.focus();
                    wnd.print();
                    // done, dispose of the iframe
                    document.body.removeChild(_this._iframe);
                    _this._iframe = null;
                }, 100);
            }
        };
        // ** implementation
        // gets a reference to the print document
        PrintDocument.prototype._getDocument = function () {
            if (!this._iframe) {
                // create iframe
                this._iframe = document.createElement('iframe');
                // initialize iframe
                var s = this._iframe.style;
                s.position = 'fixed';
                s.left = '10000px';
                s.top = '10000px';
                document.body.appendChild(this._iframe);
            }
            return this._iframe.contentDocument;
        };
        // closes the print document before printing
        PrintDocument.prototype._close = function () {
            // close document before applying title and style sheets
            var doc = this._getDocument();
            doc.close();
            // add title
            if (this.title) {
                var et = doc.querySelector('title');
                if (!et) {
                    et = doc.createElement('title');
                    doc.head.appendChild(et);
                }
                et.textContent = this.title;
            }
            // add main document style sheets
            if (this._copyCss) {
                var links = document.head.querySelectorAll('LINK');
                for (var i = 0; i < links.length; i++) {
                    var link = links[i];
                    if (link.href.match(/\.css$/i) && link.rel.match(/stylesheet/i)) {
                        var xhr = wijmo.httpRequest(link.href, { async: false });
                        this._addStyle(xhr.responseText);
                    }
                }
                var styles = document.head.querySelectorAll('STYLE');
                for (var i = 0; i < styles.length; i++) {
                    this._addStyle(styles[i].textContent);
                }
            }
            // add extra style sheets
            if (this._css) {
                for (var i = 0; i < this._css.length; i++) {
                    var es = doc.createElement('style'), xhr = wijmo.httpRequest(this._css[i], { async: false });
                    es.textContent = xhr.responseText;
                    doc.head.appendChild(es);
                }
            }
        };
        PrintDocument.prototype._addStyle = function (style) {
            var doc = this._getDocument(), es = doc.createElement('style');
            es.textContent = style;
            doc.head.appendChild(es);
        };
        return PrintDocument;
    }());
    wijmo.PrintDocument = PrintDocument;
})(wijmo || (wijmo = {}));
//# sourceMappingURL=PrintDocument.js.map
var wijmo;
(function (wijmo) {
    'use strict';
    /**
     * Class that provides masking services to an HTMLInputElement.
     */
    var _MaskProvider = (function () {
        /**
         * Initializes a new instance of the @see:_MaskProvider class.
         *
         * @param input Input element to be masked.
         * @param mask Input mask.
         * @param promptChar Character used to indicate input positions.
         */
        function _MaskProvider(input, mask, promptChar) {
            if (mask === void 0) { mask = null; }
            if (promptChar === void 0) { promptChar = '_'; }
            this._promptChar = '_';
            this._mskArr = [];
            this._full = true;
            this._hbInput = this._input.bind(this);
            this._hbKeyDown = this._keydown.bind(this);
            this._hbKeyPress = this._keypress.bind(this);
            this._hbCompositionStart = this._compositionstart.bind(this);
            this._hbCompositionEnd = this._compositionend.bind(this);
            this.mask = mask;
            this.input = input;
            this.promptChar = promptChar;
            this._connect(true);
        }
        Object.defineProperty(_MaskProvider.prototype, "input", {
            /**
             * Gets or sets the Input element to be masked.
             */
            get: function () {
                return this._tbx;
            },
            set: function (value) {
                this._connect(false);
                this._tbx = value;
                this._connect(true);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(_MaskProvider.prototype, "mask", {
            /**
             * Gets or sets the input mask used to validate input.
             */
            get: function () {
                return this._msk;
            },
            set: function (value) {
                if (value != this._msk) {
                    this._msk = wijmo.asString(value, true);
                    this._parseMask();
                    this._valueChanged();
                }
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(_MaskProvider.prototype, "promptChar", {
            /**
             * Gets or sets the input mask used to validate input.
             */
            get: function () {
                return this._promptChar;
            },
            set: function (value) {
                if (value != this._promptChar) {
                    this._promptChar = wijmo.asString(value, false);
                    wijmo.assert(this._promptChar.length == 1, 'promptChar must be a string with length 1.');
                    this._valueChanged();
                }
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(_MaskProvider.prototype, "maskFull", {
            /**
             * Gets a value that indicates whether the mask has been completely filled.
             */
            get: function () {
                return this._full;
            },
            enumerable: true,
            configurable: true
        });
        /**
         * Gets an array with the position of the first and last wildcard characters in the mask.
         */
        _MaskProvider.prototype.getMaskRange = function () {
            return this._mskArr.length ? [this._firstPos, this._lastPos] : [0, this._tbx.value.length - 1];
        };
        /**
         * Gets the raw value of the editor, excluding prompts and literals.
         */
        _MaskProvider.prototype.getRawValue = function () {
            var text = this._tbx.value, ret = '';
            // no mask? it's all raw
            if (!this.mask) {
                return text;
            }
            // get raw input (no literals or prompts)
            for (var i = 0; i < this._mskArr.length && i < text.length; i++) {
                if (!this._mskArr[i].literal && text[i] != this._promptChar) {
                    ret += text[i];
                }
            }
            // done
            return ret;
        };
        /**
         * Updates the control mask and content.
         */
        _MaskProvider.prototype.refresh = function () {
            this._parseMask();
            this._valueChanged();
        };
        // ** event handlers
        // apply mask and update cursor position after any changes
        _MaskProvider.prototype._input = function () {
            var _this = this;
            if (!this._composing) {
                setTimeout(function () {
                    _this._valueChanged();
                });
            }
        };
        // special handling for backspacing over literals
        _MaskProvider.prototype._keydown = function (e) {
            // ignore backspaces before first wildcard (TFS 199372, 199901)
            if (e.keyCode == wijmo.Key.Back) {
                var start = this._tbx.selectionStart, end = this._tbx.selectionEnd;
                if (start <= this._firstPos && end == start) {
                    e.preventDefault();
                    this._backSpace = false;
                    return;
                }
            }
            // handle backspacing over literals
            this._backSpace = (e.keyCode == wijmo.Key.Back);
        };
        // prevent flicker when invalid keys are pressed
        // NOTE: IE on Windows Phone 8.x does not raise keypress!!!
        _MaskProvider.prototype._keypress = function (e) {
            if (!e.ctrlKey && !e.metaKey && !e.altKey && !this._composing && this._preventKey(e.charCode)) {
                e.preventDefault();
            }
        };
        // handle IME composition
        _MaskProvider.prototype._compositionstart = function (e) {
            this._composing = true;
        };
        _MaskProvider.prototype._compositionend = function (e) {
            var _this = this;
            this._composing = false;
            setTimeout(function () {
                _this._valueChanged();
            });
        };
        // ** implementation
        // prevent a key from being handled if it is invalid
        _MaskProvider.prototype._preventKey = function (charCode) {
            if (charCode && this._mskArr.length) {
                var tbx = this._tbx, start = tbx.selectionStart, key = String.fromCharCode(charCode);
                // make sure we're at a placeholder
                if (start < this._firstPos) {
                    start = this._firstPos;
                    wijmo.setSelectionRange(tbx, start);
                }
                // past the end?
                if (start >= this._mskArr.length) {
                    return true;
                }
                // skip over literals and validate templates (TFS 117584)
                var m = this._mskArr[start];
                if (m.literal) {
                    this._validatePosition(start);
                }
                else if (m.wildCard != key && !this._isCharValid(m.wildCard, key)) {
                    return true;
                }
            }
            // key seems OK, do not prevent
            return false;
        };
        // connect or disconnect event handlers for the input element
        _MaskProvider.prototype._connect = function (connect) {
            var tbx = this._tbx;
            if (tbx) {
                if (connect) {
                    this._autoComplete = tbx.autocomplete;
                    this._spellCheck = tbx.spellcheck;
                    tbx.autocomplete = 'off'; // important for mobile browsers (including Chrome/Android)
                    tbx.spellcheck = false; // no spell-checking on masked inputs
                    tbx.addEventListener('input', this._hbInput);
                    tbx.addEventListener('keydown', this._hbKeyDown, true);
                    tbx.addEventListener('keypress', this._hbKeyPress, true);
                    tbx.addEventListener('compositionstart', this._hbCompositionStart, true);
                    tbx.addEventListener('compositionend', this._hbCompositionEnd, true);
                    this._valueChanged();
                }
                else {
                    tbx.autocomplete = this._autoComplete;
                    tbx.spellcheck = this._spellCheck;
                    tbx.removeEventListener('input', this._hbInput);
                    tbx.removeEventListener('keydown', this._hbKeyDown, true);
                    tbx.removeEventListener('keypress', this._hbKeyPress, true);
                    tbx.removeEventListener('compositionstart', this._hbCompositionStart, true);
                    tbx.removeEventListener('compositionend', this._hbCompositionEnd, true);
                }
            }
        };
        // apply the mask keeping the cursor position and the focus
        _MaskProvider.prototype._valueChanged = function () {
            if (this._tbx && this._msk) {
                // keep track of selection start, character at selection
                var tbx = this._tbx, start = tbx.selectionStart, oldChar = (start > 0) ? tbx.value[start - 1] : '';
                // apply the mask
                tbx.value = this._applyMask();
                // backtrack if the original character was replaced with a prompt
                var newChar = (start > 0) ? tbx.value[start - 1] : '';
                if (start > 0 && newChar == this._promptChar && oldChar != this.promptChar) {
                    start--;
                }
                // and validate the position as usual
                this._validatePosition(start);
            }
        };
        // applies the mask to the current text content, returns the result 
        // this is usually a valid string with the same length as the mask, unless 
        // (a) there's no mask, or 
        // (b) there's no text and the input is not required
        _MaskProvider.prototype._applyMask = function () {
            // assume we're complete
            this._full = true;
            // no mask? accept everything
            var text = this._tbx.value;
            if (!this._msk) {
                return text;
            }
            // no text OK if not required
            if (!text && !this._tbx.required) {
                return text;
            }
            // handle vague literals (could be interpreted as content)
            var ret = '', pos = 0;
            text = this._handleVagueLiterals(text);
            // build output string based on current content and mask
            for (var i = 0; i < this._mskArr.length; i++) {
                // get mask element
                var m = this._mskArr[i], c = m.literal;
                // if this is a literal, match with text at cursor
                if (c && c == text[pos]) {
                    pos++;
                }
                // if it is a wildcard, match with text starting at the cursor
                if (m.wildCard) {
                    c = this._promptChar;
                    if (text) {
                        for (var j = pos; j < text.length; j++) {
                            if (this._isCharValid(m.wildCard, text[j])) {
                                c = text[j];
                                switch (m.charCase) {
                                    case '>':
                                        c = c.toUpperCase();
                                        break;
                                    case '<':
                                        c = c.toLowerCase();
                                        break;
                                }
                                break;
                            }
                        }
                        pos = j + 1;
                    }
                    // still prompt? then the mask is not full
                    if (c == this._promptChar) {
                        this._full = false;
                    }
                }
                // add to output
                ret += c;
            }
            // done applying mask, return result
            return ret;
        };
        // fix text to handle vague literals (that could be interpreted as content)
        _MaskProvider.prototype._handleVagueLiterals = function (text) {
            // looks like a paste, don't mess with it (TFS 139412, 145560)
            if (text.length > this._mskArr.length + 1) {
                return text;
            }
            // see if we're shrinking or growing
            var delta = text.length - this._mskArr.length;
            if (delta != 0 && text.length > 1) {
                // see if we have a 'vague' literal
                var badIndex = -1, start = Math.max(0, this._tbx.selectionStart - delta);
                for (var i = start; i < this._mskArr.length; i++) {
                    if (this._mskArr[i].vague) {
                        badIndex = i;
                        break;
                    }
                }
                // we do, so handle it
                if (badIndex > -1) {
                    //console.log(' text: [' + text + ']');
                    if (delta < 0) {
                        var pad = Array(1 - delta).join(this._promptChar), index = badIndex + delta;
                        if (index > -1) {
                            text = text.substr(0, index) + pad + text.substr(index);
                        }
                    }
                    else {
                        while (badIndex > 0 && this._mskArr[badIndex - 1].literal) {
                            badIndex--;
                        }
                        text = text.substr(0, badIndex) + text.substr(badIndex + delta);
                    }
                }
            }
            // done
            return text;
        };
        // checks whether a character is valid for a given mask character
        _MaskProvider.prototype._isCharValid = function (mask, c) {
            var ph = this._promptChar;
            switch (mask) {
                // regular wildcards
                case '0':
                    return (c >= '0' && c <= '9') || c == ph;
                case '9':
                    return (c >= '0' && c <= '9') || c == ' ' || c == ph;
                case '#':
                    return (c >= '0' && c <= '9') || c == ' ' || c == '+' || c == '-' || c == ph;
                case 'L':
                    return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || c == ph;
                case 'l':
                    return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || c == ' ' || c == ph;
                case 'A':
                    return (c >= '0' && c <= '9') || (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || c == ph;
                case 'a':
                    return (c >= '0' && c <= '9') || (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || c == ' ' || c == ph;
                // Katakana/Hiragana wildcards
                case '\uff19':
                    return (c >= '\uFF10' && c <= '\uff19') || c == ph;
                case '\uff2a': // DBCS Hiragana
                case '\uff27':
                    if (mask == '\uff27' && _MaskProvider._X_DBCS_BIG_HIRA.indexOf(c) > -1)
                        return false; // exceptions
                    return (c >= '\u3041' && c <= '\u3096') || c == ph;
                case '\uff2b': // DBCS Katakana
                case '\uff2e':
                    if (mask == '\uff2e' && _MaskProvider._X_DBCS_BIG_KATA.indexOf(c) > -1)
                        return false; // exceptions
                    return (c >= '\u30a1' && c <= '\u30fa') || c == ph;
                case '\uff3a':
                    return (c <= '\u0021' || c >= '\u00ff') || c == ph;
                case 'H':
                    return (c >= '\u0021' && c <= '\u00ff') || c == ph;
                case 'K': // SBCS Katakana
                case 'N':
                    if (mask == 'N' && _MaskProvider._X_SBCS_BIG_KATA.indexOf(c) > -1)
                        return false; // exceptions
                    return (c >= '\uff66' && c <= '\uff9f') || c == ph;
            }
            return false;
        };
        // skip over literals
        _MaskProvider.prototype._validatePosition = function (start) {
            var msk = this._mskArr;
            // skip left if the last key pressed was a backspace
            if (this._backSpace) {
                while (start > 0 && start < msk.length && msk[start - 1].literal) {
                    start--;
                }
            }
            // skip right over literals
            if (start == 0 || !this._backSpace) {
                while (start < msk.length && msk[start].literal) {
                    start++;
                }
            }
            // move selection to start
            if (wijmo.getActiveElement() == this._tbx) {
                wijmo.setSelectionRange(this._tbx, start);
            }
            // no longer backspacing
            this._backSpace = false;
        };
        // parse mask into internal mask, literals, and case
        _MaskProvider.prototype._parseMask = function () {
            // clear internal mask info
            this._mskArr = [];
            this._firstPos = -1;
            this._lastPos = -1;
            // parse new mask
            var msk = this._msk, currCase = '|', c;
            for (var i = 0; msk && i < msk.length; i++) {
                switch (msk[i]) {
                    // wildcards
                    case '0': // digit.
                    case '9': // Digit or space.
                    case '#': // Digit, sign, or space.
                    case 'A': // Alphanumeric.
                    case 'a': // Alphanumeric or space.
                    case 'L': // Letter.
                    case 'l': // Letter or space.
                    // Katakana/Hiragana wildcards
                    case '\uff19': // DBCS Digit.
                    case '\uff2a': // DBCS Hiragana.
                    case '\uff27': // DBCS big Hiragana.
                    case '\uff2b': // DBCS Katakana.
                    case '\uff2e': // DBCS big Katakana.
                    case '\uff3a': // Any DBCS
                    case 'K': // SBCS Katakana.
                    case 'N': // SBCS big Katakana.
                    case 'H':
                        if (this._firstPos < 0) {
                            this._firstPos = this._mskArr.length;
                        }
                        this._lastPos = this._mskArr.length;
                        this._mskArr.push(new _MaskElement(msk[i], currCase));
                        break;
                    // localized literals
                    case '.': // Decimal separator.
                    case ',': // Thousands separator.
                    case ':': // Time separator.
                    case '/': // Date separator.
                    case '$':
                        switch (msk[i]) {
                            case '.':
                            case ',':
                                c = wijmo.culture.Globalize.numberFormat[msk[i]];
                                break;
                            case ':':
                            case '/':
                                c = wijmo.culture.Globalize.calendar[msk[i]];
                                break;
                            case '$':
                                c = wijmo.culture.Globalize.numberFormat.currency.symbol;
                                break;
                        }
                        for (var j = 0; j < c.length; j++) {
                            this._mskArr.push(new _MaskElement(c[j]));
                        }
                        break;
                    // case-shifting
                    case '<': // Shift down (converts characters that follow to lowercase).
                    case '>': // Shift up (converts characters that follow to uppercase).
                    case '|':
                        currCase = msk[i];
                        break;
                    // literals
                    case '\\':
                        if (i < msk.length - 1)
                            i++;
                        this._mskArr.push(new _MaskElement(msk[i]));
                        break;
                    default:
                        this._mskArr.push(new _MaskElement(msk[i]));
                        break;
                }
            }
            // keep track of vague (ambiguous) literals
            for (var i = 0; i < this._mskArr.length; i++) {
                var elem = this._mskArr[i];
                if (elem.literal) {
                    for (var j = 0; j < i; j++) {
                        var m = this._mskArr[j];
                        if (m.wildCard && this._isCharValid(m.wildCard, elem.literal)) {
                            elem.vague = true;
                            break;
                        }
                    }
                }
            }
        };
        // big DBCS/SBCS exception lists
        _MaskProvider._X_DBCS_BIG_HIRA = '\u3041\u3043\u3045\u3047\u3049\u3063\u3083\u3085\u3087\u308e\u3095\u3096';
        _MaskProvider._X_DBCS_BIG_KATA = '\u30a1\u30a3\u30a5\u30a7\u30a9\u30c3\u30e3\u30e5\u30e7\u30ee\u30f5\u30f6';
        _MaskProvider._X_SBCS_BIG_KATA = '\uff67\uff68\uff69\uff6a\uff6b\uff6c\uff6d\uff6e\uff6f';
        return _MaskProvider;
    }());
    wijmo._MaskProvider = _MaskProvider;
    /**
     * Class that contains information about a position in an input mask.
     */
    var _MaskElement = (function () {
        /**
         * Initializes a new instance of the @see:_MaskElement class.
         *
         * @param wildcardOrLiteral Wildcard or literal character
         * @param charCase Whether to convert wildcard matches to upper or lowercase.
         */
        function _MaskElement(wildcardOrLiteral, charCase) {
            if (charCase) {
                this.wildCard = wildcardOrLiteral;
                this.charCase = charCase;
            }
            else {
                this.literal = wildcardOrLiteral;
            }
        }
        return _MaskElement;
    }());
    wijmo._MaskElement = _MaskElement;
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_MaskProvider.js.map
//
// IE9 polyfills
var wijmo;
(function (wijmo) {
    'use strict';
    // expose IE flag to callers that need it
    var _isIE = navigator.userAgent.match(/MSIE |Trident\/|Edge\//) != null;
    function isIE() {
        return _isIE;
    }
    wijmo.isIE = isIE;
    // expose IE9 flag to callers that need it
    var _isIE9;
    function isIE9() {
        return _isIE9;
    }
    wijmo.isIE9 = isIE9;
    // implement HTML5 drag-drop behavior in IE9.
    if (document.doctype && navigator.appVersion.indexOf('MSIE 9') > -1) {
        // remember this is IE9...
        _isIE9 = true;
        // TFS 140812: 'selectstart' does not work in popup dialogs, so use 'mousemove'
        // instead. It's less efficient but it works, and this only matters in IE9.
        document.addEventListener('mousemove', function (e) {
            if (e.which == 1) {
                for (var el = e.target; el; el = el.parentNode) {
                    if (el.attributes && el.attributes['draggable']) {
                        el.dragDrop();
                        return false;
                    }
                }
            }
        });
    }
    // implement requestAnimationFrame/cancelAnimationFrame in IE9.
    // https://gist.github.com/rma4ok/3371337
    if (!window['requestAnimationFrame']) {
        var expectedTime = 0;
        window['requestAnimationFrame'] = function (callback) {
            var currentTime = Date.now(), adjustedDelay = 16 - (currentTime - expectedTime), delay = adjustedDelay > 0 ? adjustedDelay : 0;
            expectedTime = currentTime + delay;
            return setTimeout(function () {
                callback(expectedTime);
            }, delay);
        };
        window['cancelAnimationFrame'] = clearTimeout;
    }
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_IE9Polyfills.js.map
