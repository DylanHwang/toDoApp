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
module wijmo {
    'use strict';

    // major (ECMAScript version required).
    // year/trimester.
    // sequential
    var _VERSION = '5.20162.211';

    /**
     * Gets the version of the Wijmo library that is currently loaded.
     */
    export function getVersion(): string {
        return _VERSION;
    }

    /**
     * Specifies constants that represent keyboard codes.
     *
     * This enumeration is useful when handling <b>keyDown</b> events.
     */
    export enum Key {
        /** The backspace key. */
        Back = 8,
        /** The tab key. */
        Tab = 9,
        /** The enter key. */
        Enter = 13,
        /** The escape key. */
        Escape = 27,
        /** The space key. */
        Space = 32,
        /** The page up key. */
        PageUp = 33,
        /** The page down key. */
        PageDown = 34,
        /** The end key. */
        End = 35,
        /** The home key. */
        Home = 36,
        /** The left arrow key. */
        Left = 37,
        /** The up arrow key. */
        Up = 38,
        /** The right arrow key. */
        Right = 39,
        /** The down arrow key. */
        Down = 40,
        /** The delete key. */
        Delete = 46,
        /** The F1 key. */
        F1 = 112,
        /** The F2 key. */
        F2 = 113,
        /** The F3 key. */
        F3 = 114,
        /** The F4 key. */
        F4 = 115,
        /** The F5 key. */
        F5 = 116,
        /** The F6 key. */
        F6 = 117,
        /** The F7 key. */
        F7 = 118,
        /** The F8 key. */
        F8 = 119,
        /** The F9 key. */
        F9 = 120,
        /** The F10 key. */
        F10 = 121,
        /** The F11 key. */
        F11 = 122,
        /** The F12 key. */
        F12 = 123
    }

    /**
     * Specifies constants that represent data types.
     *
     * Use the @see:getType method to get a @see:DataType from a value.
     */
    export enum DataType {
        /** Object (anything). */
        Object,
        /** String. */
        String,
        /** Number. */
        Number,
        /** Boolean. */
        Boolean,
        /** Date (date and time). */
        Date,
        /** Array. */
        Array
    }

    // general-purpose utilities.
    // note: avoid letting this grow too much!!!

    /**
     * Allows callers to verify whether an object implements an interface.
     */
    export interface IQueryInterface {
        /**
         * Returns true if the object implements a given interface.
         *
         * @param interfaceName Name of the interface to look for.
         */
        implementsInterface(interfaceName: string): boolean;
    }
    /**
     * Casts a value to a type if possible.
     *
     * @param value Value to cast.
     * @param type Type or interface name to cast to.
     * @return The value passed in if the cast was successful, null otherwise.
     */
    export function tryCast(value: any, type: any): any {

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
    /**
     * Determines whether an object is a primitive type (string, number, boolean, or date).
     *
     * @param value Value to test.
     */
    export function isPrimitive(value: any): boolean {
        return isString(value) || isNumber(value) || isBoolean(value) || isDate(value);
    }
    /**
     * Determines whether an object is a string.
     *
     * @param value Value to test.
     */
    export function isString(value: any): boolean {
        return typeof (value) == 'string';
    }
    /**
     * Determines whether a string is null, empty, or whitespace only.
     *
     * @param value Value to test.
     */
    export function isNullOrWhiteSpace(value: string): boolean {
        return value == null ? true : value.replace(/\s/g, '').length < 1;
    }
    /**
     * Determines whether an object is a number.
     *
     * @param value Value to test.
     */
    export function isNumber(value: any): boolean {
        return typeof (value) == 'number';
    }
    /**
     * Determines whether an object is an integer.
     *
     * @param value Value to test.
     */
    export function isInt(value: any): boolean {
        return isNumber(value) && value == Math.round(value);
    }
    /**
     * Determines whether an object is a Boolean.
     *
     * @param value Value to test.
     */
    export function isBoolean(value: any): boolean {
        return typeof (value) == 'boolean';
    }
    /**
     * Determines whether an object is a function.
     *
     * @param value Value to test.
     */
    export function isFunction(value: any): boolean {
        return typeof (value) == 'function';
    }
    /**
     * Determines whether an object is undefined.
     *
     * @param value Value to test.
     */
    export function isUndefined(value: any): boolean {
        return typeof value == 'undefined'
    }
    /**
     * Determines whether an object is a Date.
     *
     * @param value Value to test.
     */
    export function isDate(value: any): boolean {
        return value instanceof Date && !isNaN(value.getTime());
    }
    /**
     * Determines whether an object is an Array.
     *
     * @param value Value to test.
     */
    export function isArray(value: any): boolean {
        return value instanceof Array;
    }
    /**
     * Determines whether a value is an object
     * (as opposed to a value type, an array, or a Date).
     *
     * @param value Value to test.
     */
    export function isObject(value: any): boolean {
        return value != null && typeof value == 'object' && !isDate(value) && !isArray(value);
    }

    /**
     * Converts mouse or touch event arguments into a @see:Point in page coordinates.
     */
    export function mouseToPage(e: any): Point {

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

    /**
     * Gets the type of a value.
     *
     * @param value Value to test.
     * @return A @see:DataType value representing the type of the value passed in.
     */
    export function getType(value: any): DataType {
        if (isNumber(value)) return DataType.Number;
        if (isBoolean(value)) return DataType.Boolean;
        if (isDate(value)) return DataType.Date;
        if (isString(value)) return DataType.String;
        if (isArray(value)) return DataType.Array;
        return DataType.Object;
    }
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
    export function changeType(value: any, type: DataType, format: string): any {
        if (value != null) {

            // convert strings to numbers, dates, or booleans
            if (isString(value)) {
                switch (type) {

                    case DataType.Number:
                        var num = Globalize.parseFloat(value, format);
                        return isNaN(num) ? value : num;

                    case DataType.Date:
                        var date = Globalize.parseDate(value, format);
                        if (!date && !format && value) {
                            date = new Date(value); // fallback on JavaScript parser
                        }
                        return date && isFinite(date.getTime()) ? date : value;

                    case DataType.Boolean:
                        switch ((<string>value).toLowerCase()) {
                            case 'true': return true;
                            case 'false': return false;
                        }
                        return value; // TFS 125067
                }
            }

            // convert anything to string
            if (type == DataType.String) {
                return Globalize.format(value, format);
            }
        }

        // did not convert...
        //console.log('did not convert "' + value + '" to type ' + DataType[type]);
        return value;
    }
    /**
     * Rounds or truncates a number to a specified precision.
     *
     * @param value Value to round or truncate.
     * @param prec Number of decimal digits for the result.
     * @param truncate Whether to truncate or round the original value.
     */
    export function toFixed(value: number, prec: number, truncate: boolean): number {
        if (truncate) {
            var s = value.toString(),
                decPos = s.indexOf('.');
            if (decPos > -1) {
                s = s.substr(0, decPos + 1 + prec);
                value = parseFloat(s);
            }
        } else {
            var s = value.toFixed(prec);
            value = parseFloat(s);
        }
        return value;
    }
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
    export function format(format: string, data: any, formatFunction?: Function): string {
        format = asString(format);
        return format.replace(/\{(.*?)(:(.*?))?\}/g, function (match, name, x, fmt) {
            var val = match;
            if (name && name[0] != '{' && data) {

                // get the value
                val = data[name];

                // apply static format
                if (fmt) {
                    val = Globalize.format(val, fmt);
                }

                // apply format function
                if (formatFunction) {
                    val = formatFunction(data, name, fmt, val);
                }
            }
            return val == null ? '' : val;
        });
    }
    /**
     * Clamps a value between a minimum and a maximum.
     *
     * @param value Original value.
     * @param min Minimum allowed value.
     * @param max Maximum allowed value.
     */
    export function clamp(value: number, min: number, max: number): number {
        if (value != null) {
            if (max != null && value > max) value = max;
            if (min != null && value < min) value = min;
        }
        return value;
    }
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
    export function copy(dst: any, src: any) {
        if (src) {
            for (var key in src) {
                if (key[0] != '_') { // skip non-public properties
                    assert(key in dst, 'Unknown property "' + key + '".');
                    var value = src[key];
                    if (!dst._copy || !dst._copy(key, value)) { // allow overrides
                        if (dst[key] instanceof Event && isFunction(value)) {
                            dst[key].addHandler(value); // add event handler
                        } else if (isObject(value) && dst[key] && key != 'itemsSource') {
                            copy(dst[key], value); // copy sub-objects
                        } else {
                            dst[key] = value; // assign values
                        }
                    }
                }
            }
        }
    }
    /**
     * Throws an exception if a condition is false.
     *
     * @param condition Condition expected to be true.
     * @param msg Message of the exception if the condition is not true.
     */
    export function assert(condition: boolean, msg: string) {
        if (!condition) {
            throw '** Assertion failed in Wijmo: ' + msg;
        }
    }
    /**
     * Outputs a message to indicate a member has been deprecated.
     *
     * @param oldMember Member that has been deprecated.
     * @param newMember Member that replaces the one that has been deprecated.
     */
    export function _deprecated(oldMember: string, newMember: string) {
        console.error('** WARNING: "' + oldMember + '" has been deprecated; please use "' + newMember + '" instead.');
    }
    /**
     * Asserts that a value is a string.
     *
     * @param value Value supposed to be a string.
     * @param nullOK Whether null values are acceptable.
     * @return The string passed in.
     */
    export function asString(value: string, nullOK = true): string {
        assert((nullOK && value == null) || isString(value), 'String expected.');
        return value;
    }
    /**
     * Asserts that a value is a number.
     *
     * @param value Value supposed to be numeric.
     * @param nullOK Whether null values are acceptable.
     * @param positive Whether to accept only positive numeric values.
     * @return The number passed in.
     */
    export function asNumber(value: number, nullOK = false, positive = false): number {
        assert((nullOK && value == null) || isNumber(value), 'Number expected.');
        if (positive && value && value < 0) throw 'Positive number expected.';
        return value;
    }
    /**
     * Asserts that a value is an integer.
     *
     * @param value Value supposed to be an integer.
     * @param nullOK Whether null values are acceptable.
     * @param positive Whether to accept only positive integers.
     * @return The number passed in.
     */
    export function asInt(value: number, nullOK = false, positive = false): number {
        assert((nullOK && value == null) || isInt(value), 'Integer expected.');
        if (positive && value && value < 0) throw 'Positive integer expected.';
        return value;
    }
    /**
     * Asserts that a value is a Boolean.
     *
     * @param value Value supposed to be Boolean.
     * @param nullOK Whether null values are acceptable.
     * @return The Boolean passed in.
     */
    export function asBoolean(value: boolean, nullOK = false): boolean {
        assert((nullOK && value == null) || isBoolean(value), 'Boolean expected.');
        return value;
    }
    /**
     * Asserts that a value is a Date.
     *
     * @param value Value supposed to be a Date.
     * @param nullOK Whether null values are acceptable.
     * @return The Date passed in.
     */
    export function asDate(value: Date, nullOK = false): Date {

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
    /**
     * Asserts that a value is a function.
     *
     * @param value Value supposed to be a function.
     * @param nullOK Whether null values are acceptable.
     * @return The function passed in.
     */
    export function asFunction(value: any, nullOK = true): Function {
        assert((nullOK && value == null) || isFunction(value), 'Function expected.');
        return value;
    }
    /**
     * Asserts that a value is an array.
     *
     * @param value Value supposed to be an array.
     * @param nullOK Whether null values are acceptable.
     * @return The array passed in.
     */
    export function asArray(value: any, nullOK = true): any[] {
        assert((nullOK && value == null) || isArray(value), 'Array expected.');
        return value;
    }
    /**
     * Asserts that a value is an instance of a given type.
     *
     * @param value Value to be checked.
     * @param type Type of value expected.
     * @param nullOK Whether null values are acceptable.
     * @return The value passed in.
     */
    export function asType(value: any, type: any, nullOK = false): any {
        value = tryCast(value, type);
        assert(nullOK || value != null, type + ' expected.');
        return value;
    }
    /**
     * Asserts that a value is a valid setting for an enumeration.
     *
     * @param value Value supposed to be a member of the enumeration.
     * @param enumType Enumeration to test for.
     * @param nullOK Whether null values are acceptable.
     * @return The value passed in.
     */
    export function asEnum(value: number, enumType: any, nullOK = false): number {
        if (value == null && nullOK) return null;
        var e = enumType[value];
        assert(e != null, 'Invalid enum value.');
        return isNumber(e) ? e : value;
    }
    /**
     * Asserts that a value is an @see:ICollectionView or an Array.
     *
     * @param value Array or @see:ICollectionView.
     * @param nullOK Whether null values are acceptable.
     * @return The @see:ICollectionView that was passed in or a @see:CollectionView
     * created from the array that was passed in.
     */
    export function asCollectionView(value: any, nullOK = true): collections.ICollectionView {
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
        return new collections.CollectionView(value);
    }
    /**
     * Checks whether an @see:ICollectionView is defined and not empty.
     *
     * @param value @see:ICollectionView to check.
     */
    export function hasItems(value: collections.ICollectionView) {
        return value && value.items && value.items.length;
    }
    /**
     * Converts a camel-cased string into a header-type string by capitalizing the first letter
     * and adding spaces before uppercase characters preceded by lower-case characters.
     *
     * For example, 'somePropertyName' becomes 'Some Property Name'.
     *
     * @param text String to convert to header case.
     */
    export function toHeaderCase(text: string): string {
        return text && text.length
            ? text[0].toUpperCase() + text.substr(1).replace(/([a-z])([A-Z])/g, '$1 $2')
            : '';
    }
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
    export function escapeHtml(text: string) {
        if (isString(text)) {
            text = text.replace(/[&<>"'\/]/g, function (s) {
                return _ENTITYMAP[s];
            });
        }
        return text;
    }
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
    export function hasClass(e: HTMLElement, className: string): boolean {

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
    /**
     * Removes a class from an element.
     *
     * @param e Element that will have the class removed.
     * @param className Class to remove form the element.
     */
    export function removeClass(e: HTMLElement, className: string) {
        if (e && className && e.setAttribute && hasClass(e, className)) {
            var rx = new RegExp('((\\s|^)' + className + '(\\s|$))', 'g'),
                cn = e.getAttribute('class');
            cn = cn.replace(rx, ' ').replace(/ +/g, ' ').trim();
            e.setAttribute('class', cn);
        }
    }
    /**
     * Adds a class to an element.
     *
     * @param e Element that will have the class added.
     * @param className Class to add to the element.
     */
    export function addClass(e: HTMLElement, className: string) {
        if (e && className && e.setAttribute && !hasClass(e, className)) {
            var cn = e.getAttribute('class');
            e.setAttribute('class', cn ? cn + ' ' + className : className);
        }
    }
    /**
     * Adds or removes a class to or from an element.
     *
     * @param e Element that will have the class added.
     * @param className Class to add or remove.
     * @param addOrRemove Whether to add or remove the class.
     * Use true to add class to element and false to remove class from element.
     */
    export function toggleClass(e: HTMLElement, className: string, addOrRemove: boolean) {
        if (addOrRemove) {
            addClass(e, className);
        } else {
            removeClass(e, className);
        }
    }
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
    export function setSelectionRange(e: HTMLInputElement, start: number, end = start) {
        e = asType(e, HTMLInputElement);
        if (contains(document.body, e) && !e.disabled && e.style.display != 'none') {
            try {
                e.focus(); // needed in Chrome (TFS 124102, 142672)
                e.setSelectionRange(asNumber(start), asNumber(end));
            } catch (x) { }
        }
    }
    /**
     * Gets a reference to the element that contains the focus,
     * accounting for shadow document fragments.
     */
    export function getActiveElement() {
        var ae = <HTMLElement>document.activeElement;
        if (ae) {
            // account for shadowRoot: https://github.com/w3c/webcomponents/issues/358)
            var shadowRoot = ae['shadowRoot'];
            if (shadowRoot && shadowRoot.activeElement) {
                ae = shadowRoot.activeElement;
            }
        }
        return ae;
    }
    /**
     * Moves the focus to the next/previous/first focusable child element.
     *
     * @param parent Parent element.
     * @param offset Offset to use when moving the focus (use zero to focus on the first focusable child).
     */
    export function moveFocus(parent: HTMLElement, offset: number) {

        // build array of focusable elements (including divs and spans)
        var inputs = parent.querySelectorAll('input,select,textarea,button,a,div,span'),
            focusable = [];
        for (var i = 0; i < inputs.length; i++) {
            var el = <HTMLInputElement>inputs[i];
            if (!el.disabled && el.tabIndex > -1 && el.clientHeight > 0 && // ignore disabled, unfocusable, hidden
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
            var el = <HTMLInputElement>focusable[index];
            el.focus();
            if (el instanceof HTMLInputElement) {
                el.select(); // TFS 190336
            }
        }
    }

    // ** jQuery replacement methods

    /**
     * Gets an element from a jQuery-style selector.
     *
     * @param selector An element, a query selector string, or a jQuery object.
     */
    export function getElement(selector: any): HTMLElement {
        if (selector instanceof Element) return selector;
        if (isString(selector)) return <HTMLElement>document.querySelector(selector);
        if (selector && selector.jquery) return selector[0];
        return null;
    }
    /**
     * Creates an element from an HTML string.
     *
     * @param html HTML fragment to convert into an HTMLElement.
     * @param appendTo Optional HTMLElement to append the new element to.
     * @return The new element.
     */
    export function createElement(html: string, appendTo?: HTMLElement): HTMLElement {
        var div = document.createElement('div');
        div.innerHTML = html;
        var e = <HTMLElement>div.removeChild(div.firstChild);
        return (appendTo instanceof HTMLElement)
            ? <HTMLElement>appendTo.appendChild(e)
            : e;
    }
    /**
     * Sets the text content of an element.
     *
     * @param e Element that will have its content updated.
     * @param text Plain text to be assigned to the element.
     */
    export function setText(e: HTMLElement, text: string) {

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
        } else if (fc || text) {
            e.textContent = text; // something else, set the textContent
        }
    }
    /**
     * Checks whether an HTML element contains another.
     *
     * @param parent Parent element.
     * @param child Child element.
     * @return True if the parent element contains the child element.
     */
    export function contains(parent: any, child: any): boolean {
        for (var e = <Node>child; e && parent;) {
            if (e === parent) return true; // found!
            e = e.parentNode || e['host']; // move up to parent node or host (shadow DOM)
        }
        return false;
    }
    /**
     * Finds the closest ancestor that satisfies a selector.
     *
     * @param e Element where the search should start.
     * @param selector A string containing a selector expression to match elements against.
     * @return The closest ancestor that satisfies the selector (including the original element), or null if not found.
     */
    export function closest(e: any, selector: string): Node {
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
    /**
     * Enables or disables an element.
     *
     * @param e Element to enable or disable.
     * @param value Whether to enable or disable the element.
     */
    export function enable(e: HTMLElement, value: boolean) {

        // update wj-state-disabled class and disabled attribute on the element
        toggleClass(e, 'wj-state-disabled', !value);
        if (value) {
            e.removeAttribute('disabled');
        } else {
            e.setAttribute('disabled', 'true');
        }

        // update disabled attribute on inner input elements (TFS 190939)
        var inputs = e.querySelectorAll('input');
        for (var i = 0; i < inputs.length; i++) {
            var input = <HTMLElement>inputs[i];
            if (value) {
                input.removeAttribute('disabled');
            } else {
                input.setAttribute('disabled', 'true');
            }
        }
    }
    /**
     * Gets the bounding rectangle of an element in page coordinates.
     *
     * This is similar to the <b>getBoundingClientRect</b> function,
     * except that uses window coordinates, which change when the
     * document scrolls.
     */
    export function getElementRect(e: Element): Rect {
        var rc = e.getBoundingClientRect();
        return new Rect(rc.left + pageXOffset, rc.top + pageYOffset, rc.width, rc.height);
    }
    /**
     * Modifies the style of an element by applying the properties specified in an object.
     *
     * @param e Element or array of elements whose style will be modified.
     * @param css Object containing the style properties to apply to the element.
     */
    export function setCss(e: any, css: any) {

        // apply to arrays
        if (isArray(e)) {
            e.forEach((item) => {
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
    export function animate(apply: Function, duration = 400, step = 35): number {
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
            if (pct >= 1) { // done!
                clearInterval(timer);
            }
        }, step);
        return timer;
    }


    // ** utility classes

    /**
     * Class that represents a point (with x and y coordinates).
     */
    export class Point {
        /**
         * Gets or sets the x coordinate of this @see:Point.
         */
        x: number;
        /**
         * Gets or sets the y coordinate of this @see:Point.
         */
        y: number;
        /**
         * Initializes a new instance of the @see:Point class.
         *
         * @param x X coordinate of the new Point.
         * @param y Y coordinate of the new Point.
         */
        constructor(x: number = 0, y: number = 0) {
            this.x = asNumber(x);
            this.y = asNumber(y);
        }
        /**
         * Returns true if a @see:Point has the same coordinates as this @see:Point.
         *
         * @param pt @see:Point to compare to this @see:Point.
         */
        equals(pt: Point): boolean {
            return (pt instanceof Point) && this.x == pt.x && this.y == pt.y;
        }
        /**
         * Creates a copy of this @see:Point.
         */
        clone(): Point {
            return new Point(this.x, this.y);
        }
    }

    /**
     * Class that represents a size (with width and height).
     */
    export class Size {
        /**
         * Gets or sets the width of this @see:Size.
         */
        width: number;
        /**
         * Gets or sets the height of this @see:Size.
         */
        height: number;
        /**
         * Initializes a new instance of the @see:Size class.
         *
         * @param width Width of the new @see:Size.
         * @param height Height of the new @see:Size.
         */
        constructor(width = 0, height = 0) {
            this.width = asNumber(width);
            this.height = asNumber(height);
        }
        /**
         * Returns true if a @see:Size has the same dimensions as this @see:Size.
         *
         * @param sz @see:Size to compare to this @see:Size.
         */
        equals(sz: Size): boolean {
            return (sz instanceof Size) && this.width == sz.width && this.height == sz.height;
        }
        /**
         * Creates a copy of this @see:Size.
         */
        clone(): Size {
            return new Size(this.width, this.height);
        }
    }

    /**
     * Class that represents a rectangle (with left, top, width, and height).
     */
    export class Rect {
        /**
         * Gets or sets the left coordinate of this @see:Rect.
         */
        left: number;
        /**
         * Gets or sets the top coordinate of this @see:Rect.
         */
        top: number;
        /**
         * Gets or sets the width of this @see:Rect.
         */
        width: number;
        /**
         * Gets or sets the height of this @see:Rect.
         */
        height: number;
        /**
         * Initializes a new instance of the @see:Rect class.
         *
         * @param left Left coordinate of the new @see:Rect.
         * @param top Top coordinate of the new @see:Rect.
         * @param width Width of the new @see:Rect.
         * @param height Height of the new @see:Rect.
         */
        constructor(left: number, top: number, width: number, height: number) {
            this.left = asNumber(left);
            this.top = asNumber(top);
            this.width = asNumber(width);
            this.height = asNumber(height);
        }
        /**
         * Gets the right coordinate of this @see:Rect.
         */
        get right(): number {
            return this.left + this.width;
        }
        /**
         * Gets the bottom coordinate of this @see:Rect.
         */
        get bottom(): number {
            return this.top + this.height;
        }
        /**
         * Returns true if a @see:Rect has the same coordinates and dimensions
         * as this @see:Rect.
         *
         * @param rc @see:Rect to compare to this @see:Rect.
         */
        equals(rc: Rect): boolean {
            return (rc instanceof Rect) && this.left == rc.left && this.top == rc.top && this.width == rc.width && this.height == rc.height;
        }
        /**
         * Creates a copy of this @see:Rect.
         */
        clone(): Rect {
            return new Rect(this.left, this.top, this.width, this.height);
        }
        /**
         * Creates a @see:Rect from <b>ClientRect</b> or <b>SVGRect</b> objects.
         *
         * @param rc Rectangle obtained by a call to the DOM's <b>getBoundingClientRect</b>
         * or <b>GetBoundingBox</b> methods.
         */
        static fromBoundingRect(rc: any): Rect {
            if (rc.left != null) {
                return new Rect(rc.left, rc.top, rc.width, rc.height);
            } else if (rc.x != null) {
                return new Rect(rc.x, rc.y, rc.width, rc.height);
            } else {
                assert(false, 'Invalid source rectangle.');
        }
        }
        /**
         * Gets a rectangle that represents the union of two rectangles.
         *
         * @param rc1 First rectangle.
         * @param rc2 Second rectangle.
         */
        static union(rc1: Rect, rc2: Rect): Rect {
            var x = Math.min(rc1.left, rc2.left),
                y = Math.min(rc1.top, rc2.top),
                right = Math.max(rc1.right, rc2.right),
                bottom = Math.max(rc1.bottom, rc2.bottom);
            return new Rect(x, y, right - x, bottom - y);
        }
        /**
         * Gets a rectangle that represents the intersection of two rectangles.
         *
         * @param rc1 First rectangle.
         * @param rc2 Second rectangle.
         */
        static intersection(rc1: Rect, rc2: Rect): Rect {
            var x = Math.max(rc1.left, rc2.left),
                y = Math.max(rc1.top, rc2.top),
                right = Math.min(rc1.right, rc2.right),
                bottom = Math.min(rc1.bottom, rc2.bottom);
            return new Rect(x, y, right - x, bottom - y);
        }
        /**
         * Determines whether the rectangle contains a given point or rectangle.
         *
         * @param pt The @see:Point or @see:Rect to ckeck.
         */
        contains(pt: any): boolean {
            if (pt instanceof Point) {
                return pt.x >= this.left && pt.x <= this.right &&
                    pt.y >= this.top && pt.y <= this.bottom;
            } else if (pt instanceof Rect) {
                var rc2 = <Rect>pt;
                return rc2.left >= this.left && rc2.right <= this.right &&
                    rc2.top >= this.top && rc2.bottom <= this.bottom;
            } else {
                assert(false, 'Point or Rect expected.');
            }
        }
        /**
         * Creates a rectangle that results from expanding or shrinking a rectangle by the specified amounts.
         *
         * @param dx The amount by which to expand or shrink the left and right sides of the rectangle.
         * @param dy The amount by which to expand or shrink the top and bottom sides of the rectangle.
         */
        inflate(dx: number, dy: number): Rect {
            return new Rect(this.left - dx, this.top - dy, this.width + 2 * dx, this.height + 2 * dy);
        }
    }

    /**
     * Provides date and time utilities.
     */
    export class DateTime {

        /**
         * Gets a new Date that adds the specified number of days to a given Date.
         *
         * @param value Original date.
         * @param days Number of days to add to the given date.
         */
        static addDays(value: Date, days: number): Date {
            return new Date(value.getFullYear(), value.getMonth(), value.getDate() + days);
        }
        /**
         * Gets a new Date that adds the specified number of months to a given Date.
         *
         * @param value Original date.
         * @param months Number of months to add to the given date.
         */
        static addMonths(value: Date, months: number): Date {
            return new Date(value.getFullYear(), value.getMonth() + months, value.getDate());
        }
        /**
         * Gets a new Date that adds the specified number of years to a given Date.
         *
         * @param value Original date.
         * @param years Number of years to add to the given date.
         */
        static addYears(value: Date, years: number): Date {
            return new Date(value.getFullYear() + years, value.getMonth(), value.getDate());
        }
        /**
         * Gets a new Date that adds the specified number of hours to a given Date.
         *
         * @param value Original date.
         * @param hours Number of hours to add to the given date.
         */
        static addHours(value: Date, hours: number): Date {
            return new Date(value.getFullYear(), value.getMonth(), value.getDate(), value.getHours() + hours);
        }
        /**
         * Gets a new Date that adds the specified number of minutes to a given Date.
         *
         * @param value Original date.
         * @param minutes Number of minutes to add to the given date.
         */
        static addMinutes(value: Date, minutes: number): Date {
            return new Date(value.getFullYear(), value.getMonth(), value.getDate(), value.getHours(), value.getMinutes() + minutes);
        }
        /**
         * Gets a new Date that adds the specified number of seconds to a given Date.
         *
         * @param value Original date.
         * @param seconds Number of seconds to add to the given date.
         */
        static addSeconds(value: Date, seconds: number): Date {
            return new Date(value.getFullYear(), value.getMonth(), value.getDate(), value.getHours(), value.getMinutes(), value.getSeconds() + seconds);
        }
        /**
         * Returns true if two Date objects refer to the same date (ignoring time).
         *
         * @param d1 First date.
         * @param d2 Second date.
         */
        static sameDate(d1: Date, d2: Date): boolean {
            return isDate(d1) && isDate(d2) &&
                d1.getFullYear() == d2.getFullYear() &&
                d1.getMonth() == d2.getMonth() &&
                d1.getDate() == d2.getDate();
        }
        /**
         * Returns true if two Date objects refer to the same time (ignoring date).
         *
         * @param d1 First date.
         * @param d2 Second date.
         */
        static sameTime(d1: Date, d2: Date): boolean {
            return isDate(d1) && isDate(d2) &&
                d1.getHours() == d2.getHours() &&
                d1.getMinutes() == d2.getMinutes() &&
                d1.getSeconds() == d2.getSeconds();
        }
        /**
         * Returns true if two Date objects refer to the same date and time.
         *
         * @param d1 First date.
         * @param d2 Second date.
         */
        static equals(d1: Date, d2: Date): boolean {
            return isDate(d1) && isDate(d2) && d1.getTime() == d2.getTime();
        }
        /**
         * Gets a Date object with the date and time set on two Date objects.
         *
         * @param date Date object that contains the date (day/month/year).
         * @param time Date object that contains the time (hour:minute:second).
         */
        static fromDateTime(date: Date, time: Date): Date {
            if (!date && !time) return null;
            if (!date) date = time;
            if (!time) time = date;
            return new Date(
                date.getFullYear(), date.getMonth(), date.getDate(),
                time.getHours(), time.getMinutes(), time.getSeconds());
        }
        /**
         * Converts a calendar date to a fiscal date using the current culture.
         *
         * @param date Calendar date.
         * @param govt Whether to use the government or corporate fiscal year.
         */
        static toFiscal(date: Date, govt: boolean) {
            var cal = wijmo.culture.Globalize.calendar;
            return isArray(cal.fiscalYearOffsets)
                ? DateTime.addMonths(date, -cal.fiscalYearOffsets[govt ? 0 : 1])
                : date;
        }
        /**
         * Converts a fiscal year date to a calendar date using the current culture.
         *
         * @param date Fiscal year date.
         * @param govt Whether to use the government or corporate fiscal year.
         */
        static fromFiscal(date: Date, govt: boolean) {
            var cal = wijmo.culture.Globalize.calendar;
            return isArray(cal.fiscalYearOffsets)
                ? DateTime.addMonths(date, +cal.fiscalYearOffsets[govt ? 0 : 1])
                : date;
        }
        /**
         * Gets a new Date object set to today's date and no time portion.
         */
        static newDate(): Date {
            var dt = new Date();
            return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
        }
        /**
         * Creates a copy of a given Date object.
         *
         * @param date Date object to copy.
         */
        static clone(date: Date): Date {
            return DateTime.fromDateTime(date, date);
        }
    }

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
    export function httpRequest(url: string, settings?: any): XMLHttpRequest {
        if (!settings) settings = {};

        // select method and basic options
        var method = settings.method ? asString(settings.method).toUpperCase() : 'GET',
            async = settings.async != null ? asBoolean(settings.async) : true,
            data = settings.data;

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
                } else if (settings.error) {
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
            } else {
                throw 'HttpRequest Error: ' + xhr.status + ' ' + xhr.statusText;
            }
        };

        // send the request
        xhr.open(method, url, async, settings.user, settings.password);
        if (settings.user && settings.password) {
            xhr.setRequestHeader('Authorization', 'Basic ' + btoa(settings.user + ':' + settings.password))
        }
        if (isJson) {
            xhr.setRequestHeader('Content-Type', 'application/json');
        }
        if (settings.requestHeaders) {
            for (var key in settings.requestHeaders) {
                xhr.setRequestHeader(key, settings.requestHeaders[key])
            }
        }
        if (isFunction(settings.beforeSend)) {
            settings.beforeSend(xhr);
        }
        xhr.send(data);

        // return the request
        return xhr;
    }
}
module wijmo {
    'use strict';

    /**
     * Gets or sets an object that contains all localizable strings in the Wijmo library.
     *
     * The culture selector is a two-letter string that represents an 
     * <a href='http://en.wikipedia.org/wiki/List_of_ISO_639-1_codes'>ISO 639 culture</a>. 
     */
    export var culture: any = {
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
    export class Globalize {

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
        static format(value: any, format: string, trim?: boolean, truncate?: boolean): string {

            // format numbers and dates, convert others to string
            if (isString(value)) {
                return value;
            } else if (isNumber(value)) {
                format = format || (value == Math.round(value) ? 'n0' : 'n2'); 
                return Globalize.formatNumber(value, format, trim, truncate);
            } else if (isDate(value)) {
                format = format || 'd';
                return Globalize.formatDate(value, format);
            } else {
                return value != null ? value.toString() : '';
            }
        }
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
        static formatNumber(value: number, format: string, trim?: boolean, truncate?: boolean): string {
            asNumber(value);
            asString(format);

            var result,
                m = format ? format.match(/([a-z])(\d*)(,*)(.*)/i) : null,
                nf = culture.Globalize.numberFormat,
                f1 = m ? m[1].toLowerCase() : 'n',
                prec = (m && m[2]) ? parseInt(m[2]) : (f1 == 'c') ? nf.currency.decimals : value == Math.round(value) ? 0 : 2,
                scale = (m && m[3]) ? 3 * m[3].length : 0,
                curr = (m && m[4]) ? m[4] : nf.currency.symbol,
                dp = nf['.'],
                ts = nf[','];

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
                value = toFixed(value * 100, prec, truncate); // TFS 147585
            }

            // truncate value
            if (truncate) {
                value = toFixed(value, prec, true);
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
                var idx = result.indexOf(dp),
                    rx = /\B(?=(\d\d\d)+(?!\d))/g;
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
        }
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
        static formatDate(value: Date, format: string): string {
            value = asDate(value);

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
            var parts = Globalize._parseDateFormat(format),
                str = '';
            for (var i = 0; i < parts.length; i++) {
                str += Globalize._formatDatePart(value, format, parts[i]);
            }

            // all done
            return str;
        }
        /**
         * Parses a string into an integer.
         *
         * @param value String to convert to an integer.
         * @param format Format to use when parsing the number.
         * @return The integer represented by the given string, 
         * or <b>NaN</b> if the string cannot be parsed into an integer.
         */
        static parseInt(value: string, format?: string): number {
            return Math.round(Globalize.parseFloat(value, format));
        }
        /**
         * Parses a string into a floating point number.
         *
         * @param value String to convert to a number.
         * @param format Format to use when parsing the number.
         * @return The floating point number represented by the given string, 
         * or <b>NaN</b> if the string cannot be parsed into a floating point number.
         */
        static parseFloat(value: string, format?: string): number {
            var neg = value.indexOf('-') > -1 || (value.indexOf('(') > -1 && value.indexOf(')') > -1) ? -1 : +1,
                pct = value.indexOf('%') > -1 ? .01 : 1,
                m = format ? format.match(/,+/) : null,
                scale = m ? m[0].length * 3 : 0;

            // hex
            if (format && (format[0] == 'x' || format[0] == 'X')) {
                value = value.replace(/[^0-9a-f]+.*$/gi, ''); // truncate at first invalid char
                return parseInt(value, 16) * neg * pct * Math.pow(10, scale);
            }

            // decimal
            var dp = culture.Globalize.numberFormat['.'],
                rx = new RegExp('[^\\d\\' + dp + ']', 'g'),
                val = value.replace(rx, '').replace(dp, '.'); // remove non-digits, replace decimal point
            return parseFloat(val) * neg * pct * Math.pow(10, scale);
        }
        /**
         * Parses a string into a Date.
         *
         * @param value String to convert to a Date.
         * @param format Format string used to parse the date.
         * @return The date represented by the given string, or null if the string
         * cannot be parsed into a Date.
         */
        static parseDate(value: string, format: string): Date {

            // make sure we have a value
            value = asString(value);
            if (!value) {
                return null;
            }

            // culture-invariant formats
            if (format == 'u') {
                return new Date(value);
            }

            // parse using RFC 3339 pattern ([yyyy-MM-dd] [hh:mm[:ss]])
            var d: Date;
            if (format == 'R' || format == 'r') {
                var rx = /(([0-9]+)\-([0-9]+)\-([0-9]+))?\s?(([0-9]+):([0-9]+)(:([0-9]+))?)?/,
                    match = value.match(rx);
                if (match[1] || match[5]) {
                    var d = match[1] // parse date
                        ? new Date(parseInt(match[2]), parseInt(match[3]) - 1, parseInt(match[4]))
                        : new Date();
                    if (match[5]) { // parse time
                        d.setHours(parseInt(match[6]));
                        d.setMinutes(parseInt(match[7]));
                        d.setSeconds(match[8] ? parseInt(match[9]) : 0);
                    }
                } else {
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
            var cal = culture.Globalize.calendar,
                cjk = Globalize._CJK,
                rxv = new RegExp(
                    '(\\' + cal['/'] + ')|(\\' + cal[':'] + ')|' + // date/time separators
                    '(\\d+)|' + // digits
                    '([' + cjk + '\\.]{2,})|' + // strings with dots
                    '([' + cjk + ']+)', // strings with no dots
                    'gi'),
                vparts = value.match(rxv),
                fparts = Globalize._parseDateFormat(format), offset = 0,
                year = -1, month = 0, day = 1, hour = 0, min = 0, tzm = 0, sec = 0, ms = 0, era = -1,
                hasDayName, hasDay, hasQuarter, hasMonth, fiscalFmt;

            // basic validation (TFS 81465, 128359)
            if (!vparts || !vparts.length || !fparts || !fparts.length) {
                return null;
            }

            // parse each element
            for (var i = 0; i < fparts.length && vparts; i++) {
                var vpi = i - offset,
                    pval = (vpi > -1 && vpi < vparts.length) ? vparts[vpi] : '',
                    plen = fparts[i].length;
                switch (fparts[i]) {

                    // ** year
                    case 'EEEE': case 'EEE': case 'EE': case 'E': // fiscal (govt)
                    case 'eeee': case 'eee': case 'ee': case 'e': // fiscal (corp)
                        fiscalFmt = fparts[i];
                        // ** fall through **
                    case 'yyyy': case 'yyy': case 'yy': case 'y': // calendar
                        if (plen > 1 && pval.length > plen) {
                            vparts[vpi] = pval.substr(plen);
                            pval = pval.substr(0, plen);
                            offset++;
                        }
                        year = parseInt(pval);
                        break;

                    // ** month
                    case 'MMMM': case 'MMM':
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
                    case 'MM': case 'M':
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
                    case 'dd': case 'd':
                        if (plen > 1 && pval.length > plen) {
                            vparts[vpi] = pval.substr(plen);
                            pval = pval.substr(0, plen);
                            offset++;
                        }
                        day = parseInt(pval);
                        hasDay = true;
                        break;

                    // ** hour
                    case 'hh': case 'h':
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
                    case 'mm': case 'm':
                        if (plen > 1 && pval.length > plen) {
                            vparts[vpi] = pval.substr(plen);
                            pval = pval.substr(0, plen);
                            offset++;
                        }
                        min = parseInt(pval);
                        break;

                    // ** second
                    case 'ss': case 's':
                        if (plen > 1 && pval.length > plen) {
                            vparts[vpi] = pval.substr(plen);
                            pval = pval.substr(0, plen);
                            offset++;
                        }
                        sec = parseInt(pval);
                        break;

                    // ** millisecond
                    case 'fffffff': case 'FFFFFFF':
                    case 'ffffff': case 'FFFFFF':
                    case 'fffff': case 'FFFFF':
                    case 'ffff': case 'FFFF':
                    case 'fff': case 'FFF':
                    case 'ff': case 'FF':
                    case 'f': case 'F':
                        ms = parseInt(pval) / Math.pow(10, plen - 3);
                        break;

                    // ** am/pm
                    case 'tt': case 't':
                        pval = pval.toUpperCase();
                        if ((cal.pm[0] && pval == cal.pm[0] && hour < 12) ||
                            (cal.pm[1] && pval == cal.pm[1] && hour < 12)) {
                            hour += 12;
                        }
                        break;

                    // ** quarter
                    case 'q': case 'Q': case 'u': case 'U':
                        hasQuarter = true;
                        break;

                    // ** era
                    case 'ggg': case 'gg': case 'g':
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
                if (isNaN(hour)) hour = 0;
                if (isNaN(min)) min = 0;
                if (isNaN(sec)) sec = 0;
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
                if (!hasMonth) { // need year and month to convert fiscal to calendar
                    return null;
                }
                d = new Date(year, month);
                d = DateTime.fromFiscal(d, fiscalFmt[0] == 'E');
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
            } else if (year < 100) {
                year += year >= 30 ? 1900 : 2000;
            }

            // return result
            d = new Date(year, month, day, hour, min + tzm, sec, ms);
            return isNaN(d.getTime()) ? null : d;
        }

        // Chinese/Japanese/Korean characters
        // http://www.rikai.com/library/kanjitables/kanji_codes.unicode.shtml
        // http://www.programminginkorean.com/programming/hangul-in-unicode/
        // NOTE: using 'replace' to keep minifier from switching the escaped Unicode chars into real Unicode.
        // NOTE: will have to expand for other cultures? (Arabic/Hebrew/Hindi etc?)
        static _CJK = 'a-z' +
                      'u00c0-u017fu3000-u30ffu4e00-u9faf'.replace(/u/g, '\\u') + // cj
                      'u1100-u11ffu3130-u318fua960-ua97fuac00-ud7afud7b0-ud7ff'.replace(/u/g, '\\u'); // k

        /**
         * Gets the first day of the week according to the current culture.
         *
         * The value returned is between zero (Sunday) and six (Saturday).
         */
        static getFirstDayOfWeek(): number {
            var fdw = culture.Globalize.calendar.firstDay;
            return fdw ? fdw : 0;
        }
        /**
         * Gets the symbol used as a decimal separator in numbers.
         */
        static getNumberDecimalSeparator(): string {
            var ndc = culture.Globalize.numberFormat['.'];
            return ndc ? ndc : '.';
        }

        // ** implementation

        // unquotes a string
        private static _unquote(s: string): string {
            if (s.length > 1 && s[0] == s[s.length - 1]) {
                if (s[0] == '\'' || s[0] == '\"') {
                    return s.substr(1, s.length - 2);
                }
            }
            return s;
        }

        // parse a date format string into its parts
        private static _dateFomatParts = {};
        private static _parseDateFormat(format: string): string[] {

            // use cache whenever possible
            if (format in Globalize._dateFomatParts) {
                return Globalize._dateFomatParts[format];
            }

            // parse the format
            var parts = [],
                str = '',
                start, end;
            for (start = 0; start > -1 && start < format.length; start++) {
                var c = format[start];
                if (c == '\'' || c == '"') { // handle quoted parts
                    end = format.indexOf(c, start + 1); // keep quotes to distinguish from regular date parts
                    if (end > -1) {
                        parts.push(format.substring(start, end + 1));
                        start = end;
                        continue;
                    }
                }
                end = start + 1;
                for (; end < format.length; end++) {
                    if (format[end] != c) break;
                }
                parts.push(format.substring(start, end));
                start = end - 1;
            }

            // cache and return
            Globalize._dateFomatParts[format] = parts;
            return parts;
        }

        // format a date part into a string
        private static _formatDatePart(d: Date, format: string, part: string): string {
            var cal = culture.Globalize.calendar,
                era = 0, year = 0, ff = 0, fd,
                plen = part.length;
            switch (part) {

                // ** year
                case 'yyyy': case 'yyy': case 'yy': case 'y': // calendar year
                case 'EEEE': case 'EEE': case 'EE': case 'E': // fiscal year (govt)
                case 'eeee': case 'eee': case 'ee': case 'e': // fiscal year (corporate)

                    // get the year (calendar or fiscal)
                    fd = part[0] == 'E' ? DateTime.toFiscal(d, true) :
                        part[0] == 'e' ? DateTime.toFiscal(d, false) :
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
                case 'MM': case 'M':
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
                case 'hh': case 'h':
                    return Globalize._zeroPad(Globalize._h12(d), plen);
                case 'HH': case 'H':
                    return Globalize._zeroPad(d.getHours(), plen);

                // ** minute
                case 'mm': case 'm':
                    return Globalize._zeroPad(d.getMinutes(), plen);

                // ** second
                case 'ss': case 's':
                    return Globalize._zeroPad(d.getSeconds(), plen);

                // ** millisecond
                case 'fffffff': case 'FFFFFFF':
                case 'ffffff': case 'FFFFFF':
                case 'fffff': case 'FFFFF':
                case 'ffff': case 'FFFF':
                case 'fff': case 'FFF':
                case 'ff': case 'FF':
                case 'f': case 'F':
                    ff = d.getMilliseconds() * Math.pow(10, plen - 3);
                    return part[0] == 'f' ? Globalize._zeroPad(ff, plen) : ff.toFixed(0);

                // ** am/pm
                case 'tt':
                    return d.getHours() < 12 ? cal.am[0] : cal.pm[0];
                case 't':
                    return d.getHours() < 12 ? cal.am[1] : cal.pm[1];

                // ** quarter
                case 'q': case 'Q': // calendar
                    return (Math.floor(d.getMonth() / 3) + 1).toString();
                case 'u': case 'U': // quarter (U: govt; u: corp)
                    fd = DateTime.toFiscal(d, part == 'U');
                    return (Math.floor(fd.getMonth() / 3) + 1).toString();

                // ** era
                case 'ggg': case 'gg': case 'g':
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
        }

        // get a date's era (used only in Japanese locales)
        private static _getEra(d: any, cal: any): number {
            if (isDate(d)) { // find era by start date
                for (var i = 0; i < cal.eras.length; i++) {
                    if (d >= cal.eras[i].start) {
                        return i;
                    }
                }
            } else if (isString(d)) { // find era by name or symbol
                for (var i = 0; i < cal.eras.length; i++) {
                    if (cal.eras[i].name) {
                        if (cal.eras[i].name.indexOf(d) == 0 || cal.eras[i].symbol.indexOf(d) == 0) {
                            return i;
                        }
                    }
                }
            }
            return -1; // not found
        }

        // expand date pattern into full date format
        private static _expandFormat(format: string): string {
            var fmt = culture.Globalize.calendar.patterns[format];
            return fmt ? fmt : format;
        }

        // format a number with leading zeros
        private static _zeroPad(num: number, places: number) {
            var n = num.toFixed(0),
                zero = places - n.length + 1;
            return zero > 0 ? Array(zero).join('0') + n : n;
        }

        // format an hour to 12 or 24 hour base depending on the calendar
        private static _h12(d: Date) {
            var cal = culture.Globalize.calendar,
                h = d.getHours();
            if (cal.am && cal.am[0]) {
                h = h % 12;
                if (h == 0) h = 12;
            }
            return h;
        }
    }
}
module wijmo {
    'use strict';

    /**
     * Provides binding to complex properties (e.g. 'customer.address.city')
     */
    export class Binding {
        _path: string;
        _parts: any[];
        _key: string;

        /**
         * Initializes a new instance of the @see:Binding class.
         *
         * @param path Name of the property to bind to.
         */
        constructor(path: string) {
            this.path = path;
        }

        /**
         * Gets or sets the path for the binding.
         * 
         * In the simplest case, the path is the name of the property of the source 
         * object to use for the binding (e.g. 'street').
         *
         * Sub-properties of a property can be specified by a syntax similar to that 
         * used in JavaScript (e.g. 'address.street').
         */
        get path(): string {
            return this._path;
        }
        set path(value: string) {
            this._path = value;
            this._parts = value ? value.split('.') : []; // e.g. 'customer.balance'
            for (var i = 0; i < this._parts.length; i++) {
                var part = this._parts[i],
                    ib = part.indexOf('['); // e.g. 'customer.balance[0]'
                if (ib > -1) {
                    this._parts[i] = part.substr(0, ib);
                    this._parts.splice(++i, 0, parseInt(part.substr(ib + 1)));
                }
            }
            this._key = this._parts.length == 1 ? this._parts[0] : null;
        }
        /**
         * Gets the binding value for a given object.
         *
         * If the object does not contain the property specified by the
         * binding @see:path, the method returns null.
         *
         * @param object The object that contains the data to be retrieved.
         */
        getValue(object: any): any {
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
        }
        /**
         * Sets the binding value on a given object.
         *
         * If the object does not contain the property specified by the
         * binding @see:path, the value is not set.
         *
         * @param object The object that contains the data to be set.
         * @param value Data value to set.
         */
        setValue(object: any, value: any) {
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
        }
    }
}
module wijmo {
    'use strict';

    /**
     * Represents an event handler.
     *
     * Event handlers are functions invoked when events are raised.
     *
     * Every event handler has two arguments:
     * <ul>
     *   <li><b>sender</b> is the object that raised the event, and</li>
     *   <li><b>args</b> is an optional object that contains the event parameters.</li>
     * </ul>
     */
    export interface IEventHandler {
        (sender: any, args: EventArgs): void;
    }
    /*
     * Represents an event handler (private class)
     */
    class EventHandler {
        handler: IEventHandler;
        self: any;
        constructor(handler: IEventHandler, self: any) {
            this.handler = handler;
            this.self = self;
        }
    }
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
    export class Event {
        private _handlers: EventHandler[] = [];

        /**
         * Adds a handler to this event.
         *
         * @param handler Function invoked when the event is raised.
         * @param self Object that defines the event handler 
         * (accessible as 'this' from the handler code).
         */
        addHandler(handler: IEventHandler, self?: any) {
            asFunction(handler);
            this._handlers.push(new EventHandler(handler, self));
        }
        /**
         * Removes a handler from this event.
         *
         * @param handler Function invoked when the event is raised.
         * @param self Object that defines the event handler (accessible as 'this' from the handler code).
         */
        removeHandler(handler: IEventHandler, self?: any) {
            asFunction(handler);
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
        }
        /**
         * Removes all handlers associated with this event.
         */
        removeAllHandlers() {
            this._handlers.length = 0;
        }
        /**
         * Raises this event, causing all associated handlers to be invoked.
         *
         * @param sender Source object.
         * @param args Event parameters. 
         */
        raise(sender: any, args = EventArgs.empty) {
            for (var i = 0; i < this._handlers.length; i++) {
                var l = this._handlers[i];
                l.handler.call(l.self, sender, args);
            }
        }
        /**
         * Gets a value that indicates whether this event has any handlers.
         */
        get hasHandlers(): boolean {
            return this._handlers.length > 0;
        }
    }
    /**
     * Base class for event arguments.
     */
    export class EventArgs {
        /**
         * Provides a value to use with events that do not have event data.
         */
        static empty = new EventArgs();
    }
    /**
     * Provides arguments for cancellable events.
     */
    export class CancelEventArgs extends EventArgs {
        /**
         * Gets or sets a value that indicates whether the event should be canceled.
         */
        cancel = false;
    }
    /**
     * Provides arguments for property change events.
     */
    export class PropertyChangedEventArgs extends EventArgs {
        _name: string;
        _oldVal: any;
        _newVal: any;

        /**
         * Initializes a new instance of the @see:PropertyChangedEventArgs class.
         *
         * @param propertyName The name of the property whose value changed.
         * @param oldValue The old value of the property.
         * @param newValue The new value of the property.
         */
        constructor(propertyName: string, oldValue: any, newValue: any) {
            super();
            this._name = propertyName;
            this._oldVal = oldValue;
            this._newVal = newValue;
        }
        /**
         * Gets the name of the property whose value changed.
         */
        get propertyName(): string {
            return this._name;
        }
        /**
         * Gets the old value of the property.
         */
        get oldValue(): any {
            return this._oldVal;
        }
        /**
         * Gets the new value of the property.
         */
        get newValue(): any {
            return this._newVal;
        }
    }
    /**
     * Provides arguments for @see:XMLHttpRequest error events.
     */
    export class RequestErrorEventArgs extends CancelEventArgs {
        _xhr: XMLHttpRequest;

        /**
         * Initializes a new instance of the @see:RequestErrorEventArgs class.
         *
         * @param xhr The @see:XMLHttpRequest that detected the error.
         * The status and statusText properties of the request object 
         * contain details about the error.
         */
        constructor(xhr: XMLHttpRequest) {
            super();
            this._xhr = xhr;
        }
        /**
         * Gets a reference to the @see:XMLHttpRequest that detected the error.
         *
         * The status and statusText properties of the request object contain
         * details about the error.
         */
        get request(): XMLHttpRequest {
            return this._xhr;
        }
    }
}
module wijmo {
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
    export class Control {
        private static _DATA_KEY = 'wj-Control';    // key used to store control reference in host element
        private static _REFRESH_INTERVAL = 10;      // interval between invalidation and refresh
        private static _wme: HTMLElement;           // watermark element
        static _touching: boolean;                  // the current event is a touch event

        private _focus = false;                     // whether the control currently contains the focus
        private _updating = 0;                      // update count (no refreshes while > 0)
        private _fullUpdate = false;                // in case there are multiple calls to invalidate(x)
        private _toInv: number;                     // invalidation timeOut
        private _szCtl: Size;                       // current control size
        private _e: HTMLElement;                    // host element
        private _orgOuter: string;                  // host element's original outerHTML
        private _orgInner: string;                  // host element's original innerHTML
        private _listeners;                         // list of event listeners attached to this control
        protected _orgTag: string;                  // host element's original tag

        /**
         * Initializes a new instance of the @see:Control class and attaches it to a DOM element.
         * 
         * @param element The DOM element that will host the control, or a selector for the host element (e.g. '#theCtrl').
         * @param options JavaScript object containing initialization data for the control.
         * @param invalidateOnResize Whether the control should be invalidated when it is resized.
         */
        constructor(element: any, options = null, invalidateOnResize = false) {

            // check that the element is not in use
            assert(Control.getControl(element) == null, 'Element is already hosting a control.');

            // get the host element
            var host = getElement(element);
            assert(host != null, 'Cannot find the host element.');

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
                this._szCtl = new Size(host.offsetWidth, host.offsetHeight);
                var hr = this._handleResize.bind(this);
                this.addEventListener(window, 'resize', hr);
            }

            // fire events for got/lost focus
            // use timeOuts since Chrome and FF sometimes move the focus to the body
            // before moving it to the new focused element
            this.addEventListener(host, 'focus', () => {
                setTimeout(() => { // TFS 146807
                    this._updateFocusState();
                });
            }, true);
            this.addEventListener(host, 'blur', () => {
                setTimeout(() => { // TFS 146807
                    this._updateFocusState();
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
                    var b = document.body,
                        ts = this._handleTouchStart,
                        te = this._handleTouchEnd;
                    if ('ontouchstart' in window) { // Chrome, FireFox, Safari
                        b.addEventListener('touchstart', ts);
                        b.addEventListener('touchend', te);
                        b.addEventListener('touchcancel', te);
                        b.addEventListener('touchleave', te);
                    } else if ('onpointerdown' in window) { // IE
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
        getTemplate(): string {
            for (var p = Object.getPrototypeOf(this); p; p = Object.getPrototypeOf(p)) {
                var tpl = p.constructor.controlTemplate;
                if (tpl) {
                    return tpl;
                }
            }
            return null;
        }
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
        applyTemplate(classNames: string, template: string, parts: Object, namePart?: string): HTMLElement {
            var host = this._e;

            // apply standard classes to host element
            if (classNames) {
                addClass(host, classNames);
            }

            // convert string into HTML template and append to host
            var tpl = null;
            if (template) {
                tpl = createElement(template, host);
            }

            // copy key attributes from the host element (name, validation) to inner input 
            // element
            // NOTE 1: do this only if there is a single input element in the template
            // NOTE 2: do not copy 'type' since it causes issues in Chrome: TFS 84900, 84901
            var inputs = host.querySelectorAll('input'),
                input = <HTMLInputElement>(inputs.length == 1 ? inputs[0] : null),
                atts = host.attributes;
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
                var label = <HTMLLabelElement>document.querySelector('label[for="' + host.id + '"]'),
                    baseId = host.id + '_input',
                    newId = baseId;
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
                var evtChange = document.createEvent('HTMLEvents'),
                    orgVal = input.value;
                evtChange.initEvent('change', true, false);
                this.addEventListener(input, 'input', () => {
                    orgVal = input.value;
                }, true);
                this.gotFocus.addHandler(() => {
                    orgVal = input.value;
                });
                this.lostFocus.addHandler(() => {
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
                        var key = 'name',
                            att = host.attributes[key];
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
        }

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
        dispose() {

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
                    var evt = <Event>this[prop[2].toLowerCase() + prop.substr(3)];
                    if (evt instanceof Event) {
                        evt.removeAllHandlers();
                    }
                }
            }

            // if the control has a collectionView property, remove handlers to stop receiving notifications
            // REVIEW: perhaps optimize by caching the CollectionView properties?
            var cv = <collections.CollectionView>this['collectionView'];
            if (cv instanceof collections.CollectionView) {
                for (var prop in cv) {
                    var evt = <Event>cv[prop];
                    if (evt instanceof Event) {
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
        }
        /**
         * Gets the control that is hosted in a given DOM element.
         *
         * @param element The DOM element that is hosting the control, or a selector for the host element (e.g. '#theCtrl').
         */
        static getControl(element: any): Control {
            var e = getElement(element);
            return e ? asType(e[Control._DATA_KEY], Control, true) : null;
        }
        /**
         * Gets the DOM element that is hosting the control.
         */
        get hostElement(): HTMLElement {
            return this._e;
        }
        /**
         * Sets the focus to this control.
         */
        focus() {
            if (this._e) { // test for disposed controls
                this._e.focus();
            }
        }
        /**
         * Checks whether this control contains the focused element.
         */
        containsFocus(): boolean {

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
            var ae = getActiveElement(),
                popup = <any>Control.getControl(closest(ae, '.wj-control.wj-popup'));
            if (popup && popup.owner && contains(this._e, popup.owner)) {
                return true;
            }

            // check for actual HTML containment
            return contains(this._e, ae);
        }
        /**
         * Invalidates the control causing an asynchronous refresh.
         *
         * @param fullUpdate Whether to update the control layout as well as the content.
         */
        invalidate(fullUpdate = true) {
            this._fullUpdate = this._fullUpdate || fullUpdate;
            if (this._toInv) {
                clearTimeout(this._toInv);
                this._toInv = null;
            }
            if (!this.isUpdating) {
                this._toInv = setTimeout(() => {
                    this.refresh(this._fullUpdate);
                }, Control._REFRESH_INTERVAL);
            }
        }
        /**
         * Refreshes the control.
         *
         * @param fullUpdate Whether to update the control layout as well as the content.
         */
        refresh(fullUpdate = true) {
            if (!this.isUpdating && this._toInv) {
                clearTimeout(this._toInv);
                this._toInv = null;
                this._fullUpdate = false;
            }
            // derived classes should override this...
        }
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
        static invalidateAll(e? : HTMLElement) {
            if (!e) e = document.body;
            if (e.children) {
                for (var i = 0; i < e.children.length; i++) {
                    Control.invalidateAll(<HTMLElement>e.children[i]);
                }
            }
            var ctl = Control.getControl(e);
            if (ctl) {
                ctl.invalidate();
            }
        }
        /**
         * Refreshes all Wijmo controls contained in an HTML element.
         *
         * This method is similar to @see:invalidateAll, except the controls
         * are updated immediately rather than after an interval.
         *
         * @param e Container element. If set to null, all Wijmo controls
         * on the page will be invalidated.
         */
        static refreshAll(e?: HTMLElement) {
            if (!e) e = document.body;
            if (e.children) {
                for (var i = 0; i < e.children.length; i++) {
                    Control.refreshAll(<HTMLElement>e.children[i]);
                }
            }
            var ctl = Control.getControl(e);
            if (ctl) {
                ctl.refresh();
            }
        }
        /**
         * Disposes of all Wijmo controls contained in an HTML element.
         *
         * @param e Container element.
         */
        static disposeAll(e?: HTMLElement) {
            var ctl = Control.getControl(e);
            if (ctl) {
                ctl.dispose();
            } else if (e.children) {
                for (var i = 0; i < e.children.length; i++) {
                    Control.disposeAll(<HTMLElement>e.children[i]);
                }
            }
        }
        /**
         * Suspends notifications until the next call to @see:endUpdate.
         */
        beginUpdate() {
            this._updating++;
        }
        /**
         * Resumes notifications suspended by calls to @see:beginUpdate.
         */
        endUpdate() {
            this._updating--;
            if (this._updating <= 0) {
                this.invalidate();
            }
        }
        /**
         * Gets a value that indicates whether the control is currently being updated.
         */
        get isUpdating(): boolean {
            return this._updating > 0;
        }
        /**
         * Executes a function within a @see:beginUpdate/@see:endUpdate block.
         *
         * The control will not be updated until the function has been executed.
         * This method ensures @see:endUpdate is called even if the function throws
         * an exception.
         *
         * @param fn Function to be executed. 
         */
        deferUpdate(fn: Function) {
            try {
                this.beginUpdate();
                fn();
            } finally {
                this.endUpdate();
            }
        }
        /**
         * Gets a value that indicates whether the control is currently handling a touch event.
         */
        get isTouching(): boolean {
            return Control._touching;
        }
        /**
         * Gets or sets a value that determines whether the control is disabled.
         *
         * Disabled controls cannot get mouse or keyboard events.
         */
        get isDisabled(): boolean {
            return this._e && this._e.getAttribute('disabled') != null;
        }
        set isDisabled(value: boolean) {
            value = asBoolean(value, true);
            if (value != this.isDisabled) {
                enable(this._e, !value);
            }
        }
        /*
         * Deprecated: use 'isDisabled' instead to avoid confusion with 'disabled' HTML attribute.
         */
        get disabled(): boolean {
            _deprecated('disabled', 'isDisabled');
            return this.isDisabled;
        }
        set disabled(value: boolean) {
            _deprecated('disabled', 'isDisabled');
            this.isDisabled = value;
        }
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
        initialize(options: any) {
            if (options) {
                this.beginUpdate();
                copy(this, options);
                this.endUpdate();
            }
        }
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
        addEventListener(target: EventTarget, type: string, fn: any, capture = false) {
            if (target) {
                target.addEventListener(type, fn, capture);
                if (this._listeners == null) {
                    this._listeners = [];
                }
                this._listeners.push({ target: target, type: type, fn: fn, capture: capture });
            }
        }
        /**
         * Removes one or more event listeners attached to elements owned by this @see:Control.
         *
         * @param target Target element for the event. If null, removes listeners attached to all targets.
         * @param type String that specifies the event. If null, removes listeners attached to all events.
         * @param capture Whether the listener is capturing. If null, removes capturing and non-capturing listeners.
         * @return The number of listeners removed.
         */
        removeEventListener(target?: EventTarget, type?: string, capture?: boolean): number {
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
        }
        /**
         * Occurs when the control gets the focus.
         */
        gotFocus = new Event();
        /**
         * Raises the @see:gotFocus event.
         */
        onGotFocus(e?: EventArgs) {
            this.gotFocus.raise(this, e);
        }
        /**
         * Occurs when the control loses the focus.
         */
        lostFocus = new Event();
        /**
         * Raises the @see:lostFocus event.
         */
        onLostFocus(e?: EventArgs) {
            this.lostFocus.raise(this, e);
        }

        // ** implementation

        // invalidates the control when its size changes
        protected _handleResize() {
            if (this._e.parentElement) {
                var sz = new Size(this._e.offsetWidth, this._e.offsetHeight);
                if (!sz.equals(this._szCtl)) {
                    this._szCtl = sz;
                    this.invalidate();
                }
            }
        }

        // update focus state and raise got/lost focus events
        protected _updateFocusState() {

            // use a timeOut since Chrome and FF sometimes move the focus to the body
            // before moving it to the new focused element (CellTemplateIntro sample)
            setTimeout(() => {

                // update focus for this control and all ancestors as well
                for (var e = this._e; e; e = e.parentElement) {
                    var ctl = Control.getControl(e);
                    if (ctl) {
                        var focus = ctl.containsFocus();
                        if (focus != ctl._focus) {
                            ctl._focus = focus;
                            if (focus) {
                                ctl.onGotFocus();
                            } else {
                                ctl.onLostFocus();
                            }
                            ctl._updateState();
                        }
                    }
                }
            }); // no interval: TFS 146949
        }

        // update state attributes for this control (focused, empty, invalid)
        protected _updateState() {
            var host = this.hostElement,
                input = <HTMLInputElement>this.hostElement.querySelector('input');
            toggleClass(host, 'wj-state-focused', this.containsFocus());
            if (input instanceof HTMLInputElement) {
                toggleClass(host, 'wj-state-empty', input.value.length == 0);
                toggleClass(host, 'wj-state-readonly', input.readOnly);
                var vm = input.validationMessage; // may be null in IE9 (TFS 204492)
                toggleClass(host, 'wj-state-invalid', vm && vm.length > 0);
            }
        }

        // keep track of touch events
        protected _handleTouchStart(e) {
            if (e.pointerType == null || e.pointerType == 'touch') {
                Control._touching = true;
                //console.log('touching = true');
            }
        }
        protected _handleTouchEnd(e) {
            if (e.pointerType == null || e.pointerType == 'touch') {
                setTimeout(function () {
                    Control._touching = false;
                    //console.log('touching = false');
                }, 400); // 300ms click event delay on IOS, plus some safety
            }
        }

        // suppress mouse and keyboard events if the control is disabled
        // (pointer-events requires IE11, doesn't prevent wheel at all)
        private _handleDisabled(e: any) {
            if (this.isDisabled) {
                e.preventDefault();
                e.stopPropagation();
                e.stopImmediatePropagation();
            }
        }

        // replaces an element with a div element, copying the child elements 
        // and the 'id' and 'style' attributes from the original element
        private _replaceWithDiv(element: HTMLElement) {

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
        }
    }
}
module wijmo {
    'use strict';

    /**
     * Specifies the type of aggregate to calculate over a group of values.
     */
    export enum Aggregate {
        /**
         * No aggregate.
         */
        None,
        /**
         * Returns the sum of the numeric values in the group.
         */
        Sum,
        /**
         * Returns the count of non-null values in the group.
         */
        Cnt,
        /**
         * Returns the average value of the numeric values in the group.
         */
        Avg,
        /**
         * Returns the maximum value in the group.
         */
        Max,
        /**
         * Returns the minimum value in the group.
         */
        Min,
        /**
         * Returns the difference between the maximum and minimum numeric values in the group.
         */
        Rng,
        /**
         * Returns the sample standard deviation of the numeric values in the group 
         * (uses the formula based on n-1).
         */
        Std,
        /**
         * Returns the sample variance of the numeric values in the group 
         * (uses the formula based on n-1).
         */
        Var,
        /**
         * Returns the population standard deviation of the values in the group 
         * (uses the formula based on n).
         */
        StdPop,
        /**
         * Returns the population variance of the values in the group 
         * (uses the formula based on n).
         */
        VarPop,
        /**
         * Returns the count of all values in the group (including nulls).
         */
        CntAll
    }
    /**
     * Calculates an aggregate value from the values in an array.
     *
     * @param aggType Type of aggregate to calculate.
     * @param items Array with the items to aggregate.
     * @param binding Name of the property to aggregate on (in case the items are not simple values).
     */
    export function getAggregate(aggType: Aggregate, items: any[], binding?: string) {
        var cnt = 0,
            cntn = 0,
            sum = 0,
            sum2 = 0,
            min = null,
            max = null,
            bnd = binding ? new Binding(binding) : null;

        // special case: overall count (including nulls)
        aggType = asEnum(aggType, Aggregate);
        if (aggType == Aggregate.CntAll) {
            return items.length;
        }

        // calculate aggregate
        for (var i = 0; i < items.length; i++) {

            // get item/value
            var val = items[i];
            if (bnd) {
                val = bnd.getValue(val);
                //assert(!isUndefined(val), 'item does not define property "' + binding + '".');
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
                if (isNumber(val) && !isNaN(val)) {
                    cntn++;
                    sum += val;
                    sum2 += val * val;
                } else if (isBoolean(val)) {
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
        switch (aggType)
        {
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
}
/**
 * Defines interfaces and classes related to data, including the @see:ICollectionView 
 * interface, @see:CollectionView and @see:ObservableArray classes.
 */     
module wijmo.collections {
    'use strict';

    /**
     * Notifies listeners of dynamic changes, such as when items get added and 
     * removed or when the collection is sorted, filtered, or grouped.
     */
    export interface INotifyCollectionChanged {
        /**
         * Occurs when the collection changes.
         */
        collectionChanged: Event;
    }
    /**
     * Describes the action that caused the @see:INotifyCollectionChanged.collectionChanged
     * event to fire.
     */
    export enum NotifyCollectionChangedAction {
        /** An item was added to the collection. */
        Add,
        /** An item was removed from the collection. */
        Remove,
        /** An item was changed or replaced. */
        Change,
        /** 
         * Several items changed simultaneously 
         * (for example, the collection was sorted, filtered, or grouped). 
         */
        Reset
    }
    /**
     * Provides data for the @see:INotifyCollectionChanged.collectionChanged event.
     */
    export class NotifyCollectionChangedEventArgs extends EventArgs {

        /**
         * Provides a reset notification.
         */
        static reset = new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset);
        /**
         * Gets the action that caused the event to fire.
         */
        action: NotifyCollectionChangedAction;
        /**
         * Gets the item that was added, removed, or changed.
         */
        item: any;
        /**
         * Gets the index at which the change occurred.
         */
        index: number;
        /**
         * Initializes a new instance of the @see:NotifyCollectionChangedEventArgs class.
         *
         * @param action Type of action that caused the event to fire.
         * @param item Item that was added or changed.
         * @param index Index of the item.
         */
        constructor(action = NotifyCollectionChangedAction.Reset, item = null, index = -1) {
            super();
            this.action = action;
            this.item = item;
            this.index = index;
        }
    }

    /**
     * Represents a method that takes an item of any type and returns a 
     * boolean that indicates whether the object meets a set of criteria.
     */
    export interface IPredicate {
        (item: any): boolean
    }

    /**
    * Represents the method that compares two objects.
    */
    export interface IComparer {
        (x: any, y: any): number;
    }

    /**
     * Describes a sorting criterion.
     */
    export class SortDescription {
        _bnd: Binding;
        _asc: boolean;

        /**
         * Initializes a new instance of the @see:SortDescription class.
         *
         * @param property Name of the property to sort on.
         * @param ascending Whether to sort in ascending order.
         */
        constructor(property: string, ascending: boolean) {
            this._bnd = new Binding(property);
            this._asc = ascending;
        }
        /**
         * Gets the name of the property used to sort.
         */
        get property(): string {
            return this._bnd.path;
        }
        /**
         * Gets a value that determines whether to sort the values in ascending order.
         */
        get ascending(): boolean {
            return this._asc;
        }
    }

    /**
     * Enables collections to have the functionalities of current record management, 
     * custom sorting, filtering, and grouping.
     *
     * This is a JavaScript version of the <b>ICollectionView</b> interface used in 
     * Microsoft's XAML platform. It provides a consistent, powerful, and  MVVM-friendly 
     * way to bind data to UI elements.
     *
     * Wijmo includes several classes that implement @see:ICollectionView. The most 
     * common is @see:CollectionView, which works based on regular JavsScript 
     * arrays.
     */
    export interface ICollectionView extends INotifyCollectionChanged, IQueryInterface {

        /**
         * Gets a value that indicates whether this view supports filtering via the 
         * @see:filter property.
         */
        canFilter: boolean;
        /**
         * Gets a value that indicates whether this view supports grouping via the 
         * @see:groupDescriptions property.
         */
        canGroup: boolean;
        /**
         * Gets a value that indicates whether this view supports sorting via the 
         * @see:sortDescriptions property.
         */
        canSort: boolean;
        /**
         * Gets the current item in the view.
         */
        currentItem: any;
        /**
         * Gets the ordinal position of the current item in the view.
         */
        currentPosition: number;
        /**
         * Gets or sets a callback used to determine if an item is suitable for 
         * inclusion in the view.
         *
         * NOTE: If the filter function needs a scope (i.e. a meaningful 'this'
         * value), then remember to set the filter using the 'bind' function to
         * specify the 'this' object. For example:
         * <pre>
         *   collectionView.filter = this._filter.bind(this);
         * </pre>
         */
        filter: IPredicate;
        /**
         * Gets a collection of @see:GroupDescription objects that describe how the 
         * items in the collection are grouped in the view.
         */
        groupDescriptions: ObservableArray;
        /**
         * Gets the top-level groups.
         */
        groups: any[];
        /**
         * Gets a value that indicates whether this view contains no items.
         */
        isEmpty: boolean;
        /**
         * Gets a collection of @see:SortDescription objects that describe how the items 
         * in the collection are sorted in the view.
         */
        sortDescriptions: ObservableArray;
        /**
         * Gets or sets the collection object from which to create this view.
         */
        sourceCollection: any;
        /**
         * Returns a value that indicates whether a given item belongs to this view.
         *
         * @param item The item to locate in the collection.
         */
        contains(item: any): boolean;
        /**
         * Sets the specified item to be the current item in the view.
         *
         * @param item The item to set as the @see:currentItem.
         */
        moveCurrentTo(item: any): boolean;
        /**
         * Sets the first item in the view as the current item.
         */
        moveCurrentToFirst(): boolean;
        /**
         * Sets the last item in the view as the current item.
         */
        moveCurrentToLast(): boolean;
        /**
         * Sets the item after the current item in the view as the current item.
         */
        moveCurrentToNext(): boolean;
        /**
         * Sets the item at the specified index in the view as the current item.
         *
         * @param index The index of the item to set as the @see:currentItem.
         */
        moveCurrentToPosition(index: number): boolean;
        /**
         * Sets the item before the current item in the view as the current item.
         */
        moveCurrentToPrevious();
        /**
         * Re-creates the view using the current sort, filter, and group parameters.
         */
        refresh();
        /**
         * Occurs after the current item changes.
         */
        currentChanged: Event;
        /**
         * Occurs before the current item changes.
         */
        currentChanging: Event;

        // since we don't have IDisposable/using:

        /**
         * Suspends refreshes until the next call to @see:endUpdate.
         */
        beginUpdate();
        /**
         * Resumes refreshes suspended by a call to @see:beginUpdate.
         */
        endUpdate();
        /**
         * Executes a function within a beginUpdate/endUpdate block.
         *
         * The collection will not be refreshed until the function has been executed.
         * This method ensures endUpdate is called even if the function throws.
         *
         * @param fn Function to be executed within the beginUpdate/endUpdate block.
         */
        deferUpdate(fn: Function);

        // since we don't have IEnumerable:

        /**
         * Gets the filtered, sorted, grouped items in the view.
         */
        items: any[];
    }

    /**
     * Defines methods and properties that extend @see:ICollectionView to provide 
     * editing capabilities.
     */
    export interface IEditableCollectionView extends ICollectionView {
        /**
         * Gets a value that indicates whether a new item can be added to the collection.
         */
        canAddNew: boolean;
        /**
         * Gets a value that indicates whether the collection view can discard pending changes 
         * and restore the original values of an edited object.
         */
        canCancelEdit: boolean;
        /**
`        * Gets a value that indicates whether items can be removed from the collection.
         */
        canRemove: boolean;
        /**
`        * Gets the item that is being added during the current add transaction.
         */
        currentAddItem: any;
        /**
`        * Gets the item that is being edited during the current edit transaction.
         */
        currentEditItem: any;
        /**
`        * Gets a value that indicates whether an add transaction is in progress.
         */
        isAddingNew: boolean;
        /**
`        * Gets a value that indicates whether an edit transaction is in progress.
         */
        isEditingItem: boolean;
        /**
`        * Adds a new item to the collection.
         *
         * @return The item that was added to the collection.
         */
        addNew(): any;
        /**
         * Ends the current edit transaction and, if possible, 
         * restores the original value to the item.
         */
        cancelEdit();
        /**
         * Ends the current add transaction and discards the pending new item.
         */
        cancelNew();
        /**
         * Ends the current edit transaction and saves the pending changes.
         */
        commitEdit();
        /**
         * Ends the current add transaction and saves the pending new item.
         */
        commitNew();
        /**
         * Begins an edit transaction of the specified item.
         *
         * @param item Item to edit.
         */
        editItem(item: any);
        /**
         * Removes the specified item from the collection.
         *
         * @param item Item to remove from the collection.
         */
        remove(item: any);
        /**
         * Removes the item at the specified index from the collection.
         *
         * @param index Index of the item to remove from the collection.
         */
        removeAt(index: number);
    }

    /**
     * Defines methods and properties that extend @see:ICollectionView to provide 
     * paging capabilities.
     */
    export interface IPagedCollectionView extends ICollectionView {
        /**
`        * Gets a value that indicates whether the @see:pageIndex value can change.
         */
        canChangePage: boolean;
        /**
`        * Gets a value that indicates whether the index is changing.
         */
        isPageChanging: boolean;
        /**
`        * Gets the number of items in the view taking paging into account.
         *
         * To get the total number of items, use the @see:totalItemCount property.
         *
         * Notice that this is different from the .NET <b>IPagedCollectionView</b>,
         * where <b>itemCount</b> and <b>totalItemCount</b> both return the count
         * before paging is applied.
         */
        itemCount: number;
        /**
`        * Gets the zero-based index of the current page.
         */
        pageIndex: number;
        /**
`        * Gets or sets the number of items to display on a page.
         */
        pageSize: number;
        /**
`        * Gets the total number of items in the view before paging is applied.
         *
         * To get the number of items in the current view not taking paging into 
         * account, use the @see:itemCount property.
         *
         * Notice that this is different from the .NET <b>IPagedCollectionView</b>,
         * where <b>itemCount</b> and <b>totalItemCount</b> both return the count
         * before paging is applied.
         */
        totalItemCount: number;
        /**
         * Sets the first page as the current page.
         */
        moveToFirstPage(): boolean;
        /**
         * Sets the last page as the current page.
         */
        moveToLastPage(): boolean;
        /**
         * Moves to the page after the current page.
         */
        moveToNextPage(): boolean;
        /**
         * Moves to the page at the specified index.
         *
         * @param index Index of the page to move to.
         */
        moveToPage(index: number): boolean;
        /**
         * Moves to the page before the current page.
         */
        moveToPreviousPage(): boolean;
        /**
        * Occurs after the page index changes.
        */
        pageChanged: Event;
        /**
         * Occurs before the page index changes.
         */
        pageChanging: Event;
    }

    /**
     * Provides data for the @see:IPagedCollectionView.pageChanging event
     */
    export class PageChangingEventArgs extends CancelEventArgs
    {
        /**
         * Gets the index of the page that is about to become current.
         */
        newPageIndex: number;

        /**
         * Initializes a new instance of the @see:PageChangingEventArgs class.
         *
         * @param newIndex Index of the page that is about to become current.
         */
        constructor(newIndex: number) {
            super();
            this.newPageIndex = newIndex;
        }
    }

    /**
     * Represents a base class for types defining grouping conditions. 
     *
     * The concrete class which is commonly used for this purpose is 
     * @see:PropertyGroupDescription.
     */
    export class GroupDescription {

        /**
         * Returns the group name for the given item.
         *
         * @param item The item to get group name for.
         * @param level The zero-based group level index.
         * @return The name of the group the item belongs to.
         */
        public groupNameFromItem(item: any, level: number): any {
            return '';
        }
        /**
         * Returns a value that indicates whether the group name and the item name
         * match (which implies that the item belongs to the group).
         *
         * @param groupName The name of the group.
         * @param itemName The name of the item.
         * @return True if the names match; otherwise, false.
         */
        public namesMatch(groupName: any, itemName: any): boolean {
            return groupName === itemName;
        }
    }

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
    export class PropertyGroupDescription extends GroupDescription {
        _bnd: Binding;
        _converter: Function;

        /**
         * Initializes a new instance of the @see:PropertyGroupDescription class.
         *
         * @param property The name of the property that specifies
         * which group an item belongs to.
         * @param converter A callback function that takes an item and 
         * a property name and returns the group name. If not specified, 
         * the group name is the property value for the item.
         */
        constructor(property: string, converter?: Function) {
            super();
            this._bnd = new Binding(property);
            this._converter = converter;
        }
        /*
         * Gets the name of the property that is used to determine which 
         * group an item belongs to.
         */
        get propertyName(): string {
            return this._bnd.path;
        }
        /**
         * Returns the group name for the given item.
         *
         * @param item The item to get group name for.
         * @param level The zero-based group level index.
         * @return The name of the group the item belongs to.
         */
        public groupNameFromItem(item: any, level: number): any {
            return this._converter
                ? this._converter(item, this.propertyName)
                : this._bnd.getValue(item);
        }
        /**
         * Returns a value that indicates whether the group name and the item name
         * match (which implies that the item belongs to the group).
         *
         * @param groupName The name of the group.
         * @param itemName The name of the item.
         * @return True if the names match; otherwise, false.
         */
       public namesMatch(groupName: any, itemName: any): boolean {
            return groupName === itemName;
        }
    }
}
module wijmo.collections {
    'use strict';

    /**
     * Base class for Array classes with notifications.
     */
    export class ArrayBase {

        /**
         * Initializes a new instance of the @see:ArrayBase class.
         */
        constructor() {
            this.length = 0;
            Array.apply(this, arguments);
        }

        // keep TypeScript happy (these will never be called, we changed the prototype)
        pop(): any {
            return null;
        }
        push(...item: any[]): number {
            return 0;
        }
        splice(index: number, count: number, value?: any): any[] {
            return null;
        }
        slice(begin: number, end?: number): any[] {
            return null;
        }
        indexOf(searchElement: any, fromIndex?: number) {
            return -1;
        }
        sort(compareFn?: Function): any[]{
            return null;
        }
        length: number;
    }

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
    export class ObservableArray extends ArrayBase implements INotifyCollectionChanged {
        private _updating = 0;

        /**
         * Initializes a new instance of the @see:ObservableArray class.
         *
         * @param data Array containing items used to populate the @see:ObservableArray.
         */
        constructor(data? : any[]) {
            super();

            // initialize the array
            if (data) {
                data = asArray(data);
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
        push(...item: any[]): number {
            var length = this.length;
            for (var i = 0; item && i < item.length; i++) {
                length = super.push(item[i]);
                if (!this._updating) {
                    this._raiseCollectionChanged(NotifyCollectionChangedAction.Add, item[i], length - 1);
                }
            }
            return length;
        }
        /*
         * Removes the last item from the array.
         *
         * @return The item that was removed from the array.
         */
        pop(): any {
            var item = super.pop();
            this._raiseCollectionChanged(NotifyCollectionChangedAction.Remove, item, this.length);
            return item;
        }
        /**
         * Removes and/or adds items to the array.
         *
         * @param index Position where items will be added or removed.
         * @param count Number of items to remove from the array.
         * @param item Item to add to the array.
         * @return An array containing the removed elements.
         */
        splice(index: number, count: number, item?: any): any[] {
            var rv;
            if (count && item) { // add and remove items (argh)
                rv = super.splice(index, count, item);
                if (count == 1) {
                    this._raiseCollectionChanged(NotifyCollectionChangedAction.Change, item, index);
                } else {
                    this._raiseCollectionChanged();
                }
                return rv;
            } else if (item) { // add a value to the array
                rv = super.splice(index, 0, item);
                this._raiseCollectionChanged(NotifyCollectionChangedAction.Add, item, index);
                return rv;
            } else { // remove one or more items from the array
                rv = super.splice(index, count);
                if (count == 1) {
                    this._raiseCollectionChanged(NotifyCollectionChangedAction.Remove, rv[0], index);
                } else {
                    this._raiseCollectionChanged();
                }
                return rv;
            }
        }
        /**
         * Creates a shallow copy of a portion of an array.
         *
         * @param begin Position where the copy starts.
         * @param end Position where the copy ends.
         * @return A shallow copy of a portion of an array.
         */
        slice(begin?: number, end?: number): any[] {
            return super.slice(begin, end);
        }
        /**
         * Searches for an item in the array.
         *
         * @param searchElement Element to locate in the array.
         * @param fromIndex The index where the search should start.
         * @return The index of the item in the array, or -1 if the item was not found.
         */
        indexOf(searchElement: any, fromIndex?: number): number {
            return super.indexOf(searchElement, fromIndex);
        }
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
        sort(compareFn?: Function): any[] {
            var rv = super.sort(compareFn);
            this._raiseCollectionChanged();
            return rv;
        }
        /**
         * Inserts an item at a specific position in the array.
         *
         * @param index Position where the item will be added.
         * @param item Item to add to the array.
         */
        insert(index: number, item: any) {
            this.splice(index, 0, item);
        }
        /**
         * Removes an item from the array.
         *
         * @param item Item to remove.
         * @return True if the item was removed, false if it wasn't found in the array.
         */
        remove(item: any): boolean {
            var index = this.indexOf(item);
            if (index > -1) {
                this.removeAt(index);
                return true;
            }
            return false;
        }
        /**
         * Removes an item at a specific position in the array.
         *
         * @param index Position of the item to remove.
         */
        removeAt(index: number) {
            this.splice(index, 1);
        }
        /**
         * Assigns an item at a specific position in the array.
         *
         * @param index Position where the item will be assigned.
         * @param item Item to assign to the array.
         */
        setAt(index: number, item: any) {

            // make sure we have enough elements to set at the right index!
            if (index > this.length) {
                this.length = index;
            }

            // go ahead and splice now
            this.splice(index, 1, item);
        }
        /**
         * Removes all items from the array.
         */
        clear() {
            if (this.length !== 0) {
                this.splice(0, this.length); // safer than setting length = 0
                //this.length = 0; // fastest way to clear an array
                this._raiseCollectionChanged();
            }
        }
        /**
         * Suspends notifications until the next call to @see:endUpdate.
         */
        beginUpdate() {
            this._updating++;
        }
        /**
         * Resumes notifications suspended by a call to @see:beginUpdate.
         */
        endUpdate() {
            if (this._updating > 0) {
                this._updating--;
                if (this._updating == 0) {
                    this._raiseCollectionChanged();
                }
            }
        }
        /**
         * Gets a value that indicates whether notifications are currently suspended
         * (see @see:beginUpdate and @see:endUpdate).
         */
        get isUpdating() {
            return this._updating > 0;
        }
        /**
         * Executes a function within a @see:beginUpdate/@see:endUpdate block.
         *
         * The collection will not be refreshed until the function finishes. 
         * This method ensures @see:endUpdate is called even if the function throws
         * an exception.
         *
         * @param fn Function to be executed without updates. 
         */
        deferUpdate(fn: Function) {
            try {
                this.beginUpdate();
                fn();
            } finally {
                this.endUpdate();
            }
        }

        // ** IQueryInterface

        /**
         * Returns true if the caller queries for a supported interface.
         *
         * @param interfaceName Name of the interface to look for.
         * @return True if the caller queries for a supported interface.
         */
        implementsInterface(interfaceName: string): boolean {
            return interfaceName == 'INotifyCollectionChanged';
        }

        // ** INotifyCollectionChanged

        /**
         * Occurs when the collection changes.
         */
        collectionChanged = new Event();
        /**
         * Raises the @see:collectionChanged event.
         *
         * @param e Contains a description of the change.
         */
        onCollectionChanged(e = NotifyCollectionChangedEventArgs.reset) {
            if (!this.isUpdating) {
                this.collectionChanged.raise(this, e);
            }
        }

        // creates event args and calls onCollectionChanged
        private _raiseCollectionChanged(action = NotifyCollectionChangedAction.Reset, item?: any, index?: number) {
            if (!this.isUpdating) {
                var e = new NotifyCollectionChangedEventArgs(action, item, index);
                this.onCollectionChanged(e);
            }
        }
    }
}
module wijmo.collections {
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
    export class CollectionView implements IEditableCollectionView, IPagedCollectionView {
        _src: any[];
        _ncc: INotifyCollectionChanged;
        _view: any[];
        _pgView: any[];
        _groups: CollectionViewGroup[];
        _fullGroups: CollectionViewGroup[];
        _digest: string;
        _idx = -1;
        _filter: IPredicate;
        _srtDsc = new ObservableArray();
        _grpDesc = new ObservableArray();
        _newItem = null;
        _edtItem = null;
        _edtClone: any;
        _committing: boolean;
        _canceling: boolean;
        _pgSz = 0;
        _pgIdx = 0;
        _updating = 0;
        _itemCreator: Function;
        _stableSort: boolean;
        _canFilter = true;
        _canGroup = true;
        _canSort = true;
        _canAddNew = true;
        _canCancelEdit = true;
        _canRemove = true;
        _canChangePage = true;
        _trackChanges = false;
        _chgAdded = new ObservableArray();
        _chgRemoved = new ObservableArray();
        _chgEdited = new ObservableArray();
        _srtCvt: Function;
        _srtCmp: Function;

        /**
         * Initializes a new instance of the @see:CollectionView class.
         * 
         * @param sourceCollection Array that serves as a source for this 
         * @see:CollectionView.
         * @param options JavaScript object containing initialization data for the control.
         */
        constructor(sourceCollection?: any, options?: any) {

            // check that sortDescriptions contains SortDescriptions
            this._srtDsc.collectionChanged.addHandler(() => {
                var arr = this._srtDsc;
                for (var i = 0; i < arr.length; i++) {
                    assert(arr[i] instanceof SortDescription, 'sortDescriptions array must contain SortDescription objects.');
                }
                if (this.canSort) {
                    this.refresh();
                }
            });

            // check that groupDescriptions contains GroupDescriptions
            this._grpDesc.collectionChanged.addHandler(() => {
                var arr = this._grpDesc;
                for (var i = 0; i < arr.length; i++) {
                    assert(arr[i] instanceof GroupDescription, 'groupDescriptions array must contain GroupDescription objects.');
                }
                if (this.canGroup) {
                    this.refresh();
                }
            });

            // initialize the source collection
            this.sourceCollection = sourceCollection ? sourceCollection : new ObservableArray();

            // apply options
            if (options) {
                this.beginUpdate();
                copy(this, options);
                this.endUpdate();
            }
        }

        // method used in JSON-style initialization
        _copy(key: string, value: any): boolean {
            if (key == 'sortDescriptions') {
                this.sortDescriptions.clear();
                var arr = asArray(value);
                for (var i = 0; i < arr.length; i++) {
                    var val = arr[i];
                    if (isString(val)) {
                        val = new SortDescription(val, true);
                    }
                    this.sortDescriptions.push(val);
                }
                return true;
            }
            if (key == 'groupDescriptions') {
                this.groupDescriptions.clear();
                var arr = asArray(value);
                for (var i = 0; i < arr.length; i++) {
                    var val = arr[i];
                    if (isString(val)) {
                        val = new PropertyGroupDescription(val);
                    }
                    this.groupDescriptions.push(val);
                }
                return true;
            }
            return false;
        }
        
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
        get newItemCreator(): Function {
            return this._itemCreator;
        }
        set newItemCreator(value: Function) {
            this._itemCreator = asFunction(value);
        }
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
        get sortConverter(): Function {
            return this._srtCvt;
        }
        set sortConverter(value: Function) {
            if (value != this._srtCvt) {
                this._srtCvt = asFunction(value, true);
            }
        }
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
        get sortComparer(): Function {
            return this._srtCmp;
        }
        set sortComparer(value: Function) {
            if (value != this._srtCmp) {
                this._srtCmp = asFunction(value, true);
            }
        }
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
        get useStableSort(): boolean {
            return this._stableSort;
        }
        set useStableSort(value: boolean) {
            this._stableSort = asBoolean(value);
        }
        /**
         * Calculates an aggregate value for the items in this collection.
         *
         * @param aggType Type of aggregate to calculate.
         * @param binding Property to aggregate on.
         * @param currentPage Whether to include only items on the current page.
         * @return The aggregate value.
         */
        getAggregate(aggType: Aggregate, binding: string, currentPage?: boolean) {
            var items = currentPage ? this._pgView : this._view;
            return getAggregate(aggType, items, binding);
        }

        // ** IQueryInterface

        /**
         * Returns true if the caller queries for a supported interface.
         *
         * @param interfaceName Name of the interface to look for.
         */
        implementsInterface(interfaceName: string): boolean {
            switch (interfaceName) {
                case 'ICollectionView':
                case 'IEditableCollectionView':
                case 'IPagedCollectionView':
                case 'INotifyCollectionChanged':
                    return true;
            }
            return false;
        }

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
        get trackChanges(): boolean {
            return this._trackChanges;
        }
        set trackChanges(value: boolean) {
            this._trackChanges = asBoolean(value);
        }
        /** 
         * Gets an @see:ObservableArray containing the records that were added to
         * the collection since @see:trackChanges was enabled.
         */
        get itemsAdded(): ObservableArray {
            return this._chgAdded;
        }
        /** 
         * Gets an @see:ObservableArray containing the records that were removed from
         * the collection since @see:trackChanges was enabled.
         */
        get itemsRemoved(): ObservableArray {
            return this._chgRemoved;
        }
        /** 
         * Gets an @see:ObservableArray containing the records that were edited in
         * the collection since @see:trackChanges was enabled.
         */
        get itemsEdited(): ObservableArray {
            return this._chgEdited;
        }
        /**
         * Clears all changes by removing all items in the @see:itemsAdded, 
         * @see:itemsRemoved, and @see:itemsEdited collections.
         *
         * Call this method after committing changes to the server or 
         * after refreshing the data from the server.
         */
        clearChanges() {
            this._chgAdded.clear();
            this._chgRemoved.clear();
            this._chgEdited.clear();
        }

        // ** INotifyCollectionChanged

        /**
         * Occurs when the collection changes.
         */
        collectionChanged = new Event();
        /**
         * Raises the @see:collectionChanged event.
         *
         * @param e Contains a description of the change.
         */
        onCollectionChanged(e = NotifyCollectionChangedEventArgs.reset) {

            // track changes applied to items outside of editItem/commitEdit blocks (TFS 204805)
            if (!this._committing && !this._canceling &&
                e.action == NotifyCollectionChangedAction.Change &&
                e.item != this.currentEditItem) {
                this._trackItemChanged(e.item);
            }

            // raise the event as usual
            this.collectionChanged.raise(this, e);
        }

        // creates event args and calls onCollectionChanged
        private _raiseCollectionChanged(action = NotifyCollectionChangedAction.Reset, item?: any, index?: number) {
            //console.log('** collection changed: ' + NotifyCollectionChangedAction[action] + ' **');
            var e = new NotifyCollectionChangedEventArgs(action, item, index);
            this.onCollectionChanged(e);
        }

        /**
         * Occurs before the value of the @see:sourceCollection property changes.
         */
        sourceCollectionChanging = new Event();
        /**
         * Raises the @see:sourceCollectionChanging event.
         *
         * @param e @see:CancelEventArgs that contains the event data.
         */
        onSourceCollectionChanging(e: CancelEventArgs): boolean {
            this.sourceCollectionChanging.raise(this, e);
            return !e.cancel;
        }
        /**
         * Occurs after the value of the @see:sourceCollection property changes.
         */
        sourceCollectionChanged = new Event();
        /**
         * Raises the @see:sourceCollectionChanged event.
         */
        onSourceCollectionChanged(e?: EventArgs) {
            this.sourceCollectionChanged.raise(this, e);
        }

        // ** ICollectionView

        /**
         * Gets a value that indicates whether this view supports filtering via the 
         * @see:filter property.
         */
        get canFilter(): boolean {
            return this._canFilter;
        }
        set canFilter(value: boolean) {
            this._canFilter = asBoolean(value);
        }
        /**
         * Gets a value that indicates whether this view supports grouping via the 
         * @see:groupDescriptions property.
         */
        get canGroup(): boolean {
            return this._canGroup;
        }
        set canGroup(value: boolean) {
            this._canGroup = asBoolean(value);
        }
        /**
         * Gets a value that indicates whether this view supports sorting via the 
         * @see:sortDescriptions property.
         */
        get canSort(): boolean {
            return this._canSort;
        }
        set canSort(value: boolean) {
            this._canSort = asBoolean(value);
        }
        /**
         * Gets or sets the current item in the view.
         */
        get currentItem(): any {
            return this._pgView && this._idx > -1 && this._idx < this._pgView.length
                ? this._pgView[this._idx]
                : null;
        }
        set currentItem(value: any) {
            this.moveCurrentTo(value);
        }
        /**
         * Gets the ordinal position of the current item in the view.
         */
        get currentPosition(): number {
            return this._idx;
        }
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
        get filter(): IPredicate {
            return this._filter;
        }
        set filter(value: IPredicate) {
            if (this._filter != value) {
                this._filter = <IPredicate>asFunction(value);
                if (this.canFilter) {
                    this.refresh();
                }
            }
        }
        /**
         * Gets a collection of @see:GroupDescription objects that describe how the 
         * items in the collection are grouped in the view.
         */
        get groupDescriptions(): ObservableArray {
            return this._grpDesc;
        }
        /**
         * Gets an array of @see:CollectionViewGroup objects that represents the 
         * top-level groups.
         */
        get groups(): CollectionViewGroup[] {
            return this._groups;
        }
        /**
         * Gets a value that indicates whether this view contains no items.
         */
        get isEmpty(): boolean {
            return !this._pgView || !this._pgView.length;
        }
        /**
         * Gets a collection of @see:SortDescription objects that describe how the items 
         * in the collection are sorted in the view.
         */
        get sortDescriptions(): ObservableArray {
            return this._srtDsc;
        }
        /**
         * Gets or sets the underlying (unfiltered and unsorted) collection.
         */
        get sourceCollection(): any {
            return this._src;
        }
        set sourceCollection(sourceCollection: any) {
            if (sourceCollection != this._src) {

                // raise changing event
                if (!this.onSourceCollectionChanging(new CancelEventArgs())) {
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
                this._src = asArray(sourceCollection, false);
                this._ncc = <INotifyCollectionChanged>tryCast(this._src, 'INotifyCollectionChanged');
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
        }
        // handle notifications from the source collection
        private _sourceChanged(s: INotifyCollectionChanged, e: NotifyCollectionChangedEventArgs) {
            if (this._updating <= 0) {
                this.refresh(); // TODO: optimize
            }
        }
        /**
         * Returns a value indicating whether a given item belongs to this view.
         *
         * @param item Item to seek.
         */
        contains(item: any): boolean {
            return this._pgView.indexOf(item) > -1;
        }
        /**
         * Sets the specified item to be the current item in the view.
         *
         * @param item Item that will become current.
         */
        moveCurrentTo(item: any): boolean {
            return this.moveCurrentToPosition(this._pgView.indexOf(item));
        }
        /**
         * Sets the first item in the view as the current item.
         */
        moveCurrentToFirst(): boolean {
            return this.moveCurrentToPosition(0);
        }
        /**
         * Sets the last item in the view as the current item.
         */
        moveCurrentToLast(): boolean {
            return this.moveCurrentToPosition(this._pgView.length - 1);
        }
        /**
         * Sets the item before the current item in the view as the current item.
         */
        moveCurrentToPrevious(): boolean {
            return this._idx > 0 ? this.moveCurrentToPosition(this._idx - 1): false;
        }
        /**
         * Sets the item after the current item in the view as the current item.
         */
        moveCurrentToNext(): boolean {
            return this.moveCurrentToPosition(this._idx + 1);
        }
        /**
         * Sets the item at the specified index in the view as the current item.
         *
         * @param index Index of the item that will become current.
         */
        moveCurrentToPosition(index: number): boolean {
            if (index >= -1 && index < this._pgView.length) {
                var e = new CancelEventArgs();
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
        }
        /**
         * Re-creates the view using the current sort, filter, and group parameters.
         */
        refresh() {

            // not while updating, adding, or editing
            if (this._updating > 0 || this._newItem || this._edtItem) {
                return;
            }

            // perform the refresh
            this._performRefresh();

            // notify listeners
            this.onCollectionChanged();
        }

        // performs the refresh (without issuing notifications)
        _performRefresh() {

            // benchmark
            //var start = new Date();

            // save current item
            var current = this.currentItem;

            // create filtered view
            if (!this._src) {
                this._view = [];
            } else if (!this._filter || !this.canFilter) {
                this._view = (this._srtDsc.length > 0 && this.canSort)
                    ? this._src.slice(0) // clone source array
                    : this._src; // don't waste time cloning
            } else {
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
            this._pgIdx = clamp(this._pgIdx, 0, this.pageCount - 1);
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
        }

        // sorts an array in-place using the current sort descriptions
        _performSort(items: any[]) {

            if (this._stableSort) {

                // stable sort (nice, but 30-50% slower)
                // https://bugs.chromium.org/p/v8/issues/detail?id=90
                var arrIndexed = items.map(function (item, index) { return { item: item, index: index } }),
                    compare = this._compareItems();
                arrIndexed.sort(function (a, b) {
                    var r = compare(a.item, b.item);
                    return r == 0 ? a.index - b.index : r;
                });
                for (var i = 0; i < items.length; i++) {
                    items[i] = arrIndexed[i].item;
                }

            } else {

                // regular sort, not stable but very fast
                items.sort(this._compareItems());
            }

        }

        // this function is used in some of our samples, so
        // if we remove it or change its name some things will break...
        _compareItems() {
            var srtDsc = this._srtDsc,
                srtCvt = this._srtCvt,
                srtCmp = this._srtCmp,
                init = true,
                cmp = 0;
            return function (a, b) {
                for (var i = 0; i < srtDsc.length; i++) {

                    // get values
                    var sd = <SortDescription>srtDsc[i],
                        v1 = sd._bnd.getValue(a),
                        v2 = sd._bnd.getValue(b);

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
                    if (v1 !== v1) v1 = null;
                    if (v2 !== v2) v2 = null;

                    // ignore case when sorting unless the values are strings that only differ in case
                    // (to keep the sort consistent, TFS 131135)
                    if (typeof (v1) === 'string' && typeof (v2) === 'string') {
                        var lc1 = v1.toLowerCase(),
                            lc2 = v2.toLowerCase();
                        if (lc1 != lc2) {
                            v1 = lc1;
                            v2 = lc2;
                        }
                    }

                    // nulls always at the bottom (like excel)
                    if (v1 != null && v2 == null) return -1;
                    if (v1 == null && v2 != null) return +1;

                    // compare the values (at last!)
                    cmp = (v1 < v2) ? -1 : (v1 > v2) ? +1 : 0;
                    if (cmp != 0) {
                        return sd.ascending ? +cmp : -cmp;
                    }
                }
                return 0;
            }
        }

        // returns an array filtered using the current filter definition
        _performFilter(items: any[]): any[] {
            return this.canFilter && this._filter
                ? items.filter(this._filter, this)
                : items;
        }

        /**
         * Occurs after the current item changes.
         */
        currentChanged = new Event();
        /**
         * Raises the @see:currentChanged event.
         */
        onCurrentChanged(e = EventArgs.empty) {
            this.currentChanged.raise(this, e);
        }
        /**
         * Occurs before the current item changes.
         */
        currentChanging = new Event();
        /**
         * Raises the @see:currentChanging event.
         *
         * @param e @see:CancelEventArgs that contains the event data.
         */
        onCurrentChanging(e: CancelEventArgs): boolean {
            this.currentChanging.raise(this, e);
            return !e.cancel;
        }
        /**
         * Gets items in the view.
         */
        get items(): any[] {
            return this._pgView;
        }
        /**
         * Suspend refreshes until the next call to @see:endUpdate.
         */
        beginUpdate() {
            this._updating++;
        }
        /**
         * Resume refreshes suspended by a call to @see:beginUpdate.
         */
        endUpdate() {
            this._updating--;
            if (this._updating <= 0) {
                this.refresh();
            }
        }
        /**
         * Gets a value that indicates whether notifications are currently suspended
         * (see @see:beginUpdate and @see:endUpdate).
         */
        get isUpdating() {
            return this._updating > 0;
        }
        /**
         * Executes a function within a @see:beginUpdate/@see:endUpdate block.
         *
         * The collection will not be refreshed until the function finishes. 
         * This method ensures @see:endUpdate is called even if the function throws
         * an exception.
         *
         * @param fn Function to be executed without updates. 
         */
        deferUpdate(fn: Function) {
            try {
                this.beginUpdate();
                fn();
            } finally {
                this.endUpdate();
            }
        }

        // ** IEditableCollectionView

        /**
         * Gets a value that indicates whether a new item can be added to the collection.
         */
        get canAddNew(): boolean {
            return this._canAddNew;
        }
        set canAddNew(value: boolean) {
            this._canAddNew = asBoolean(value);
        }
        /**
         * Gets a value that indicates whether the collection view can discard pending changes 
         * and restore the original values of an edited object.
         */
        get canCancelEdit(): boolean {
            return this._canCancelEdit;
        }
        set canCancelEdit(value: boolean) {
            this._canCancelEdit = asBoolean(value);
        }
        /**
         * Gets a value that indicates whether items can be removed from the collection.
         */
        get canRemove(): boolean {
            return this._canRemove;
        }
        set canRemove(value: boolean) {
            this._canRemove = asBoolean(value);
        }
        /**
         * Gets the item that is being added during the current add transaction.
         */
        get currentAddItem(): any {
            return this._newItem;
        }
        /**
         * Gets the item that is being edited during the current edit transaction.
         */
        get currentEditItem(): any {
            return this._edtItem;
        }
        /**
         * Gets a value that indicates whether an add transaction is in progress.
         */
        get isAddingNew(): boolean {
            return this._newItem != null;
        }
        /**
         * Gets a value that indicates whether an edit transaction is in progress.
         */
        get isEditingItem(): boolean {
            return this._edtItem != null;
        }
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
        addNew(): any {

            // sanity
            if (arguments.length > 0) {
                assert(false, 'addNew does not take any parameters, it creates the new items.');
            }

            // commit pending changes
            this.commitEdit();
            this.commitNew();

            // honor canAddNew
            if (!this.canAddNew) {
                assert(false, 'cannot add items (canAddNew == false).');
                return null;
            }

            // create new item
            var item = null,
                src = this.sourceCollection;
            if (this.newItemCreator) {
                item = this.newItemCreator();
            } else if (src && src.length) {
                item = new src[0].constructor();
            } else {
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
                this._raiseCollectionChanged(NotifyCollectionChangedAction.Add, item, this._pgView.length - 1);

                // select the new item
                this.moveCurrentTo(item);
            }

            // done
            return this._newItem;
        }
        /**
         * Ends the current edit transaction and, if possible, 
         * restores the original value to the item.
         */
        cancelEdit() {
            var item = this._edtItem;
            if (item != null) {
                this._edtItem = null;

                // honor canCancelEdit
                if (!this.canCancelEdit) {
                    assert(false, 'cannot cancel edits (canCancelEdit == false).');
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
                this._raiseCollectionChanged(NotifyCollectionChangedAction.Change, item, index);
                this._canceling = false;
            }
        }
        /**
         * Ends the current add transaction and discards the pending new item.
         */
        cancelNew() {
            var item = this._newItem;
            if (item != null) {
                this.remove(item);
            }
        }
        /**
         * Ends the current edit transaction and saves the pending changes.
         */
        commitEdit() {
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
                    this._raiseCollectionChanged(NotifyCollectionChangedAction.Change, item, index);
                } else {
                    this._raiseCollectionChanged(); // full refresh
                }

                // done committing
                this._committing = false;
            }
        }
        /**
         * Ends the current add transaction and saves the pending new item.
         */
        commitNew() {
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
                    if (idx > -1) { // remove from changed if it's there
                        this._chgEdited.removeAt(idx);
                    }
                    if (this._chgAdded.indexOf(item) < 0) { // add to added if not there
                        this._chgAdded.push(item);
                    }
                }

                // notify (full refresh if the item moved)
                if (this._pgView.indexOf(item) == index && digest == this._digest) {
                    this._raiseCollectionChanged(NotifyCollectionChangedAction.Change, item, index);
                } else {
                    this._raiseCollectionChanged(); // full refresh
                }
            }
        }
        /**
         * Begins an edit transaction of the specified item.
         *
         * @param item Item to be edited.
         */
        editItem(item: any) {

            // commit pending changes if not already editing/adding this item
            if (item != this._edtItem && this.moveCurrentTo(item)) {
                this.commitEdit();
                this._edtItem = item;
                this._edtClone = {};
                this._extend(this._edtClone, this._edtItem);
            }
        }
        /**
         * Removes the specified item from the collection.
         *
         * @param item Item to be removed from the collection.
         */
        remove(item: any) {

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
                assert(false, 'cannot remove items (canRemove == false).');
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
                } else {
                    this._raiseCollectionChanged(NotifyCollectionChangedAction.Remove, item, index);
                }

                // raise currentChanged if needed
                if (this.currentItem !== current) {
                    this.onCurrentChanged();
                }
            }
        }
        /**
         * Removes the item at the specified index from the collection.
         *
         * @param index Index of the item to be removed from the collection.
         * The index is relative to the view, not to the source collection.
         */
        removeAt(index: number) {
            index = asInt(index);
            this.remove(this._pgView[index]);
        }

        // track changes applied to an item (not necessarily the current edit item)
        _trackItemChanged(item: any) {
            if (this._trackChanges) {
                var idx = this._chgEdited.indexOf(item);
                if (idx < 0 && this._chgAdded.indexOf(item) < 0) { // add item to changed list
                    this._chgEdited.push(item);
                } else if (idx > -1) { // item already on changed list, notify of change
                    var e = new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Change, item, idx);
                    this._chgEdited.onCollectionChanged(e);
                } else { // item on added list, notify of change
                    idx = this._chgAdded.indexOf(item);
                    if (idx > -1) {
                        var e = new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Change, item, idx);
                        this._chgAdded.onCollectionChanged(e);
                    }
                }
            }
        }

        // extends an object (shallow copy)
        _extend(dst: any, src: any) {
            for (var key in src) {
                dst[key] = src[key];
            }
        }

        // checks whether two objects have the same content
        _sameContent(dst: any, src: any) {
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
        }

        // checks whether two values are the same
        _sameValue(v1: any, v2: any) {
            return v1 == v2 || DateTime.equals(v1, v2);
        }

        // ** IPagedCollectionView

        /**
         * Gets a value that indicates whether the @see:pageIndex value can change.
         */
        get canChangePage(): boolean {
            return this._canChangePage;
        }
        set canChangePage(value: boolean) {
            this._canChangePage = asBoolean(value);
        }
        /**
         * Gets a value that indicates whether the page index is changing.
         */
        get isPageChanging(): boolean {
            return false;
        }
        /**
         * Gets the total number of items in the view taking paging into account.
         */
        get itemCount(): number {
            return this._pgView.length;
        }
        /**
         * Gets the zero-based index of the current page.
         */
        get pageIndex(): number {
            return this._pgIdx;
        }
        /**
         * Gets or sets the number of items to display on a page.
         */
        get pageSize(): number {
            return this._pgSz;
        }
        set pageSize(value: number) {
            if (value != this._pgSz) {
                this._pgSz = asInt(value);
                this.refresh();
            }
        }
        /**
         * Gets the total number of items in the view before paging is applied.
         */
        get totalItemCount(): number {
            return this._view.length;
        }
        /**
         * Gets the total number of pages.
         */
        get pageCount(): number {
            return this.pageSize ? Math.ceil(this.totalItemCount / this.pageSize) : 1;
        }
        /**
         * Sets the first page as the current page.
         *
         * @return True if the page index was changed successfully.
         */
        moveToFirstPage(): boolean {
            return this.moveToPage(0);
        }
        /**
         * Sets the last page as the current page.
         *
         * @return True if the page index was changed successfully.
         */
        moveToLastPage(): boolean {
            return this.moveToPage(this.pageCount - 1);
        }
        /**
         * Moves to the page after the current page.
         *
         * @return True if the page index was changed successfully.
         */
        moveToNextPage(): boolean {
            return this.moveToPage(this.pageIndex + 1);
        }
        /**
         * Moves to the page at the specified index.
         *
         * @param index Index of the page to move to.
         * @return True if the page index was changed successfully.
         */
        moveToPage(index: number): boolean {
            var newIndex = clamp(index, 0, this.pageCount - 1);
            if (newIndex != this._pgIdx) {

                // honor canChangePage
                if (!this.canChangePage) {
                    assert(false, 'cannot change pages (canChangePage == false).');
                }

                // raise pageChanging
                var e = new PageChangingEventArgs(newIndex);
                if (this.onPageChanging(e)) {

                    // change the page
                    this._pgIdx = newIndex;
                    this._pgView = this._getPageView();
                    this._idx = 0;

                    // raise pageChanged and collectionChanged, or refresh if grouping
                    if (!this.groupDescriptions || this.groupDescriptions.length == 0) {
                        this.onPageChanged();
                        this.onCollectionChanged();
                    } else {
                        this.refresh();
                    }
                }
            }
            return this._pgIdx == index;
        }
        /**
         * Moves to the page before the current page.
         *
         * @return True if the page index was changed successfully.
         */
        moveToPreviousPage(): boolean {
            return this.moveToPage(this.pageIndex - 1);
        }
        /**
        * Occurs after the page index changes.
        */
        pageChanged = new Event();
        /**
         * Raises the @see:pageChanged event.
         */
        onPageChanged(e = EventArgs.empty) {
            this.pageChanged.raise(this, e);
        }
        /**
         * Occurs before the page index changes.
         */
        pageChanging = new Event();
        /**
         * Raises the @see:pageChanging event.
         *
         * @param e @see:PageChangingEventArgs that contains the event data.
         */
        onPageChanging(e: PageChangingEventArgs): boolean {
            this.pageChanging.raise(this, e);
            return !e.cancel;
        }

        // gets the full group that corresponds to a paged group view
        _getFullGroup(g: CollectionViewGroup): CollectionViewGroup {

            // look for the group by level and name
            // this gets the full (unpaged) and updated group (TFS 109119)
            var fg = this._getGroupByPath(this._fullGroups, g.level, g._path);
            if (fg != null) {
                g = fg;
            }

            // return the group
            return g;
        }

        // gets a group from a collection by path
        _getGroupByPath(groups: CollectionViewGroup[], level: number, path: string) {
            for (var i = 0; i < groups.length; i++) {
                var g = groups[i];
                if (g.level == level && g._path == path) {
                    return g;
                }
                if (g.level < level && path.indexOf(g._path) == 0) { // TFS 139570
                    g = this._getGroupByPath(g.groups, level, path);
                    if (g != null) {
                        return g;
                    }
                }
            }
            return null;
        }

        // gets the list that corresponds to the current page
        _getPageView() {

            // not paging? return the whole view
            if (this.pageSize <= 0 || this._pgIdx < 0) {
                return this._view;
            }

            // slice the current page out of the view
            var start = this._pgSz * this._pgIdx,
                end = Math.min(start + this._pgSz, this._view.length);
            return this._view.slice(start, end);
        }

        // creates a grouped view of the current page
        _createGroups(items: any[]): CollectionViewGroup[] {

            // not grouping? return null
            if (!this._grpDesc || !this._grpDesc.length) {
                return null;
            }

            // build group tree
            var root: CollectionViewGroup[] = [],
                maps = {},
                map = null;
            for (var i = 0; i < items.length; i++) {

                // get the item
                var item = items[i],
                    groups = root,
                    levels = this._grpDesc.length;

                // add this item to the tree
                var path = '';
                for (var level = 0; level < levels; level++) {

                    // get the group name for this level
                    var gd = this._grpDesc[level],
                        name = gd.groupNameFromItem(item, level),
                        last = level == levels - 1;

                    // get the group map for this level (optimization)
                    map = maps[path];
                    if (!map && isPrimitive(name)) {
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
        }

        // gets a string digest of the current groups 
        // this is used to check whether changes require a full refresh
        private _getGroupsDigest(groups): string {
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
        }

        // gets an array that contains all the children for a list of groups
        // NOTE: use "push.apply" instead of "concat" for much better performance
        // NOTE2: use explicit loop for even better performance and to avoid stack overflows (TFS 15921)
        private _mergeGroupItems(groups: CollectionViewGroup[]): any[] {
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
        }

        // finds or creates a group
        private _getGroup(gd: GroupDescription, groups: CollectionViewGroup[], map: any, name: string, level: number, isBottomLevel: boolean): CollectionViewGroup {
            var g: CollectionViewGroup;

            // find existing group
            if (map && isPrimitive(name)) {
                g = map[name];
                if (g) {
                    return g;
                }
            } else {
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
        }
    }

    /**
     * Represents a group created by a @see:CollectionView object based on
     * its @see:CollectionView.groupDescriptions property.
     */
    export class CollectionViewGroup {
        _gd: GroupDescription;
        _name: string;
        _path: string;
        _level: number;
        _isBottomLevel: boolean;
        _groups: CollectionViewGroup[];
        _items: any[];

        /**
         * Initializes a new instance of the @see:CollectionViewGroup class.
         *
         * @param groupDescription @see:GroupDescription that owns the new group.
         * @param name Name of the new group.
         * @param level Level of the new group.
         * @param isBottomLevel Whether this group has any subgroups.
         */
        constructor(groupDescription: GroupDescription, name: string, level: number, isBottomLevel: boolean) {
            this._gd = groupDescription;
            this._name = name;
            this._level = level;
            this._isBottomLevel = isBottomLevel;
            this._groups = [];
            this._items = [];
        }
        /*
         * Gets the name of this group.
         */
        get name(): string {
            return this._name;
        }
        /*
         * Gets the level of this group.
         */
        get level(): number {
            return this._level;
        }
        /*
         * Gets a value that indicates whether this group has any subgroups.
         */
        get isBottomLevel(): boolean {
            return this._isBottomLevel;
        }
        /*
         * Gets an array containing the items included in this group (including all subgroups).
         */
        get items(): any[] {
            return this._items;
        }
        /*
         * Gets an array containing the this group's subgroups.
         */
        get groups(): CollectionViewGroup[] {
            return this._groups;
        }
        /*
         * Gets the @see:GroupDescription that owns this group.
         */
        get groupDescription(): GroupDescription {
            return this._gd;
        }
        /**
         * Calculates an aggregate value for the items in this group.
         *
         * @param aggType Type of aggregate to calculate.
         * @param binding Property to aggregate on.
         * @param view CollectionView that owns this group.
         * @return The aggregate value.
         */
        getAggregate(aggType: Aggregate, binding: string, view?: ICollectionView) {
            var cv = <CollectionView>tryCast(view, CollectionView),
                group = cv ? cv._getFullGroup(this): this;
            return getAggregate(aggType, group.items, binding);
        }
    }
}
module wijmo {
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
    export class Tooltip {

        // tooltip element
        private static _eTip: HTMLElement;

        // private stuff
        private _toShow: number;
        private _toHide: number;
        private _showAutoTipBnd = this._showAutoTip.bind(this);
        private _hideAutoTipBnd = this._hideAutoTip.bind(this);

        // property storage
        private _html = true;
        private _gap = 6;
        private _showAtMouse = false;
        private _showDelay = 500; // http://msdn.microsoft.com/en-us/library/windows/desktop/bb760404(v=vs.85).aspx
        private _hideDelay = 0; // do not hide
        private _tips: ElementContent[] = [];

        /**
         * Initializes a new instance of the @see:Tooltip class.
         *
         * @param options JavaScript object containing initialization data for the @see:Tooltip.
         */
        constructor(options?: any) {
            copy(this, options);
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
        setTooltip(element: any, content: string) {

            // get element and tooltip content
            element = getElement(element);
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
        }
        /**
         * Shows the tooltip with the specified content, next to the specified element.
         *
         * @param element Element, element ID, or control that the tooltip explains.
         * @param content Tooltip content or ID of the element that contains the tooltip content.
         * @param bounds Optional element that defines the bounds of the area that the tooltip 
         * targets. If not provided, the bounds of the element are used (as reported by the
         * <b>getBoundingClientRect</b> method).
         */
        show(element: any, content: string, bounds?: Rect) {

            // get element and tooltip content
            element = getElement(element);
            content = this._getContent(content);
            if (!bounds) {
                bounds = Rect.fromBoundingRect(element.getBoundingClientRect());
            }

            // create tooltip element if necessary
            var tip = Tooltip._eTip;
            if (!tip) {
                tip = Tooltip._eTip = document.createElement('div');
                addClass(tip, 'wj-tooltip');
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
                bounds = new Rect(
                    bounds.left - (tip.offsetWidth - bounds.width) / 2,
                    bounds.top - this.gap,
                    tip.offsetWidth,
                    bounds.height + 2 * this.gap);

                // show tooltip
                showPopup(tip, bounds, true);

                // hide when the mouse goes down
                document.addEventListener('mousedown', this._hideAutoTipBnd);
            }
        }
        /**
         * Hides the tooltip if it is currently visible.
         */
        hide() {
            if (Tooltip._eTip) {
                Tooltip._eTip.style.visibility = 'hidden';
                Tooltip._eTip.innerHTML = '';
            }
            document.removeEventListener('mousedown', this._hideAutoTipBnd);
        }
        /**
         * Gets whether the tooltip is currently visible. 
         */
        get isVisible(): boolean {
            return Tooltip._eTip && Tooltip._eTip.style.visibility != 'hidden';
        }
        /**
         * Gets or sets a value that determines whether the tooltip contents 
         * should be displayed as plain text or as HTML.
         */
        get isContentHtml(): boolean {
            return this._html;
        }
        set isContentHtml(value: boolean) {
            this._html = asBoolean(value);
        }
        /**
         * Gets or sets the distance between the tooltip and the target element.
         */
        get gap(): number {
            return this._gap;
        }
        set gap(value: number) {
            this._gap = asNumber(value);
        }
        /**
         * Gets or sets a value that determines whether the tooltip should be
         * positioned with respect to the mouse position rather than the
         * target element.
         */
        get showAtMouse(): boolean {
            return this._showAtMouse;
        }
        set showAtMouse(value: boolean) {
            this._showAtMouse = asBoolean(value);
        }
        /**
         * Gets or sets the delay, in milliseconds, before showing the tooltip after the 
         * mouse enters the target element.
         */
        get showDelay(): number {
            return this._showDelay;
        }
        set showDelay(value: number) {
            this._showDelay = asInt(value);
        }
        /**
         * Gets or sets the delay, in milliseconds, before hiding the tooltip after the 
         * mouse leaves the target element.
         */
        get hideDelay(): number {
            return this._hideDelay;
        }
        set hideDelay(value: number) {
            this._hideDelay = asInt(value);
        }

        /**
         * Occurs before the tooltip content is displayed.
         * 
         * The event handler may customize the tooltip content or suppress the 
         * tooltip display by changing the event parameters.
         */
        public popup = new Event();
        /**
         * Raises the @see:popup event.
         *
         * @param e @see:TooltipEventArgs that contains the event data.
         */
        onPopup(e: TooltipEventArgs) {
            if (this.popup) {
                this.popup.raise(this, e);
            }
        }

        // implementation

        // finds an element in the auto-tooltip list
        private _indexOf(e: Element): number {
            for (var i = 0; i < this._tips.length; i++) {
                if (this._tips[i].element == e) {
                    return i;
                }
            }
            return -1;
        }

        // add event listeners to show and hide tooltips for an element
        private _attach(e: HTMLElement) {
            e.addEventListener('mouseenter', this._showAutoTipBnd);
            e.addEventListener('mouseleave', this._hideAutoTipBnd);
            e.addEventListener('click', this._showAutoTipBnd);
        }

        // remove event listeners used to show and hide tooltips for an element
        private _detach(e: HTMLElement) {
            e.removeEventListener('mouseenter', this._showAutoTipBnd);
            e.removeEventListener('mouseleave', this._hideAutoTipBnd);
            e.removeEventListener('click', this._showAutoTipBnd);
        }

        // shows an automatic tooltip
        private _showAutoTip(e: MouseEvent) {
            var showDelay = e.type == 'mouseenter' ? this._showDelay : 0;
            this._clearTimeouts();
            this._toShow = setTimeout(() => {
                var i = this._indexOf(<Element>e.target);
                if (i > -1) {
                    var tip = this._tips[i],
                        bounds = this._showAtMouse ? new Rect(e.clientX, e.clientY, 0, 0) : null;
                    this.show(tip.element, tip.content, bounds);
                    if (this._hideDelay > 0) {
                        this._toHide = setTimeout(() => {
                            this.hide();
                        }, this._hideDelay);
                    }
                }
            }, showDelay);
        }

        // hides an automatic tooltip
        private _hideAutoTip() {
            this._clearTimeouts();
            this.hide();
        }

        // clears the timeouts used to show and hide automatic tooltips
        private _clearTimeouts() {
            if (this._toShow) {
                clearTimeout(this._toShow);
                this._toShow = null;
            }
            if (this._toHide) {
                clearTimeout(this._toHide);
                this._toHide = null;
            }
        }

        // gets content that may be a string or an element id
        private _getContent(content: string): string {
            content = asString(content);
            if (content && content[0] == '#') {
                var e = getElement(content);
                if (e) {
                    content = e.innerHTML;
                }
            }
            return content;
        }

        // assigns content to the tooltip element
        private _setContent(content: string) {
            var tip = Tooltip._eTip;
            if (tip) {
                if (this.isContentHtml) {
                    tip.innerHTML = content;
                } else {
                    tip.textContent = content;
                }
            }
        }
    }

    // helper class to hold element/tooltip information
    class ElementContent {
        element: HTMLElement;
        content: string;
    }

    /**
     * Provides arguments for the @see:Tooltip.popup event.
     */
    export class TooltipEventArgs extends CancelEventArgs {
        private _content: string;

        /**
         * Initializes a new instance of the @see:TooltipEventArgs class.
         *
         * @param content String to show in the tooltip.
         */
        constructor(content: string) {
            super();
            this._content = asString(content);
        }

        /**
         * Gets or sets the content to show in the tooltip.
         *
         * This parameter can be used while handling the @see:Tooltip.popup
         * event to modify the content of the tooltip.
         */
        get content(): string {
            return this._content;
        }
        set content(value: string) {
            this._content = asString(value);
        }
    }
} 
module wijmo {
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
    export class Color {

        // fields
        _r = 0;
        _g = 0;
        _b = 0;
        _a = 1;

        /**
         * Initializes a new @see:Color from a CSS color specification.
         *
         * @param color CSS color specification.
         */
        constructor(color: string) {
            if (color) {
                this._parse(color);
            }
        }

        /**
         * Gets or sets the red component of this @see:Color,
         * in a range from 0 to 255.
         */
        get r(): number {
            return this._r;
        }
        set r(value: number) {
            this._r = clamp(asNumber(value), 0, 255);
        }
        /**
         * Gets or sets the green component of this @see:Color,
         * in a range from 0 to 255.
         */
        get g(): number {
            return this._g;
        }
        set g(value: number) {
            this._g = clamp(asNumber(value), 0, 255);
        }
        /**
         * Gets or sets the blue component of this @see:Color,
         * in a range from 0 to 255.
         */
        get b(): number {
            return this._b;
        }
        set b(value: number) {
            this._b = clamp(asNumber(value), 0, 255);
        }
        /**
         * Gets or sets the alpha component of this @see:Color,
         * in a range from 0 to 1 (zero is transparent, one is solid).
         */
        get a(): number {
            return this._a;
        }
        set a(value: number) {
            this._a = clamp(asNumber(value), 0, 1);
        }
        /**
         * Returns true if a @see:Color has the same value as this @see:Color.
         *
         * @param clr @see:Color to compare to this @see:Color.
         */
        equals(clr: Color): boolean {
            return (clr instanceof Color) &&
                this.r == clr.r && this.g == clr.g && this.b == clr.b &&
                this.a == clr.a;
        }
        /**
         * Gets a string representation of this @see:Color.
         */
        toString(): string {
            var a = Math.round(this.a * 100);
            return a > 99
                ? '#' + ((1 << 24) + (this.r << 16) + (this.g << 8) + this.b).toString(16).slice(1)
                : 'rgba(' + this.r + ',' + this.g + ',' + this.b + ',' + (a / 100) +')';
        }
        /**
         * Creates a new @see:Color using the specified RGBA color channel values.
         *
         * @param r Value for the red channel, from 0 to 255.
         * @param g Value for the green channel, from 0 to 255.
         * @param b Value for the blue channel, from 0 to 255.
         * @param a Value for the alpha channel, from 0 to 1.
         */
        static fromRgba(r: number, g: number, b: number, a = 1): Color {
            var c = new Color(null);
            c.r = Math.round(clamp(asNumber(r), 0, 255));
            c.g = Math.round(clamp(asNumber(g), 0, 255));
            c.b = Math.round(clamp(asNumber(b), 0, 255));
            c.a = clamp(asNumber(a), 0, 1);
            return c;
        }
        /**
         * Creates a new @see:Color using the specified HSB values.
         *
         * @param h Hue value, from 0 to 1.
         * @param s Saturation value, from 0 to 1.
         * @param b Brightness value, from 0 to 1.
         * @param a Alpha value, from 0 to 1.
         */
        static fromHsb(h: number, s: number, b: number, a = 1): Color {
            var rgb = Color._hsbToRgb(clamp(asNumber(h), 0, 1), clamp(asNumber(s), 0, 1), clamp(asNumber(b), 0, 1));
            return Color.fromRgba(rgb[0], rgb[1], rgb[2], a);
        }
        /**
         * Creates a new @see:Color using the specified HSL values.
         *
         * @param h Hue value, from 0 to 1.
         * @param s Saturation value, from 0 to 1.
         * @param l Lightness value, from 0 to 1.
         * @param a Alpha value, from 0 to 1.
         */
        static fromHsl(h: number, s: number, l: number, a = 1): Color {
            var rgb = Color._hslToRgb(clamp(asNumber(h), 0, 1), clamp(asNumber(s), 0, 1), clamp(asNumber(l), 0, 1));
            return Color.fromRgba(rgb[0], rgb[1], rgb[2], a);
        }
        /**
         * Creates a new @see:Color from a CSS color string.
         *
         * @param value String containing a CSS color specification.
         * @return A new @see:Color, or null if the string cannot be parsed into a color.
         */
        static fromString(value: string): Color {
            var c = new Color(null);
            return c._parse(asString(value)) ? c : null;
        }
        /**
         * Gets an array with this color's HSB components.
         */
        getHsb(): number[] {
            return Color._rgbToHsb(this.r, this.g, this.b)
        }
        /**
         * Gets an array with this color's HSL components.
         */
        getHsl(): number[] {
            return Color._rgbToHsl(this.r, this.g, this.b)
        }
        /**
         * Creates a @see:Color by interpolating between two colors.
         *
         * @param c1 First color.
         * @param c2 Second color.
         * @param pct Value between zero and one that determines how close the
         * interpolation should be to the second color.
         */
        static interpolate(c1: Color, c2: Color, pct: number): Color {

            // sanity
            pct = clamp(asNumber(pct), 0, 1);

            // convert rgb to hsl
            var h1 = Color._rgbToHsl(c1.r, c1.g, c1.b),
                h2 = Color._rgbToHsl(c2.r, c2.g, c2.b);

            // interpolate
            var qct = 1 - pct,
                alpha = c1.a * qct + c2.a * pct,
                h3 = [
                    h1[0] * qct + h2[0] * pct,
                    h1[1] * qct + h2[1] * pct,
                    h1[2] * qct + h2[2] * pct
                ];

            // convert back to rgb
            var rgb = Color._hslToRgb(h3[0], h3[1], h3[2]);
            return Color.fromRgba(rgb[0], rgb[1], rgb[2], alpha);
        }
        /**
         * Gets the closest opaque color to a given color.
         *
         * @param c @see:Color to be converted to an opaque color 
         * (the color may also be specified as a string).
         * @param bkg Background color to use when removing the transparency
         * (defaults to white).
         */
        static toOpaque(c: any, bkg?: any): Color {

            // get color
            c = isString(c) ? Color.fromString(c) : asType(c, Color);

            // no alpha? no work
            if (c.a == 1) return c;

            // get background
            bkg = bkg == null ? Color.fromRgba(255, 255, 255, 1) :
                  isString(bkg) ? Color.fromString(bkg) :
                  <Color>asType(bkg, Color);

            // interpolate in RGB space
            var p = c.a, q = 1 - p;
            return Color.fromRgba(c.r * p + bkg.r * q, c.g * p + bkg.g * q, c.b * p + bkg.b * q);
        }

        // ** implementation

        // parses a color string into r/b/g/a
        _parse(c: string): boolean {

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
                if (cc == c) {                                  // same value? 
                    cc = window.getComputedStyle(e).color;      // then get computed style
                    if (!cc) {                                  // not yet? (Chrome/named colors)
                        document.body.appendChild(e);           // then add element to document
                        cc = window.getComputedStyle(e).color;  // and try again
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
                } else if (c.length == 7) {
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
                var op = c.indexOf('('),
                    ep = c.indexOf(')');
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
                var op = c.indexOf('('),
                    ep = c.indexOf(')');
                if (op > -1 && ep > -1) {
                    var p = c.substr(op + 1, ep - (op + 1)).split(',');
                    if (p.length > 2) {
                        var h = parseInt(p[0]) / 360,
                            s = parseInt(p[1]),
                            l = parseInt(p[2]);
                        if (p[1].indexOf('%') > -1) s /= 100;
                        if (p[2].indexOf('%') > -1) l /= 100;
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
        }
        /**
         * Converts an HSL color value to RGB.
         *
         * @param h The hue (between zero and one).
         * @param s The saturation (between zero and one).
         * @param l The lightness (between zero and one).
         * @return An array containing the R, G, and B values (between zero and 255).
         */
        static _hslToRgb(h: number, s: number, l: number): number[] {
            assert(h >= 0 && h <= 1 && s >= 0 && s <= 1 && l >= 0 && l <= 1, 'bad HSL values');
            var r: number, g: number, b: number;
            if (s == 0) {
                r = g = b = l; // achromatic
            } else {
                var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
                var p = 2 * l - q;
                r = Color._hue2rgb(p, q, h + 1 / 3);
                g = Color._hue2rgb(p, q, h);
                b = Color._hue2rgb(p, q, h - 1 / 3);
            }
            return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
        }
        static _hue2rgb(p: number, q: number, t: number): number {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1 / 6) return p + (q - p) * 6 * t;
            if (t < 1 / 2) return q;
            if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
            return p;
        }
        /**
         * Converts an RGB color value to HSL.
         *
         * @param r The value of the red channel (between zero and 255).
         * @param g The value of the green channel (between zero and 255).
         * @param b The value of the blue channel (between zero and 255).
         * @return An array containing the H, S, and L values (between zero and one).
         */
        static _rgbToHsl(r: number, g: number, b: number): number[] {
            assert(r >= 0 && r <= 255 && g >= 0 && g <= 255 && b >= 0 && b <= 255, 'bad RGB values');
            r /= 255, g /= 255, b /= 255;
            var max = Math.max(r, g, b),
                min = Math.min(r, g, b),
                h, s, l = (max + min) / 2;
            if (max == min) {
                h = s = 0;
            } else {
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
        }
        /**
         * Converts an RGB color value to HSB.
         *
         * @param r The value of the red channel (between zero and 255).
         * @param g The value of the green channel (between zero and 255).
         * @param b The value of the blue channel (between zero and 255).
         * @return An array containing the H, S, and B values (between zero and one).
         */
        static _rgbToHsb(r: number, g: number, b: number): number[]{
            assert(r >= 0 && r <= 255 && g >= 0 && g <= 255 && b >= 0 && b <= 255, 'bad RGB values');
            var hsl = Color._rgbToHsl(r, g, b);
            return Color._hslToHsb(hsl[0], hsl[1], hsl[2]);
        }
        /**
         * Converts an HSB color value to RGB.
         *
         * @param h The hue (between zero and one).
         * @param s The saturation (between zero and one).
         * @param b The brightness (between zero and one).
         * @return An array containing the R, G, and B values (between zero and 255).
         */
        static _hsbToRgb(h: number, s: number, b: number): number[] {
            //assert(h >= 0 && h <= 1 && s >= 0 && s <= 1 && b >= 0 && b <= 1, 'bad HSB values');
            var hsl = Color._hsbToHsl(h, s, b);
            return Color._hslToRgb(hsl[0], hsl[1], hsl[2]);
        }
        /**
         * Converts an HSB color value to HSL.
         *
         * @param h The hue (between zero and one).
         * @param s The saturation (between zero and one).
         * @param b The brightness (between zero and one).
         * @return An array containing the H, S, and L values (between zero and one).
         */
        static _hsbToHsl(h: number, s: number, b: number): number[]{
            assert(h >= 0 && h <= 1 && s >= 0 && s <= 1 && b >= 0 && b <= 1, 'bad HSB values');
            var ll = clamp(b * (2 - s) / 2, 0, 1),
                div = 1 - Math.abs(2 * ll - 1),
                ss = clamp(div > 0 ? b * s / div : s/*0*/, 0, 1);
            assert(!isNaN(ll) && !isNaN(ss), 'bad conversion to HSL');
            return [h, ss, ll];
        }
        /**
         * Converts an HSL color value to HSB.
         *
         * @param h The hue (between zero and one).
         * @param s The saturation (between zero and one).
         * @param l The lightness (between zero and one).
         * @return An array containing the H, S, and B values (between zero and one).
         */
        static _hslToHsb(h: number, s: number, l: number): number[] {
            assert(h >= 0 && h <= 1 && s >= 0 && s <= 1 && l >= 0 && l <= 1, 'bad HSL values');
            var bb = clamp(l == 1 ? 1 : (2 * l + s * (1 - Math.abs(2 * l - 1))) / 2, 0, 1);
            var ss = clamp(bb > 0 ? 2 * (bb - l) / bb : s/*0*/, 0, 1);
            assert(!isNaN(bb) && !isNaN(ss), 'bad conversion to HSB');
            return [h, ss, bb];
        }
    }
}

module wijmo {
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
    export class Clipboard {

        /**
         * Copies a string to the clipboard.
         *
         * This method only works if invoked immediately after the user 
         * pressed a clipboard copy command (such as ctrl+c).
         *
         * @param text Text to copy to the clipboard.
         */
        static copy(text: string) {
            Clipboard._copyPasteInternal(text);
        }
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
        static paste(callback: Function) {
            Clipboard._copyPasteInternal(callback);
        }

        // ** implementation

        private static _copyPasteInternal(textOrCallback: any) {

            // get active element to restore later
            var activeElement = getActiveElement();

            // find parent for temporary input element
            // (body may not be focusable when modal dialogs are used..., TFS 202992)

            // start with closest wijmo control
            var parent = closest(activeElement, '.wj-control');

            // move on to non-control parent (since most controls handle the keyboard themselves)
            while (parent && Control.getControl(parent)) {
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
                setTimeout(() => {
                    var text = el.value;
                    parent.removeChild(el);
                    activeElement.focus();
                    if (typeof (textOrCallback) == 'function') {
                        textOrCallback(text);
                    }
                }, 100); // Apple needs extra timeOut
            }
        }
    }
}

module wijmo {
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
    export function showPopup(popup: HTMLElement, ref: any = null, above = false, fadeIn = false, copyStyles = true) {

        // the popup parent will be the document body (to avoid clipping issues)
        var parent = document.body;

        // if the reference parameter is an element
        if (ref instanceof HTMLElement) {

            // 1 - make sure the reference is in the DOM
            if (!contains(document.body, ref)) {
                return;
            }

            // 2 - adjust the parent to account for ancestors with fixed position or scrollable content
            for (var e = <HTMLElement>ref.parentElement; e; e = e.parentElement) {
                if (getComputedStyle(e).position == 'fixed') {
                    parent = e;
                    break;
                }
            }
        } else {

            // 3 - no reference element, adjust the parent based on the current active element
            // (e.g.grid in dialog, jQuery panel)
            for (var e = getActiveElement(); e; e = e.parentElement) {
                if (!wijmo.hasClass(e, 'wj-popup') && // not our own popups: TFS 206009
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
            var sref = getComputedStyle(ref),
                bkg = new Color(sref.backgroundColor);
            if (bkg.a) { // don't copy anything if the background is transparent!!
                setCss(popup, {
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
        setCss(popup, {
            position: 'absolute',
            visibility: 'hidden',
            display: ''
        });

        // update layout for any Wijmo controls in the popup
        Control.refreshAll(popup);

        // ready to compute margins, size
        var sp = getComputedStyle(popup),
            my = parseFloat(sp.marginTop) + parseFloat(sp.marginBottom),
            mx = parseFloat(sp.marginLeft) + parseFloat(sp.marginRight),
            sz = new Size(popup.offsetWidth + mx, popup.offsetHeight + my);

        // ref can be a mouse event, a point, an element, or a rect
        var pos = new Point(),
            rc = null;
        if (ref && ref.clientX != null && ref.clientY != null && ref.pageX != null && ref.pageY != null) {
            if (ref.clientX <= 0 && ref.clientY <= 0 && ref.target) {
                // this looks like a fake mouse event (e.g. context menu key),
                // so use the event target as a reference TFS 117115
                rc = ref.target.getBoundingClientRect();
            } else {
                // use pageX/Y-offsetX/Y instead of clientX/Y, which gives wrong results in Chrome/Android
                // REVIEW: this seems to work OK, although in some other places we do the opposite... sigh...
                pos.x = Math.max(0, ref.pageX - pageXOffset);
                pos.y = Math.max(0, ref.pageY - pageYOffset);
                //pos.x = Math.max(0, ref.clientX);
                //pos.y = Math.max(0, ref.clientY);
            }
        } else if (ref instanceof Point) { 
            pos = ref;
        } else if (ref instanceof HTMLElement) { 
            rc = ref.getBoundingClientRect();
        } else if (ref && ref.top != null && ref.left != null) {
            rc = ref;
        } else if (ref == null) { // no reference: center a little above the window center
            pos.x = Math.max(0, (innerWidth - sz.width) / 2);
            pos.y = Math.max(0, Math.round((innerHeight - sz.height) / 2 * .7));
        } else {
            throw 'Invalid ref parameter.';
        }

        // calculate min width for the popup
        var minWidth = parseFloat(sp.minWidth);

        // if we have a rect, position popup above or below the rect
        if (rc) {
            var spcAbove = rc.top,
                spcBelow = innerHeight - rc.bottom,
                rtl = getComputedStyle(popup).direction == 'rtl';
            if (rtl) {
                pos.x = Math.max(0, rc.right - sz.width);
            } else { // TODO: honor this
                pos.x = Math.max(0, Math.min(rc.left, innerWidth - sz.width));
            }
            if (above) {
                pos.y = spcAbove > sz.height || spcAbove > spcBelow
                    ? Math.max(0, rc.top - sz.height)
                    : rc.bottom;
            } else {
                pos.y = spcBelow > sz.height || spcBelow > spcAbove
                    ? rc.bottom
                    : Math.max(0, rc.top - sz.height);
            }

            // make popup at least as wide as the element
            minWidth = Math.max(minWidth, rc.width);
        }

        // handle scroll offset (TFS 202906)
        var ptOffset = new Point(0, 0);
        if (ref != null) {
            if (parent == document.body) {
                ptOffset = new Point(-pageXOffset, -pageYOffset);
            } else if (parent) {
                var bcr = parent.getBoundingClientRect();
                ptOffset = new Point(bcr.left - parent.scrollLeft, bcr.top - parent.scrollTop);
            }
        }

        // calculate popup position
        var css: any = {
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
            animate(function (pct) {
                popup.style.opacity = (pct == 1) ? '' : pct.toString();
            });
        }

        // show it
        setCss(popup, css);

        // hide the popup if the user scrolls an ancestor other than the body
        var anchor = ref instanceof MouseEvent ? ref.target : ref;
        if (anchor instanceof HTMLElement) {
            var ctl: Control; // listener to scroll events
            for (var se = anchor.parentElement; se && se != document.body; se = se.parentElement) {
                if (getComputedStyle(se).overflowY == 'auto' && se.scrollHeight > se.offsetHeight) {
                    if (!ctl) {
                        ctl = new Control(document.createElement('div'));
                    }
                    ctl.addEventListener(se, 'scroll', (e) => {
                        _hidePopup(popup, true);
                        ctl.dispose();
                    });
                }
            }
        }
    }
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
    export function hidePopup(popup: HTMLElement, remove = true, fadeOut = false) {
        if (fadeOut) {
            animate(function (pct) {
                popup.style.opacity = (1 - pct).toString();
                if (pct == 1) {
                    _hidePopup(popup, remove);
                    popup.style.opacity = '';
                }
            });
        } else {
            _hidePopup(popup, remove);
        }
    }
    function _hidePopup(popup: HTMLElement, remove) {
        popup.style.display = 'none';
        if (remove && popup.parentElement) {
            popup.parentElement.removeChild(popup);
        }
    }
}

module wijmo {
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
    export class PrintDocument {
        _iframe: HTMLIFrameElement;
        _title: string;
        _css: string[];
        _copyCss = true;

        // ** ctor

        /**
         * Initializes a new instance of the @see:PrintDocument class.
         *
         * @param options JavaScript object containing initialization data for the @see:PrintDocument.
         */
        constructor(options?: any) {
            if (options != null) {
                wijmo.copy(this, options);
            }
        }

        // ** object model

        /**
         * Gets or sets the document title.
         */
        get title(): string {
            return this._title;
        }
        set title(value: string) {
            this._title = asString(value);
        }
        /**
         * Gets or sets a value that determines whether the @see:PrintDocument should include the CSS
         * style sheets defined in the main document.
         */
        get copyCss(): boolean {
            return this._copyCss;
        }
        set copyCss(value: boolean) {
            this._copyCss = asBoolean(value);
        }
        /**
         * Adds a CSS style sheet to the document.
         *
         * @param href URL of the CSS file that should be added to the document.
         */
        addCSS(href: string) {
            if (!this._css) {
                this._css = [];
            }
            this._css.push(href);
        }
        /**
         * Appends an HTML element or string to the document.
         *
         * @param child HTML element or string to append to the document.
         */
        append(child: any) {
            var doc = this._getDocument();
            if (isString(child)) {
                doc.write(child);
            } else if (child instanceof HTMLElement) {
                doc.write(child.outerHTML);
            } else {
                assert(false, 'child should be an HTML element or a string.');
            }
        }
        /**
         * Prints the document.
         */
        print() {
            if (this._iframe) {

                // close the document
                this._close();

                // give it some time before printing/disposing
                setTimeout(() => {

                    // print the document
                    var wnd = this._iframe.contentWindow;
                    wnd.focus();
                    wnd.print();

                    // done, dispose of the iframe
                    document.body.removeChild(this._iframe);
                    this._iframe = null;
                }, 100);
            }
        }

        // ** implementation

        // gets a reference to the print document
        _getDocument(): Document {
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
        }

        // closes the print document before printing
        _close() {

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
                    var link = <HTMLLinkElement>links[i];
                    if (link.href.match(/\.css$/i) && link.rel.match(/stylesheet/i)) {
                        var xhr = httpRequest(link.href, { async: false });
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
                    var es = doc.createElement('style'),
                        xhr = httpRequest(this._css[i], { async: false });
                    es.textContent = xhr.responseText;
                    doc.head.appendChild(es);
                }
            }
        }
        _addStyle(style: string) {
            var doc = this._getDocument(),
                es = doc.createElement('style');
            es.textContent = style;
            doc.head.appendChild(es);
        }
    }
}
module wijmo {
    'use strict';

    /**
     * Class that provides masking services to an HTMLInputElement.
     */
    export class _MaskProvider {
        _tbx: HTMLInputElement;
        _msk: string;
        _promptChar = '_';
        _mskArr: _MaskElement[] = [];
        _firstPos: number;
        _lastPos: number;
        _backSpace: boolean;
        _composing: boolean;
        _full = true;
        _autoComplete: string;
        _spellCheck: boolean;
        _hbInput = this._input.bind(this);
        _hbKeyDown = this._keydown.bind(this);
        _hbKeyPress = this._keypress.bind(this);
        _hbCompositionStart = this._compositionstart.bind(this);
        _hbCompositionEnd = this._compositionend.bind(this);

        // big DBCS/SBCS exception lists
        static _X_DBCS_BIG_HIRA = '\u3041\u3043\u3045\u3047\u3049\u3063\u3083\u3085\u3087\u308e\u3095\u3096';
        static _X_DBCS_BIG_KATA = '\u30a1\u30a3\u30a5\u30a7\u30a9\u30c3\u30e3\u30e5\u30e7\u30ee\u30f5\u30f6';
        static _X_SBCS_BIG_KATA = '\uff67\uff68\uff69\uff6a\uff6b\uff6c\uff6d\uff6e\uff6f';

        /**
         * Initializes a new instance of the @see:_MaskProvider class.
         * 
         * @param input Input element to be masked.
         * @param mask Input mask.
         * @param promptChar Character used to indicate input positions.
         */
        constructor(input: HTMLInputElement, mask = null, promptChar = '_') {
            this.mask = mask;
            this.input = input;
            this.promptChar = promptChar;
            this._connect(true);
        }

        /**
         * Gets or sets the Input element to be masked.
         */
        get input(): HTMLInputElement {
            return this._tbx;
        }
        set input(value: HTMLInputElement) {
            this._connect(false);
            this._tbx = value;
            this._connect(true);
        }
        /**
         * Gets or sets the input mask used to validate input.
         */
        get mask(): string {
            return this._msk;
        }
        set mask(value: string) {
            if (value != this._msk) {
                this._msk = asString(value, true);
                this._parseMask();
                this._valueChanged();
            }
        }
        /**
         * Gets or sets the input mask used to validate input.
         */
        get promptChar(): string {
            return this._promptChar;
        }
        set promptChar(value: string) {
            if (value != this._promptChar) {
                this._promptChar = asString(value, false);
                assert(this._promptChar.length == 1, 'promptChar must be a string with length 1.');
                this._valueChanged();
            }
        }
        /**
         * Gets a value that indicates whether the mask has been completely filled.
         */
        get maskFull(): boolean {
            return this._full;
        }
        /**
         * Gets an array with the position of the first and last wildcard characters in the mask.
         */
        getMaskRange(): number[] {
            return this._mskArr.length ? [this._firstPos, this._lastPos] : [0, this._tbx.value.length - 1];
        }
        /**
         * Gets the raw value of the editor, excluding prompts and literals.
         */
        getRawValue(): string {
            var text = this._tbx.value,
                ret = '';

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
        }
        /**
         * Updates the control mask and content.
         */
        refresh() {
            this._parseMask();
            this._valueChanged();
        }

        // ** event handlers

        // apply mask and update cursor position after any changes
        _input() {
            if (!this._composing) {
                setTimeout(() => { // this timeOut is **important** for Windows Phone/Android/Safari
                    this._valueChanged();
                });
            }
        }

        // special handling for backspacing over literals
        _keydown(e: KeyboardEvent) {

            // ignore backspaces before first wildcard (TFS 199372, 199901)
            if (e.keyCode == Key.Back) {
                var start = this._tbx.selectionStart,
                    end = this._tbx.selectionEnd;
                if (start <= this._firstPos && end == start) {
                    e.preventDefault();
                    this._backSpace = false;
                    return;
                }
            }

            // handle backspacing over literals
            this._backSpace = (e.keyCode == Key.Back);
        }

        // prevent flicker when invalid keys are pressed
        // NOTE: IE on Windows Phone 8.x does not raise keypress!!!
        _keypress(e: KeyboardEvent) {
            if (!e.ctrlKey && !e.metaKey && !e.altKey && !this._composing && this._preventKey(e.charCode)) {
                e.preventDefault();
            }
        }

        // handle IME composition
        _compositionstart(e: KeyboardEvent) {
            this._composing = true;
        }
        _compositionend(e: KeyboardEvent) {
            this._composing = false;
            setTimeout(() => {
                this._valueChanged();
            });
        }

        // ** implementation

        // prevent a key from being handled if it is invalid
        _preventKey(charCode: number): boolean {
            if (charCode && this._mskArr.length) {
                var tbx = this._tbx,
                    start = tbx.selectionStart,
                    key = String.fromCharCode(charCode);

                // make sure we're at a placeholder
                if (start < this._firstPos) {
                    start = this._firstPos;
                    setSelectionRange(tbx, start);
                }

                // past the end?
                if (start >= this._mskArr.length) {
                    return true;
                }

                // skip over literals and validate templates (TFS 117584)
                var m = this._mskArr[start];
                if (m.literal) {
                    this._validatePosition(start);
                } else if (m.wildCard != key && !this._isCharValid(m.wildCard, key)) {
                    return true;
                }
            }

            // key seems OK, do not prevent
            return false;
        }


        // connect or disconnect event handlers for the input element
        _connect(connect: boolean) {
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
                } else {
                    tbx.autocomplete = this._autoComplete;
                    tbx.spellcheck = this._spellCheck;
                    tbx.removeEventListener('input', this._hbInput);
                    tbx.removeEventListener('keydown', this._hbKeyDown, true);
                    tbx.removeEventListener('keypress', this._hbKeyPress, true);
                    tbx.removeEventListener('compositionstart', this._hbCompositionStart, true);
                    tbx.removeEventListener('compositionend', this._hbCompositionEnd, true);
                }
            }
        }

        // apply the mask keeping the cursor position and the focus
        _valueChanged() {
            if (this._tbx && this._msk) {

                // keep track of selection start, character at selection
                var tbx = this._tbx,
                    start = tbx.selectionStart,
                    oldChar = (start > 0) ? tbx.value[start - 1] : '';

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
        }

        // applies the mask to the current text content, returns the result 
        // this is usually a valid string with the same length as the mask, unless 
        // (a) there's no mask, or 
        // (b) there's no text and the input is not required
        _applyMask(): string {

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
            var ret = '',
                pos = 0;
            text = this._handleVagueLiterals(text);

            // build output string based on current content and mask
            for (var i = 0; i < this._mskArr.length; i++) {

                // get mask element
                var m = this._mskArr[i],
                    c = m.literal;

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
        }

        // fix text to handle vague literals (that could be interpreted as content)
        _handleVagueLiterals(text: string): string {

            // looks like a paste, don't mess with it (TFS 139412, 145560)
            if (text.length > this._mskArr.length + 1) {
                return text;
            }

            // see if we're shrinking or growing
            var delta = text.length - this._mskArr.length;
            if (delta != 0 && text.length > 1) {

                // see if we have a 'vague' literal
                var badIndex = -1,
                    start = Math.max(0, this._tbx.selectionStart - delta);
                for (var i = start; i < this._mskArr.length; i++) {
                    if (this._mskArr[i].vague) {
                        badIndex = i;
                        break;
                    }
                }

                // we do, so handle it
                if (badIndex > -1) {
                    //console.log(' text: [' + text + ']');
                    if (delta < 0) { // deleted: pad
                        var pad = Array(1 - delta).join(this._promptChar),
                            index = badIndex + delta;
                        if (index > -1) {
                            text = text.substr(0, index) + pad + text.substr(index);
                        }
                    } else { // added: line up
                        while (badIndex > 0 && this._mskArr[badIndex - 1].literal) {
                            badIndex--;
                        }
                        text = text.substr(0, badIndex) + text.substr(badIndex + delta);
                    }
                    //console.log('fixed: [' + text + ']');
                }
            }

            // done
            return text;
        }

        // checks whether a character is valid for a given mask character
        _isCharValid(mask: string, c: string) {
            var ph = this._promptChar;
            switch (mask) {

                // regular wildcards
                case '0': // Digit
                    return (c >= '0' && c <= '9') || c == ph;
                case '9': // Digit or space
                    return (c >= '0' && c <= '9') || c == ' ' || c == ph;
                case '#': // Digit, sign, or space
                    return (c >= '0' && c <= '9') || c == ' ' || c == '+' || c == '-' || c == ph;
                case 'L': // Letter
                    return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || c == ph;
                case 'l': // Letter or space
                    return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || c == ' ' || c == ph;
                case 'A': // Alphanumeric
                    return (c >= '0' && c <= '9') || (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || c == ph;
                case 'a': // Alphanumeric or space
                    return (c >= '0' && c <= '9') || (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || c == ' ' || c == ph;

                // Katakana/Hiragana wildcards
                case '\uff19': // DBCS Digit
                    return (c >= '\uFF10' && c <= '\uff19') || c == ph;
                case '\uff2a': // DBCS Hiragana
                case '\uff27': // DBCS big Hiragana
                    if (mask == '\uff27' && _MaskProvider._X_DBCS_BIG_HIRA.indexOf(c) > -1) return false;  // exceptions
                    return (c >= '\u3041' && c <= '\u3096') || c == ph;
                case '\uff2b': // DBCS Katakana
                case '\uff2e': // DBCS big Katakana
                    if (mask == '\uff2e' && _MaskProvider._X_DBCS_BIG_KATA.indexOf(c) > -1) return false; // exceptions
                    return (c >= '\u30a1' && c <= '\u30fa') || c == ph;
                case '\uff3a': // Any DBCS
                    return (c <= '\u0021' || c >= '\u00ff') || c == ph;
                case 'H': // Any SBCS
                    return (c >= '\u0021' && c <= '\u00ff') || c == ph;
                case 'K': // SBCS Katakana
                case 'N': // SBCS big Katakana
                    if (mask == 'N' && _MaskProvider._X_SBCS_BIG_KATA.indexOf(c) > -1) return false; // exceptions
                    return (c >= '\uff66' && c <= '\uff9f') || c == ph;
            }
            return false;
        }

        // skip over literals
        _validatePosition(start: number) {
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
            if (getActiveElement() == this._tbx) {
                setSelectionRange(this._tbx, start);
            }

            // no longer backspacing
            this._backSpace = false;
        }

        // parse mask into internal mask, literals, and case
        _parseMask() {

            // clear internal mask info
            this._mskArr = [];
            this._firstPos = -1;
            this._lastPos = -1;

            // parse new mask
            var msk = this._msk,
                currCase = '|',
                c: string;
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
                    case 'H': // Any SBCS character.

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
                    case '$': // Currency symbol.
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
                    case '|': // Disable any previous shifts.
                        currCase = msk[i];
                        break;

                    // literals
                    case '\\': // Escape next character into literal.
                        if (i < msk.length - 1) i++;
                        this._mskArr.push(new _MaskElement(msk[i]));
                        break;
                    default: // All others: Literals.
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
        }
    }

    /**
     * Class that contains information about a position in an input mask.
     */
    export class _MaskElement {
        wildCard: string;   // wildcard to match
        charCase: string;   // upper/lower case
        literal: string;    // literal to match
        vague: boolean      // literal that can be interpreted as content

        /**
         * Initializes a new instance of the @see:_MaskElement class.
         *
         * @param wildcardOrLiteral Wildcard or literal character
         * @param charCase Whether to convert wildcard matches to upper or lowercase.
         */
        constructor(wildcardOrLiteral: string, charCase?: string) {
            if (charCase) {
                this.wildCard = wildcardOrLiteral;
                this.charCase = charCase;
            } else {
                this.literal = wildcardOrLiteral;
            }
        }
    }

}
//
// IE9 polyfills
module wijmo {
    'use strict';

    // expose IE flag to callers that need it
    var _isIE = navigator.userAgent.match(/MSIE |Trident\/|Edge\//) != null;
    export function isIE(): boolean {
        return _isIE;
    }
    
    // expose IE9 flag to callers that need it
    var _isIE9: boolean;
    export function isIE9(): boolean {
        return _isIE9;
    }

    // implement HTML5 drag-drop behavior in IE9.
    if (document.doctype && navigator.appVersion.indexOf('MSIE 9') > -1) {

        // remember this is IE9...
        _isIE9 = true;

        // TFS 140812: 'selectstart' does not work in popup dialogs, so use 'mousemove'
        // instead. It's less efficient but it works, and this only matters in IE9.
        document.addEventListener('mousemove', function (e: MouseEvent) {
            if (e.which == 1) {
                for (var el = <Node>e.target; el; el = el.parentNode) {
                    if (el.attributes && el.attributes['draggable']) {
                        (<HTMLElement>el).dragDrop();
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
            var
                currentTime = Date.now(),
                adjustedDelay = 16 - (currentTime - expectedTime),
                delay = adjustedDelay > 0 ? adjustedDelay : 0;
            expectedTime = currentTime + delay;
            return setTimeout(function () {
                callback(expectedTime);
            }, delay);
        };
        window['cancelAnimationFrame'] = clearTimeout;
    }
}
