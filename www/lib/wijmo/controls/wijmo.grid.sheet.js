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
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var sheet;
        (function (sheet_1) {
            'use strict';
            /*
             * Defines the CalcEngine class.
             *
             * It deals with the calculation for the flexsheet control.
             */
            var _CalcEngine = (function () {
                /*
                 * Initializes a new instance of the @see:CalcEngine class.
                 *
                 * @param owner The @see: FlexSheet control that the CalcEngine works for.
                 */
                function _CalcEngine(owner) {
                    this._expressionCache = {};
                    this._idChars = '$:!';
                    this._functionTable = {};
                    this._cacheSize = 0;
                    /*
                     * Occurs when the @see:_CalcEngine meets the unknown formula.
                     */
                    this.unknownFunction = new wijmo.Event();
                    this._owner = owner;
                    this._buildSymbolTable();
                    this._registerAggregateFunction();
                    this._registerMathFunction();
                    this._registerLogicalFunction();
                    this._registerTextFunction();
                    this._registerDateFunction();
                    this._registLookUpReferenceFunction();
                    this._registFinacialFunction();
                }
                /*
                 * Raises the unknownFunction event.
                 */
                _CalcEngine.prototype.onUnknownFunction = function (funcName, params) {
                    var paramsList, eventArgs;
                    if (params && params.length > 0) {
                        paramsList = [];
                        for (var i = 0; i < params.length; i++) {
                            paramsList[i] = params[i].evaluate();
                        }
                    }
                    eventArgs = new sheet_1.UnknownFunctionEventArgs(funcName, paramsList);
                    this.unknownFunction.raise(this, eventArgs);
                    if (eventArgs.value != null) {
                        return new sheet_1._Expression(eventArgs.value);
                    }
                    throw 'The function "' + funcName + '"' + ' has not supported in FlexSheet yet.';
                };
                /*
                 * Evaluates an expression.
                 *
                 * @param expression the expression need to be evaluated to value.
                 * @param format the format string used to convert raw values into display.
                 * @param sheet The @see:Sheet is referenced by the @see:Expression.
                 * @param rowIndex The row index of the cell where the expression located in.
                 * @param columnIndex The column index of the cell where the expression located in.
                 */
                _CalcEngine.prototype.evaluate = function (expression, format, sheet, rowIndex, columnIndex) {
                    var expr, result;
                    try {
                        if (expression && expression.length > 1 && expression[0] === '=') {
                            expr = this._checkCache(expression);
                            result = expr.evaluate(sheet, rowIndex, columnIndex);
                            while (result instanceof sheet_1._Expression) {
                                result = result.evaluate(sheet);
                            }
                            if (format && wijmo.isPrimitive(result)) {
                                return wijmo.Globalize.format(result, format);
                            }
                            return result;
                        }
                        return expression ? expression : '';
                    }
                    catch (e) {
                        return "Error: " + e;
                    }
                };
                /*
                 * Add a custom function to the @see:_CalcEngine.
                 *
                 * @param name the name of the custom function, the function name should be lower case.
                 * @param func the custom function.
                 * @param minParamsCount the minimum count of the parameter that the function need.
                 * @param maxParamsCount the maximum count of the parameter that the function need.
                 *        If the count of the parameters in the custom function is arbitrary, the
                 *        minParamsCount and maxParamsCount should be set to null.
                 */
                _CalcEngine.prototype.addCustomFunction = function (name, func, minParamsCount, maxParamsCount) {
                    var self = this;
                    name = name.toLowerCase();
                    this._functionTable[name] = new _FunctionDefinition(function (params) {
                        var param, paramsList = [];
                        if (params.length > 0) {
                            for (var i = 0; i < params.length; i++) {
                                param = params[i];
                                if (param instanceof sheet_1._CellRangeExpression) {
                                    paramsList[i] = param.cells;
                                }
                                else {
                                    paramsList[i] = param.evaluate();
                                }
                            }
                        }
                        return func.apply(self, paramsList);
                    }, maxParamsCount, minParamsCount);
                };
                // Clear the expression cache.
                _CalcEngine.prototype._clearExpressionCache = function () {
                    this._expressionCache = null;
                    this._expressionCache = {};
                    this._cacheSize = 0;
                };
                // Parse the string expression to an Expression instance that can be evaluated to value.
                _CalcEngine.prototype._parse = function (expression) {
                    this._expression = expression;
                    this._expressLength = expression ? expression.length : 0;
                    this._pointer = 0;
                    // skip leading equals sign
                    if (this._expressLength > 0 && this._expression[0] === '=') {
                        this._pointer++;
                    }
                    return this._parseExpression();
                };
                // Build static token table.
                _CalcEngine.prototype._buildSymbolTable = function () {
                    if (!this._tokenTable) {
                        this._tokenTable = {};
                        this._addToken('+', _TokenID.ADD, _TokenType.ADDSUB);
                        this._addToken('-', _TokenID.SUB, _TokenType.ADDSUB);
                        this._addToken('(', _TokenID.OPEN, _TokenType.GROUP);
                        this._addToken(')', _TokenID.CLOSE, _TokenType.GROUP);
                        this._addToken('*', _TokenID.MUL, _TokenType.MULDIV);
                        this._addToken(',', _TokenID.COMMA, _TokenType.GROUP);
                        this._addToken('.', _TokenID.PERIOD, _TokenType.GROUP);
                        this._addToken('/', _TokenID.DIV, _TokenType.MULDIV);
                        this._addToken('\\', _TokenID.DIVINT, _TokenType.MULDIV);
                        this._addToken('=', _TokenID.EQ, _TokenType.COMPARE);
                        this._addToken('>', _TokenID.GT, _TokenType.COMPARE);
                        this._addToken('<', _TokenID.LT, _TokenType.COMPARE);
                        this._addToken('^', _TokenID.POWER, _TokenType.POWER);
                        this._addToken("<>", _TokenID.NE, _TokenType.COMPARE);
                        this._addToken(">=", _TokenID.GE, _TokenType.COMPARE);
                        this._addToken("<=", _TokenID.LE, _TokenType.COMPARE);
                        this._addToken('&', _TokenID.CONCAT, _TokenType.CONCAT);
                    }
                };
                // Register the aggregate function for the CalcEngine.
                _CalcEngine.prototype._registerAggregateFunction = function () {
                    var self = this;
                    self._functionTable['sum'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getAggregateResult(wijmo.Aggregate.Sum, params, sheet);
                    });
                    self._functionTable['average'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getAggregateResult(wijmo.Aggregate.Avg, params, sheet);
                    });
                    self._functionTable['max'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getAggregateResult(wijmo.Aggregate.Max, params, sheet);
                    });
                    self._functionTable['min'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getAggregateResult(wijmo.Aggregate.Min, params, sheet);
                    });
                    self._functionTable['var'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getAggregateResult(wijmo.Aggregate.Var, params, sheet);
                    });
                    self._functionTable['varp'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getAggregateResult(wijmo.Aggregate.VarPop, params, sheet);
                    });
                    self._functionTable['stdev'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getAggregateResult(wijmo.Aggregate.Std, params, sheet);
                    });
                    self._functionTable['stdevp'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getAggregateResult(wijmo.Aggregate.StdPop, params, sheet);
                    });
                    self._functionTable['count'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getFlexSheetAggregateResult(_FlexSheetAggregate.Count, params, sheet);
                    });
                    self._functionTable['counta'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getFlexSheetAggregateResult(_FlexSheetAggregate.CountA, params, sheet);
                    });
                    self._functionTable['countblank'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getFlexSheetAggregateResult(_FlexSheetAggregate.ConutBlank, params, sheet);
                    });
                    self._functionTable['countif'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getFlexSheetAggregateResult(_FlexSheetAggregate.CountIf, params, sheet);
                    }, 2, 2);
                    self._functionTable['countifs'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getFlexSheetAggregateResult(_FlexSheetAggregate.CountIfs, params, sheet);
                    }, 254, 2);
                    self._functionTable['sumif'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getFlexSheetAggregateResult(_FlexSheetAggregate.SumIf, params, sheet);
                    }, 3, 2);
                    self._functionTable['sumifs'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getFlexSheetAggregateResult(_FlexSheetAggregate.SumIfs, params, sheet);
                    }, 255, 2);
                    self._functionTable['rank'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getFlexSheetAggregateResult(_FlexSheetAggregate.Rank, params, sheet);
                    }, 3, 2);
                    self._functionTable['product'] = new _FunctionDefinition(function (params, sheet) {
                        return self._getFlexSheetAggregateResult(_FlexSheetAggregate.Product, params, sheet);
                    }, 255, 1);
                    self._functionTable['subtotal'] = new _FunctionDefinition(function (params, sheet) {
                        return self._handleSubtotal(params, sheet);
                    }, 255, 2);
                    self._functionTable['dcount'] = new _FunctionDefinition(function (params, sheet) {
                        return self._handleDCount(params, sheet);
                    }, 3, 3);
                };
                // Register the math function for the calcEngine.
                _CalcEngine.prototype._registerMathFunction = function () {
                    var self = this, unaryFuncs = ['abs', 'acos', 'asin', 'atan', 'ceiling', 'cos', 'exp', 'floor', 'ln', 'sin', 'sqrt', 'tan'], roundFuncs = ['round', 'rounddown', 'roundup'];
                    self._functionTable['pi'] = new _FunctionDefinition(function () {
                        return Math.PI;
                    }, 0, 0);
                    self._functionTable['rand'] = new _FunctionDefinition(function () {
                        return Math.random();
                    }, 0, 0);
                    self._functionTable['power'] = new _FunctionDefinition(function (params, sheet) {
                        return Math.pow(sheet_1._Expression.toNumber(params[0], sheet), sheet_1._Expression.toNumber(params[1], sheet));
                    }, 2, 2);
                    self._functionTable['atan2'] = new _FunctionDefinition(function (params, sheet) {
                        var x = sheet_1._Expression.toNumber(params[0], sheet), y = sheet_1._Expression.toNumber(params[1], sheet);
                        if (x === 0 && y === 0) {
                            throw 'The x number and y number can\'t both be zero for the atan2 function';
                        }
                        return Math.atan2(y, x);
                    }, 2, 2);
                    self._functionTable['mod'] = new _FunctionDefinition(function (params, sheet) {
                        return sheet_1._Expression.toNumber(params[0], sheet) % sheet_1._Expression.toNumber(params[1], sheet);
                    }, 2, 2);
                    self._functionTable['trunc'] = new _FunctionDefinition(function (params, sheet) {
                        var num = sheet_1._Expression.toNumber(params[0], sheet), precision = params.length === 2 ? sheet_1._Expression.toNumber(params[1], sheet) : 0, multiple;
                        if (precision === 0) {
                            if (num >= 0) {
                                return Math.floor(num);
                            }
                            else {
                                return Math.ceil(num);
                            }
                        }
                        else {
                            multiple = Math.pow(10, precision);
                            if (num >= 0) {
                                return Math.floor(num * multiple) / multiple;
                            }
                            else {
                                return Math.ceil(num * multiple) / multiple;
                            }
                        }
                    }, 2, 1);
                    roundFuncs.forEach(function (val) {
                        self._functionTable[val] = new _FunctionDefinition(function (params, sheet) {
                            var num = sheet_1._Expression.toNumber(params[0], sheet), precision = sheet_1._Expression.toNumber(params[1], sheet), result, format, multiple;
                            if (precision === 0) {
                                switch (val) {
                                    case 'rounddown':
                                        if (num >= 0) {
                                            result = Math.floor(num);
                                        }
                                        else {
                                            result = Math.ceil(num);
                                        }
                                        break;
                                    case 'roundup':
                                        if (num >= 0) {
                                            result = Math.ceil(num);
                                        }
                                        else {
                                            result = Math.floor(num);
                                        }
                                        break;
                                    case 'round':
                                        result = Math.round(num);
                                        break;
                                    default:
                                        result = Math.floor(num);
                                        break;
                                }
                                format = 'n0';
                            }
                            else if (precision > 0 && wijmo.isInt(precision)) {
                                multiple = Math.pow(10, precision);
                                switch (val) {
                                    case 'rounddown':
                                        if (num >= 0) {
                                            result = Math.floor(num * multiple) / multiple;
                                        }
                                        else {
                                            result = Math.ceil(num * multiple) / multiple;
                                        }
                                        break;
                                    case 'roundup':
                                        if (num >= 0) {
                                            result = Math.ceil(num * multiple) / multiple;
                                        }
                                        else {
                                            result = Math.floor(num * multiple) / multiple;
                                        }
                                        break;
                                    case 'round':
                                        result = Math.round(num * multiple) / multiple;
                                        break;
                                }
                                format = 'n' + precision;
                            }
                            if (result != null) {
                                return {
                                    value: result,
                                    format: format
                                };
                            }
                            throw 'Invalid precision!';
                        }, 2, 2);
                    });
                    unaryFuncs.forEach(function (val) {
                        self._functionTable[val] = new _FunctionDefinition(function (params, sheet) {
                            switch (val) {
                                case 'ceiling':
                                    return Math.ceil(sheet_1._Expression.toNumber(params[0], sheet));
                                case 'ln':
                                    return Math.log(sheet_1._Expression.toNumber(params[0], sheet));
                                default:
                                    return Math[val](sheet_1._Expression.toNumber(params[0], sheet));
                            }
                        }, 1, 1);
                    });
                };
                // Register the logical function for the calcEngine.
                _CalcEngine.prototype._registerLogicalFunction = function () {
                    // and(true,true,false,...)
                    this._functionTable['and'] = new _FunctionDefinition(function (params, sheet) {
                        var result = true, index;
                        for (index = 0; index < params.length; index++) {
                            result = result && sheet_1._Expression.toBoolean(params[index], sheet);
                            if (!result) {
                                break;
                            }
                        }
                        return result;
                    }, Number.MAX_VALUE, 1);
                    // or(false,true,true,...)
                    this._functionTable['or'] = new _FunctionDefinition(function (params, sheet) {
                        var result = false, index;
                        for (index = 0; index < params.length; index++) {
                            result = result || sheet_1._Expression.toBoolean(params[index], sheet);
                            if (result) {
                                break;
                            }
                        }
                        return result;
                    }, Number.MAX_VALUE, 1);
                    // not(false)
                    this._functionTable['not'] = new _FunctionDefinition(function (params, sheet) {
                        return !sheet_1._Expression.toBoolean(params[0], sheet);
                    }, 1, 1);
                    // if(true,a,b)
                    this._functionTable['if'] = new _FunctionDefinition(function (params, sheet) {
                        return sheet_1._Expression.toBoolean(params[0], sheet) ? params[1].evaluate(sheet) : params[2].evaluate(sheet);
                    }, 3, 3);
                    // true()
                    this._functionTable['true'] = new _FunctionDefinition(function () {
                        return true;
                    }, 0, 0);
                    // false()
                    this._functionTable['false'] = new _FunctionDefinition(function () {
                        return false;
                    }, 0, 0);
                };
                // register the text process function
                _CalcEngine.prototype._registerTextFunction = function () {
                    // char(65, 66, 67,...) => "abc"
                    this._functionTable['char'] = new _FunctionDefinition(function (params, sheet) {
                        var index, result = '';
                        for (index = 0; index < params.length; index++) {
                            result += String.fromCharCode(sheet_1._Expression.toNumber(params[index], sheet));
                        }
                        return result;
                    }, Number.MAX_VALUE, 1);
                    // code("A")
                    this._functionTable['code'] = new _FunctionDefinition(function (params, sheet) {
                        var str = sheet_1._Expression.toString(params[0], sheet);
                        if (str && str.length > 0) {
                            return str.charCodeAt(0);
                        }
                        return -1;
                    }, 1, 1);
                    // concatenate("abc","def","ghi",...) => "abcdefghi"
                    this._functionTable['concatenate'] = new _FunctionDefinition(function (params, sheet) {
                        var index, result = '';
                        for (index = 0; index < params.length; index++) {
                            result = result.concat(sheet_1._Expression.toString(params[index], sheet));
                        }
                        return result;
                    }, Number.MAX_VALUE, 1);
                    // left("Abcdefgh", 5) => "Abcde"
                    this._functionTable['left'] = new _FunctionDefinition(function (params, sheet) {
                        var str = sheet_1._Expression.toString(params[0], sheet), length = Math.floor(sheet_1._Expression.toNumber(params[1], sheet));
                        if (str && str.length > 0) {
                            return str.slice(0, length);
                        }
                        return undefined;
                    }, 2, 2);
                    // right("Abcdefgh", 5) => "defgh"
                    this._functionTable['right'] = new _FunctionDefinition(function (params, sheet) {
                        var str = sheet_1._Expression.toString(params[0], sheet), length = Math.floor(sheet_1._Expression.toNumber(params[1], sheet));
                        if (str && str.length > 0) {
                            return str.slice(-length);
                        }
                        return undefined;
                    }, 2, 2);
                    // find("abc", "abcdefgh") 
                    // this function is case-sensitive.
                    this._functionTable['find'] = new _FunctionDefinition(function (params, sheet) {
                        var search = sheet_1._Expression.toString(params[0], sheet), text = sheet_1._Expression.toString(params[1], sheet), result;
                        if (text != null && search != null) {
                            result = text.indexOf(search);
                            if (result > -1) {
                                return result + 1;
                            }
                        }
                        return -1;
                    }, 2, 2);
                    // search("abc", "ABCDEFGH") 
                    // this function is not case-sensitive.
                    this._functionTable['search'] = new _FunctionDefinition(function (params, sheet) {
                        var search = sheet_1._Expression.toString(params[0], sheet), text = sheet_1._Expression.toString(params[1], sheet), searchRegExp, result;
                        if (text != null && search != null) {
                            searchRegExp = new RegExp(search, 'i');
                            result = text.search(searchRegExp);
                            if (result > -1) {
                                return result + 1;
                            }
                        }
                        return -1;
                    }, 2, 2);
                    // len("abcdefgh")
                    this._functionTable['len'] = new _FunctionDefinition(function (params, sheet) {
                        var str = sheet_1._Expression.toString(params[0], sheet);
                        if (str) {
                            return str.length;
                        }
                        return -1;
                    }, 1, 1);
                    //  mid("abcdefgh", 2, 3) => "bcd"
                    this._functionTable['mid'] = new _FunctionDefinition(function (params, sheet) {
                        var text = sheet_1._Expression.toString(params[0], sheet), start = Math.floor(sheet_1._Expression.toNumber(params[1], sheet)), length = Math.floor(sheet_1._Expression.toNumber(params[2], sheet));
                        if (text && text.length > 0 && start > 0) {
                            return text.substr(start - 1, length);
                        }
                        return undefined;
                    }, 3, 3);
                    // lower("ABCDEFGH")
                    this._functionTable['lower'] = new _FunctionDefinition(function (params, sheet) {
                        var str = sheet_1._Expression.toString(params[0], sheet);
                        if (str && str.length > 0) {
                            return str.toLowerCase();
                        }
                        return undefined;
                    }, 1, 1);
                    // upper("abcdefgh")
                    this._functionTable['upper'] = new _FunctionDefinition(function (params, sheet) {
                        var str = sheet_1._Expression.toString(params[0], sheet);
                        if (str && str.length > 0) {
                            return str.toUpperCase();
                        }
                        return undefined;
                    }, 1, 1);
                    // proper("abcdefgh") => "Abcdefgh"
                    this._functionTable['proper'] = new _FunctionDefinition(function (params, sheet) {
                        var str = sheet_1._Expression.toString(params[0], sheet);
                        if (str && str.length > 0) {
                            return str[0].toUpperCase() + str.substring(1).toLowerCase();
                        }
                        return undefined;
                    }, 1, 1);
                    // trim("   abcdefgh   ") => "abcdefgh"
                    this._functionTable['trim'] = new _FunctionDefinition(function (params, sheet) {
                        var str = sheet_1._Expression.toString(params[0], sheet);
                        if (str && str.length > 0) {
                            return str.trim();
                        }
                        return undefined;
                    }, 1, 1);
                    // replace("abcdefg", 2, 3, "xyz") => "axyzefg"
                    this._functionTable['replace'] = new _FunctionDefinition(function (params, sheet) {
                        var text = sheet_1._Expression.toString(params[0], sheet), start = Math.floor(sheet_1._Expression.toNumber(params[1], sheet)), length = Math.floor(sheet_1._Expression.toNumber(params[2], sheet)), replaceText = sheet_1._Expression.toString(params[3], sheet);
                        if (text && text.length > 0 && start > 0) {
                            return text.substring(0, start - 1) + replaceText + text.slice(start - 1 + length);
                        }
                        return undefined;
                    }, 4, 4);
                    // substitute("abcabcdabcdefgh", "ab", "xy") => "xycxycdxycdefg"
                    this._functionTable['substitute'] = new _FunctionDefinition(function (params, sheet) {
                        var text = sheet_1._Expression.toString(params[0], sheet), oldText = sheet_1._Expression.toString(params[1], sheet), newText = sheet_1._Expression.toString(params[2], sheet), searhRegExp;
                        if (text && text.length > 0 && oldText && oldText.length > 0) {
                            searhRegExp = new RegExp(oldText, 'g');
                            return text.replace(searhRegExp, newText);
                        }
                        return undefined;
                    }, 3, 3);
                    // rept("abc", 3) => "abcabcabc"
                    this._functionTable['rept'] = new _FunctionDefinition(function (params, sheet) {
                        var text = sheet_1._Expression.toString(params[0], sheet), repeatTimes = Math.floor(sheet_1._Expression.toNumber(params[1], sheet)), result = '', i;
                        if (text && text.length > 0 && repeatTimes > 0) {
                            for (i = 0; i < repeatTimes; i++) {
                                result = result.concat(text);
                            }
                        }
                        return result;
                    }, 2, 2);
                    // text("1234", "n2") => "1234.00"
                    this._functionTable['text'] = new _FunctionDefinition(function (params, sheet) {
                        var value = params[0].evaluate(), format = sheet_1._Expression.toString(params[1], sheet);
                        return wijmo.Globalize.format(value, format);
                    }, 2, 2);
                    // value("1234") => 1234
                    this._functionTable['value'] = new _FunctionDefinition(function (params, sheet) {
                        return sheet_1._Expression.toNumber(params[0], sheet);
                    }, 1, 1);
                };
                // Register the datetime function for the calcEngine.
                _CalcEngine.prototype._registerDateFunction = function () {
                    this._functionTable['now'] = new _FunctionDefinition(function () {
                        return {
                            value: new Date(),
                            format: 'M/d/yyyy h:mm'
                        };
                    }, 0, 0);
                    this._functionTable['today'] = new _FunctionDefinition(function () {
                        return {
                            value: new Date(),
                            format: 'd'
                        };
                    }, 0, 0);
                    // year("11/25/2015") => 2015
                    this._functionTable['year'] = new _FunctionDefinition(function (params, sheet) {
                        var date = sheet_1._Expression.toDate(params[0], sheet);
                        if (!wijmo.isPrimitive(date) && date) {
                            return date.value;
                        }
                        if (wijmo.isDate(date)) {
                            return date.getFullYear();
                        }
                        return 1900;
                    }, 1, 1);
                    // month("11/25/2015") => 11
                    this._functionTable['month'] = new _FunctionDefinition(function (params, sheet) {
                        var date = sheet_1._Expression.toDate(params[0], sheet);
                        if (!wijmo.isPrimitive(date) && date) {
                            return date.value;
                        }
                        if (wijmo.isDate(date)) {
                            return date.getMonth() + 1;
                        }
                        return 1;
                    }, 1, 1);
                    // day("11/25/2015") => 25
                    this._functionTable['day'] = new _FunctionDefinition(function (params, sheet) {
                        var date = sheet_1._Expression.toDate(params[0], sheet);
                        if (!wijmo.isPrimitive(date) && date) {
                            return date.value;
                        }
                        if (wijmo.isDate(date)) {
                            return date.getDate();
                        }
                        return 0;
                    }, 1, 1);
                    // hour("11/25/2015 16:50") => 16 or hour(0.5) => 12
                    this._functionTable['hour'] = new _FunctionDefinition(function (params, sheet) {
                        var val = params[0].evaluate(sheet);
                        if (wijmo.isNumber(val) && !isNaN(val)) {
                            return Math.floor(24 * (val - Math.floor(val)));
                        }
                        else if (wijmo.isDate(val)) {
                            return val.getHours();
                        }
                        val = sheet_1._Expression.toDate(params[0], sheet);
                        if (!wijmo.isPrimitive(val) && val) {
                            val = val.value;
                        }
                        if (wijmo.isDate(val)) {
                            return val.getHours();
                        }
                        throw 'Invalid parameter.';
                    }, 1, 1);
                    // time(10, 23, 11) => 10:23:11 AM
                    this._functionTable['time'] = new _FunctionDefinition(function (params, sheet) {
                        var hour = params[0].evaluate(sheet), minute = params[1].evaluate(sheet), second = params[2].evaluate(sheet);
                        if (wijmo.isNumber(hour) && wijmo.isNumber(minute) && wijmo.isNumber(second)) {
                            hour %= 24;
                            minute %= 60;
                            second %= 60;
                            return {
                                value: new Date(0, 0, 0, hour, minute, second),
                                format: 't'
                            };
                        }
                        throw 'Invalid parameters.';
                    }, 3, 3);
                    // time(2015, 11, 25) => 11/25/2015
                    this._functionTable['date'] = new _FunctionDefinition(function (params, sheet) {
                        var year = params[0].evaluate(sheet), month = params[1].evaluate(sheet), day = params[2].evaluate(sheet);
                        if (wijmo.isNumber(year) && wijmo.isNumber(month) && wijmo.isNumber(day)) {
                            return {
                                value: new Date(year, month - 1, day),
                                format: 'd'
                            };
                        }
                        throw 'Invalid parameters.';
                    }, 3, 3);
                    this._functionTable['datedif'] = new _FunctionDefinition(function (params, sheet) {
                        var startDate = sheet_1._Expression.toDate(params[0], sheet), endDate = sheet_1._Expression.toDate(params[1], sheet), unit = params[2].evaluate(sheet), startDateTime, endDateTime, diffDays, diffMonths, diffYears;
                        if (!wijmo.isPrimitive(startDate) && startDate) {
                            startDate = startDate.value;
                        }
                        if (!wijmo.isPrimitive(endDate) && endDate) {
                            endDate = endDate.value;
                        }
                        if (wijmo.isDate(startDate) && wijmo.isDate(endDate) && wijmo.isString(unit)) {
                            startDateTime = startDate.getTime();
                            endDateTime = endDate.getTime();
                            if (startDateTime > endDateTime) {
                                throw 'Start date is later than end date.';
                            }
                            diffDays = endDate.getDate() - startDate.getDate();
                            diffMonths = endDate.getMonth() - startDate.getMonth();
                            diffYears = endDate.getFullYear() - startDate.getFullYear();
                            switch (unit.toUpperCase()) {
                                case 'Y':
                                    if (diffMonths > 0) {
                                        return diffYears;
                                    }
                                    else if (diffMonths < 0) {
                                        return diffYears - 1;
                                    }
                                    else {
                                        if (diffDays >= 0) {
                                            return diffYears;
                                        }
                                        else {
                                            return diffYears - 1;
                                        }
                                    }
                                case 'M':
                                    if (diffDays >= 0) {
                                        return diffYears * 12 + diffMonths;
                                    }
                                    else {
                                        return diffYears * 12 + diffMonths - 1;
                                    }
                                case 'D':
                                    return (endDateTime - startDateTime) / (1000 * 3600 * 24);
                                case 'YM':
                                    if (diffDays >= 0) {
                                        diffMonths = diffYears * 12 + diffMonths;
                                    }
                                    else {
                                        diffMonths = diffYears * 12 + diffMonths - 1;
                                    }
                                    return diffMonths % 12;
                                case 'YD':
                                    if (diffMonths > 0) {
                                        return (new Date(startDate.getFullYear(), endDate.getMonth(), endDate.getDate()).getTime() - startDate.getTime()) / (1000 * 3600 * 24);
                                    }
                                    else if (diffMonths < 0) {
                                        return (new Date(startDate.getFullYear() + 1, endDate.getMonth(), endDate.getDate()).getTime() - startDate.getTime()) / (1000 * 3600 * 24);
                                    }
                                    else {
                                        if (diffDays >= 0) {
                                            return diffDays;
                                        }
                                        else {
                                            return (new Date(startDate.getFullYear() + 1, endDate.getMonth(), endDate.getDate()).getTime() - startDate.getTime()) / (1000 * 3600 * 24);
                                        }
                                    }
                                case 'MD':
                                    if (diffDays >= 0) {
                                        return diffDays;
                                    }
                                    else {
                                        diffDays = new Date(endDate.getFullYear(), endDate.getMonth(), 0).getDate() - new Date(endDate.getFullYear(), endDate.getMonth() - 1, 1).getDate() + 1 + diffDays;
                                        return diffDays;
                                    }
                                default:
                                    throw 'Invalid unit.';
                            }
                        }
                        throw 'Invalid parameters.';
                    }, 3, 3);
                };
                // Register the cell reference and look up related functions for the calcEngine.
                _CalcEngine.prototype._registLookUpReferenceFunction = function () {
                    var self = this;
                    self._functionTable['column'] = new _FunctionDefinition(function (params, sheet, rowIndex, columnIndex) {
                        var cellExpr;
                        if (params == null) {
                            return columnIndex + 1;
                        }
                        cellExpr = params[0];
                        cellExpr = self._ensureNonFunctionExpression(cellExpr);
                        if (cellExpr instanceof sheet_1._CellRangeExpression) {
                            return cellExpr.cells.col + 1;
                        }
                        throw 'Invalid Cell Reference.';
                    }, 1, 0);
                    self._functionTable['columns'] = new _FunctionDefinition(function (params, sheet) {
                        var cellExpr = params[0];
                        cellExpr = self._ensureNonFunctionExpression(cellExpr);
                        if (cellExpr instanceof sheet_1._CellRangeExpression) {
                            return cellExpr.cells.columnSpan;
                        }
                        throw 'Invalid Cell Reference.';
                    }, 1, 1);
                    self._functionTable['row'] = new _FunctionDefinition(function (params, sheet, rowIndex, columnIndex) {
                        var cellExpr;
                        if (params == null) {
                            return rowIndex + 1;
                        }
                        cellExpr = params[0];
                        cellExpr = self._ensureNonFunctionExpression(cellExpr);
                        if (cellExpr instanceof sheet_1._CellRangeExpression) {
                            return cellExpr.cells.row + 1;
                        }
                        throw 'Invalid Cell Reference.';
                    }, 1, 0);
                    self._functionTable['rows'] = new _FunctionDefinition(function (params, sheet) {
                        var cellExpr = params[0];
                        cellExpr = self._ensureNonFunctionExpression(cellExpr);
                        if (cellExpr instanceof sheet_1._CellRangeExpression) {
                            return cellExpr.cells.rowSpan;
                        }
                        throw 'Invalid Cell Reference.';
                    }, 1, 1);
                    self._functionTable['choose'] = new _FunctionDefinition(function (params, sheet) {
                        var index = sheet_1._Expression.toNumber(params[0], sheet);
                        if (isNaN(index)) {
                            throw 'Invalid index number.';
                        }
                        if (index < 1 || index >= params.length) {
                            throw 'The index number is out of the list range.';
                        }
                        return params[index].evaluate(sheet);
                    }, 255, 2);
                    self._functionTable['index'] = new _FunctionDefinition(function (params, sheet) {
                        var cellExpr = params[0], cells, rowNum = sheet_1._Expression.toNumber(params[1], sheet), colNum = params[2] != null ? sheet_1._Expression.toNumber(params[2], sheet) : 0;
                        if (isNaN(rowNum) || rowNum < 0) {
                            throw 'Invalid Row Number.';
                        }
                        if (isNaN(colNum) || colNum < 0) {
                            throw 'Invalid Column Number.';
                        }
                        cellExpr = self._ensureNonFunctionExpression(cellExpr);
                        if (cellExpr instanceof sheet_1._CellRangeExpression) {
                            cells = cellExpr.cells;
                            if (rowNum > cells.rowSpan || colNum > cells.columnSpan) {
                                throw 'Index is out of the cell range.';
                            }
                            if (rowNum > 0 && colNum > 0) {
                                return self._owner.getCellValue(cells.topRow + rowNum - 1, cells.leftCol + colNum - 1, true, sheet);
                            }
                            if (rowNum === 0 && colNum === 0) {
                                return cellExpr;
                            }
                            if (rowNum === 0) {
                                return new sheet_1._CellRangeExpression(new grid.CellRange(cells.topRow, cells.leftCol + colNum - 1, cells.bottomRow, cells.leftCol + colNum - 1), cellExpr.sheetRef, self._owner);
                            }
                            if (colNum === 0) {
                                return new sheet_1._CellRangeExpression(new grid.CellRange(cells.topRow + rowNum - 1, cells.leftCol, cells.topRow + rowNum - 1, cells.rightCol), cellExpr.sheetRef, self._owner);
                            }
                        }
                        throw 'Invalid Cell Reference.';
                    }, 4, 2);
                    self._functionTable['hlookup'] = new _FunctionDefinition(function (params, sheet) {
                        return self._handleHLookup(params, sheet);
                    }, 4, 3);
                };
                // Register the finacial function for the calcEngine.
                _CalcEngine.prototype._registFinacialFunction = function () {
                    var self = this;
                    self._functionTable['rate'] = new _FunctionDefinition(function (params, sheet) {
                        var rate = self._calculateRate(params, sheet);
                        return {
                            value: rate,
                            format: 'p2'
                        };
                    }, 6, 3);
                };
                // Add token into the static token table.
                _CalcEngine.prototype._addToken = function (symbol, id, type) {
                    var token = new _Token(symbol, id, type);
                    this._tokenTable[symbol] = token;
                };
                // Parse expression
                _CalcEngine.prototype._parseExpression = function () {
                    this._getToken();
                    return this._parseCompareOrConcat();
                };
                // Parse compare expression
                _CalcEngine.prototype._parseCompareOrConcat = function () {
                    var x = this._parseAddSub(), t, exprArg;
                    while (this._token.tokenType === _TokenType.COMPARE || this._token.tokenType === _TokenType.CONCAT) {
                        t = this._token;
                        this._getToken();
                        exprArg = this._parseAddSub();
                        x = new sheet_1._BinaryExpression(t, x, exprArg);
                    }
                    return x;
                };
                // Parse add/sub expression.
                _CalcEngine.prototype._parseAddSub = function () {
                    var x = this._parseMulDiv(), t, exprArg;
                    while (this._token.tokenType === _TokenType.ADDSUB) {
                        t = this._token;
                        this._getToken();
                        exprArg = this._parseMulDiv();
                        x = new sheet_1._BinaryExpression(t, x, exprArg);
                    }
                    return x;
                };
                // Parse multiple/division expression.
                _CalcEngine.prototype._parseMulDiv = function () {
                    var x = this._parsePower(), t, exprArg;
                    while (this._token.tokenType === _TokenType.MULDIV) {
                        t = this._token;
                        this._getToken();
                        exprArg = this._parsePower();
                        x = new sheet_1._BinaryExpression(t, x, exprArg);
                    }
                    return x;
                };
                // Parse power expression.
                _CalcEngine.prototype._parsePower = function () {
                    var x = this._parseUnary(), t, exprArg;
                    while (this._token.tokenType === _TokenType.POWER) {
                        t = this._token;
                        this._getToken();
                        exprArg = this._parseUnary();
                        x = new sheet_1._BinaryExpression(t, x, exprArg);
                    }
                    return x;
                };
                // Parse unary expression
                _CalcEngine.prototype._parseUnary = function () {
                    var t, exprArg;
                    // unary plus and minus
                    if (this._token.tokenID === _TokenID.ADD || this._token.tokenID === _TokenID.SUB) {
                        t = this._token;
                        this._getToken();
                        exprArg = this._parseAtom();
                        return new sheet_1._UnaryExpression(t, exprArg);
                    }
                    // not unary, return atom
                    return this._parseAtom();
                };
                // Parse atomic expression
                _CalcEngine.prototype._parseAtom = function () {
                    var x = null, id, funcDefinition, params, pCnt, cellRef;
                    switch (this._token.tokenType) {
                        // literals
                        case _TokenType.LITERAL:
                            x = new sheet_1._Expression(this._token);
                            break;
                        // identifiers
                        case _TokenType.IDENTIFIER:
                            // get identifier
                            id = this._token.value.toString();
                            funcDefinition = this._functionTable[id.toLowerCase()];
                            // look for functions
                            if (funcDefinition) {
                                params = this._getParameters();
                                pCnt = params ? params.length : 0;
                                if (funcDefinition.paramMin !== -1 && pCnt < funcDefinition.paramMin) {
                                    throw 'Too few parameters.';
                                }
                                if (funcDefinition.paramMax !== -1 && pCnt > funcDefinition.paramMax) {
                                    throw 'Too many parameters.';
                                }
                                x = new sheet_1._FunctionExpression(funcDefinition, params);
                                break;
                            }
                            // look for Cell Range.
                            cellRef = this._getCellRange(id);
                            if (cellRef) {
                                x = new sheet_1._CellRangeExpression(cellRef.cellRange, cellRef.sheetRef, this._owner);
                                break;
                            }
                            // trigger the unknownFunction event.
                            params = this._getParameters();
                            x = this.onUnknownFunction(id, params);
                            break;
                        // sub-expressions
                        case _TokenType.GROUP:
                            // anything other than opening parenthesis is illegal here
                            if (this._token.tokenID !== _TokenID.OPEN) {
                                throw 'Expression expected.';
                            }
                            // get expression
                            this._getToken();
                            x = this._parseCompareOrConcat();
                            // check that the parenthesis was closed
                            if (this._token.tokenID !== _TokenID.CLOSE) {
                                throw 'Unbalanced parenthesis.';
                            }
                            break;
                    }
                    // make sure we got something...
                    if (x === null) {
                        throw '';
                    }
                    // done
                    this._getToken();
                    return x;
                };
                // Get token for the expression.
                _CalcEngine.prototype._getToken = function () {
                    var i, c, lastChar, isLetter, isDigit, id = '', sheetRef = '', 
                    // About the Japanese characters checking
                    // Please refer http://stackoverflow.com/questions/15033196/using-javascript-to-check-whether-a-string-contains-japanese-characters-includi
                    // And http://www.rikai.com/library/kanjitables/kanji_codes.unicode.shtml
                    japaneseRegExp = new RegExp('[\u3000-\u303f\u3040-\u309f\u30a0-\u30ff\uff00-\uff9f\u4e00-\u9faf\u3400-\u4dbf]');
                    // eat white space 
                    while (this._pointer < this._expressLength && this._expression[this._pointer] === ' ') {
                        this._pointer++;
                    }
                    // are we done?
                    if (this._pointer >= this._expressLength) {
                        this._token = new _Token(null, _TokenID.END, _TokenType.GROUP);
                        return;
                    }
                    // prepare to parse
                    c = this._expression[this._pointer];
                    // operators
                    // this gets called a lot, so it's pretty optimized.
                    // note that operators must start with non-letter/digit characters.
                    isLetter = (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || japaneseRegExp.test(c);
                    isDigit = (c >= '0' && c <= '9') || c == '.';
                    if (!isLetter && !isDigit) {
                        var tk = this._tokenTable[c];
                        if (tk) {
                            // save token we found
                            this._token = tk;
                            this._pointer++;
                            // look for double-char tokens (special case)
                            if (this._pointer < this._expressLength && (c === '>' || c === '<')) {
                                tk = this._tokenTable[this._expression.substring(this._pointer - 1, this._pointer + 1)];
                                if (tk) {
                                    this._token = tk;
                                    this._pointer++;
                                }
                            }
                            return;
                        }
                    }
                    // parse numbers token
                    if (isDigit) {
                        this._parseDigit();
                        return;
                    }
                    // parse strings token
                    if (c === '\"') {
                        this._parseString();
                        return;
                    }
                    if (c === '\'') {
                        sheetRef = this._parseSheetRef();
                        if (!sheetRef) {
                            return;
                        }
                    }
                    // parse dates token
                    if (c === '#') {
                        this._parseDate();
                        return;
                    }
                    // identifiers (functions, objects) must start with alpha or underscore
                    if (!isLetter && c !== '_' && this._idChars.indexOf(c) < 0 && !sheetRef) {
                        throw 'Identifier expected.';
                    }
                    // and must contain only letters/digits/_idChars
                    for (i = 1; i + this._pointer < this._expressLength; i++) {
                        c = this._expression[this._pointer + i];
                        isLetter = (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || japaneseRegExp.test(c);
                        isDigit = c >= '0' && c <= '9';
                        if (c === '\'' && lastChar === ':') {
                            id = sheetRef + this._expression.substring(this._pointer, this._pointer + i);
                            this._pointer += i;
                            sheetRef = this._parseSheetRef();
                            i = 0;
                            continue;
                        }
                        lastChar = c;
                        if (!isLetter && !isDigit && c !== '_' && this._idChars.indexOf(c) < 0) {
                            break;
                        }
                    }
                    // got identifier
                    id += sheetRef + this._expression.substring(this._pointer, this._pointer + i);
                    this._pointer += i;
                    this._token = new _Token(id, _TokenID.ATOM, _TokenType.IDENTIFIER);
                };
                // Parse digit token
                _CalcEngine.prototype._parseDigit = function () {
                    var div = -1, sci = false, pct = false, val = 0.0, i, c, lit;
                    for (i = 0; i + this._pointer < this._expressLength; i++) {
                        c = this._expression[this._pointer + i];
                        // digits always OK
                        if (c >= '0' && c <= '9') {
                            val = val * 10 + (+c - 0);
                            if (div > -1) {
                                div *= 10;
                            }
                            continue;
                        }
                        // one decimal is OK
                        if (c === '.' && div < 0) {
                            div = 1;
                            continue;
                        }
                        // scientific notation?
                        if ((c === 'E' || c === 'e') && !sci) {
                            sci = true;
                            c = this._expression[this._pointer + i + 1];
                            if (c === '+' || c === '-')
                                i++;
                            continue;
                        }
                        // percentage?
                        if (c === '%') {
                            pct = true;
                            i++;
                            break;
                        }
                        // end of literal
                        break;
                    }
                    // end of number, get value
                    if (!sci) {
                        // much faster than ParseDouble
                        if (div > 1) {
                            val /= div;
                        }
                        if (pct) {
                            val /= 100.0;
                        }
                    }
                    else {
                        lit = this._expression.substring(this._pointer, this._pointer + i);
                        val = +lit;
                    }
                    // build token
                    this._token = new _Token(val, _TokenID.ATOM, _TokenType.LITERAL);
                    // advance pointer and return
                    this._pointer += i;
                };
                // Parse string token
                _CalcEngine.prototype._parseString = function () {
                    var i, c, cNext, lit;
                    // look for end quote, skip double quotes
                    for (i = 1; i + this._pointer < this._expressLength; i++) {
                        c = this._expression[this._pointer + i];
                        if (c !== '\"') {
                            continue;
                        }
                        cNext = i + this._pointer < this._expressLength - 1 ? this._expression[this._pointer + i + 1] : ' ';
                        if (cNext !== '\"') {
                            break;
                        }
                        i++;
                    }
                    // check that we got the end of the string
                    if (c !== '\"') {
                        throw 'Can\'t find final quote.';
                    }
                    // end of string
                    lit = this._expression.substring(this._pointer + 1, this._pointer + i);
                    this._pointer += i + 1;
                    if (this._expression[this._pointer] === '!') {
                        throw 'Illegal cross sheet reference.';
                    }
                    this._token = new _Token(lit.replace('\"\"', '\"'), _TokenID.ATOM, _TokenType.LITERAL);
                };
                // Parse datetime token
                _CalcEngine.prototype._parseDate = function () {
                    var i, c, lit;
                    // look for end #
                    for (i = 1; i + this._pointer < this._expressLength; i++) {
                        c = this._expression[this._pointer + i];
                        if (c === '#') {
                            break;
                        }
                    }
                    // check that we got the end of the date
                    if (c !== '#') {
                        throw 'Can\'t find final date delimiter ("#").';
                    }
                    // end of date
                    lit = this._expression.substring(this._pointer + 1, this._pointer + i);
                    this._pointer += i + 1;
                    this._token = new _Token(Date.parse(lit), _TokenID.ATOM, _TokenType.LITERAL);
                };
                // Parse the sheet reference.
                _CalcEngine.prototype._parseSheetRef = function () {
                    var i, c, cNext, lit;
                    // look for end quote, skip double quotes
                    for (i = 1; i + this._pointer < this._expressLength; i++) {
                        c = this._expression[this._pointer + i];
                        if (c !== '\'') {
                            continue;
                        }
                        cNext = i + this._pointer < this._expressLength - 1 ? this._expression[this._pointer + i + 1] : ' ';
                        if (cNext !== '\'') {
                            break;
                        }
                        i++;
                    }
                    // check that we got the end of the string
                    if (c !== '\'') {
                        throw 'Can\'t find final quote.';
                    }
                    // end of string
                    lit = this._expression.substring(this._pointer + 1, this._pointer + i);
                    this._pointer += i + 1;
                    if (this._expression[this._pointer] === '!') {
                        return lit.replace(/\'\'/g, '\'');
                    }
                    else {
                        return '';
                    }
                };
                // Gets the cell range by the identifier.
                // For e.g. A1:C3 to cellRange(row=0, col=0, row1=2, col1=2)
                _CalcEngine.prototype._getCellRange = function (identifier) {
                    var cells, cell, cell2, sheetRef, rng, rng2;
                    if (identifier) {
                        cells = identifier.split(':');
                        if (cells.length > 0 && cells.length < 3) {
                            cell = this._parseCell(cells[0]);
                            rng = cell.cellRange;
                            if (rng && cells.length === 2) {
                                cell2 = this._parseCell(cells[1]);
                                rng2 = cell2.cellRange;
                                if (cell.sheetRef && !cell2.sheetRef) {
                                    cell2.sheetRef = cell.sheetRef;
                                }
                                if (cell.sheetRef !== cell2.sheetRef) {
                                    throw 'The cell reference must be in the same sheet!';
                                }
                                if (rng2) {
                                    rng.col2 = rng2.col;
                                    rng.row2 = rng2.row;
                                }
                                else {
                                    rng = null;
                                }
                            }
                        }
                    }
                    if (rng == null) {
                        return null;
                    }
                    return {
                        cellRange: rng,
                        sheetRef: cell.sheetRef
                    };
                };
                // Parse the single string cell identifier to cell range;
                // For e.g. A1 to cellRange(row=0, col=0).
                _CalcEngine.prototype._parseCellRange = function (cell) {
                    var col = -1, row = -1, absCol = false, absRow = false, index, c;
                    // parse column
                    for (index = 0; index < cell.length; index++) {
                        c = cell[index];
                        if (c === '$' && !absCol) {
                            absCol = true;
                            continue;
                        }
                        if (!(c >= 'a' && c <= 'z') && !(c >= 'A' && c <= 'Z')) {
                            break;
                        }
                        if (col < 0) {
                            col = 0;
                        }
                        col = 26 * col + (c.toUpperCase().charCodeAt(0) - 'A'.charCodeAt(0) + 1);
                    }
                    // parse row
                    for (; index < cell.length; index++) {
                        c = cell[index];
                        if (c === '$' && !absRow) {
                            absRow = true;
                            continue;
                        }
                        if (!(c >= '0' && c <= '9')) {
                            break;
                        }
                        if (row < 0) {
                            row = 0;
                        }
                        row = 10 * row + (+c - 0);
                    }
                    // sanity
                    if (index < cell.length) {
                        row = col = -1;
                    }
                    if (row === -1 || col === -1) {
                        return null;
                    }
                    // done
                    return new grid.CellRange(row - 1, col - 1);
                };
                // Parse the single cell reference string to cell reference object.
                // For e.g. 'sheet1!A1' to { sheetRef: 'sheet1', cellRange: CellRange(row = 0, col = 0)}
                _CalcEngine.prototype._parseCell = function (cell) {
                    var rng, sheetRefIndex, cellsRef, sheetRef;
                    sheetRefIndex = cell.lastIndexOf('!');
                    if (sheetRefIndex > 0 && sheetRefIndex < cell.length - 1) {
                        sheetRef = cell.substring(0, sheetRefIndex);
                        cellsRef = cell.substring(sheetRefIndex + 1);
                    }
                    else if (sheetRefIndex <= 0) {
                        cellsRef = cell;
                    }
                    else {
                        return null;
                    }
                    rng = this._parseCellRange(cellsRef);
                    return {
                        cellRange: rng,
                        sheetRef: sheetRef
                    };
                };
                // Gets the parameters for the function.
                // e.g. myfun(a, b, c+2)
                _CalcEngine.prototype._getParameters = function () {
                    // check whether next token is a (, 
                    // restore state and bail if it's not
                    var pos = this._pointer, tk = this._token, parms, expr;
                    this._getToken();
                    if (this._token.tokenID !== _TokenID.OPEN) {
                        this._pointer = pos;
                        this._token = tk;
                        return null;
                    }
                    // check for empty Parameter list
                    pos = this._pointer;
                    this._getToken();
                    if (this._token.tokenID === _TokenID.CLOSE) {
                        return null;
                    }
                    this._pointer = pos;
                    // get Parameters until we reach the end of the list
                    parms = new Array();
                    expr = this._parseExpression();
                    parms.push(expr);
                    while (this._token.tokenID === _TokenID.COMMA) {
                        expr = this._parseExpression();
                        parms.push(expr);
                    }
                    // make sure the list was closed correctly
                    if (this._token.tokenID !== _TokenID.CLOSE) {
                        throw 'Syntax error.';
                    }
                    // done
                    return parms;
                };
                // Get the aggregate result for the CalcEngine.
                _CalcEngine.prototype._getAggregateResult = function (aggType, params, sheet) {
                    var list = this._getItemList(params, sheet), result;
                    result = wijmo.getAggregate(aggType, list.items);
                    if (list.isDate) {
                        result = new Date(result);
                    }
                    return result;
                };
                // Get the flexsheet aggregate result for the CalcEngine
                _CalcEngine.prototype._getFlexSheetAggregateResult = function (aggType, params, sheet) {
                    var list, sumList, num, order;
                    switch (aggType) {
                        case _FlexSheetAggregate.Count:
                            list = this._getItemList(params, sheet, true, false);
                            return this._countNumberCells(list.items);
                        case _FlexSheetAggregate.CountA:
                            list = this._getItemList(params, sheet, false, false);
                            return list.items.length;
                        case _FlexSheetAggregate.ConutBlank:
                            list = this._getItemList(params, sheet, false, true);
                            return this._countBlankCells(list.items);
                        case _FlexSheetAggregate.Rank:
                            num = sheet_1._Expression.toNumber(params[0], sheet);
                            order = params[2] ? sheet_1._Expression.toNumber(params[2], sheet) : 0;
                            if (isNaN(num)) {
                                throw 'Invalid number.';
                            }
                            if (isNaN(order)) {
                                throw 'Invalid order.';
                            }
                            params[1] = this._ensureNonFunctionExpression(params[1]);
                            if (params[1] instanceof sheet_1._CellRangeExpression) {
                                list = this._getItemList([params[1]], sheet);
                                return this._getRankOfCellRange(num, list.items, order);
                            }
                            throw 'Invalid Cell Reference.';
                        case _FlexSheetAggregate.CountIf:
                            params[0] = this._ensureNonFunctionExpression(params[0]);
                            if (params[0] instanceof sheet_1._CellRangeExpression) {
                                list = this._getItemList([params[0]], sheet, false);
                                return this._countCellsByCriterias([list.items], [params[1]], sheet);
                            }
                            throw 'Invalid Cell Reference.';
                        case _FlexSheetAggregate.CountIfs:
                            return this._handleCountIfs(params, sheet);
                        case _FlexSheetAggregate.SumIf:
                            params[0] = this._ensureNonFunctionExpression(params[0]);
                            if (params[0] instanceof sheet_1._CellRangeExpression) {
                                list = this._getItemList([params[0]], sheet, false);
                                params[2] = this._ensureNonFunctionExpression(params[2]);
                                if (params[2] != null && params[2] instanceof sheet_1._CellRangeExpression) {
                                    sumList = this._getItemList([params[2]], sheet);
                                }
                                return this._sumCellsByCriterias([list.items], [params[1]], sumList ? sumList.items : null, sheet);
                            }
                            throw 'Invalid Cell Reference.';
                        case _FlexSheetAggregate.SumIfs:
                            return this._handleSumIfs(params, sheet);
                        case _FlexSheetAggregate.Product:
                            list = this._getItemList(params, sheet);
                            return this._getProductOfNumbers(list.items);
                    }
                    throw 'Invalid aggregate type.';
                };
                // Get item list for aggregate processing.
                _CalcEngine.prototype._getItemList = function (params, sheet, needParseToNum, isGetEmptyValue, isGetHiddenValue, columnIndex) {
                    if (needParseToNum === void 0) { needParseToNum = true; }
                    if (isGetEmptyValue === void 0) { isGetEmptyValue = false; }
                    if (isGetHiddenValue === void 0) { isGetHiddenValue = true; }
                    var items = new Array(), item, index, cellIndex, cellValues, param, isDate = true;
                    for (index = 0; index < params.length; index++) {
                        param = params[index];
                        // When meets the CellRangeExpression, 
                        // we need set the value of the each cell in the cell range into the array to get the aggregate result.
                        param = this._ensureNonFunctionExpression(param);
                        if (param instanceof sheet_1._CellRangeExpression) {
                            cellValues = param.getValues(isGetHiddenValue, columnIndex, sheet);
                            cells: for (cellIndex = 0; cellIndex < cellValues.length; cellIndex++) {
                                item = cellValues[cellIndex];
                                if (!isGetEmptyValue && (item == null || item === '')) {
                                    continue cells;
                                }
                                isDate = isDate && (item instanceof Date);
                                item = needParseToNum ? +item : item;
                                items.push(item);
                            }
                        }
                        else {
                            item = param instanceof sheet_1._Expression ? param.evaluate(sheet) : param;
                            if (!isGetEmptyValue && (item == null || item === '')) {
                                continue;
                            }
                            isDate = isDate && (item instanceof Date);
                            item = needParseToNum ? +item : item;
                            items.push(item);
                        }
                    }
                    if (items.length === 0) {
                        isDate = false;
                    }
                    return {
                        isDate: isDate,
                        items: items
                    };
                };
                // Count blank cells
                _CalcEngine.prototype._countBlankCells = function (items) {
                    var i = 0, count = 0, item;
                    for (; i < items.length; i++) {
                        item = items[i];
                        if (item == null || (wijmo.isString(item) && item === '') || (wijmo.isNumber(item) && isNaN(item))) {
                            count++;
                        }
                    }
                    return count;
                };
                // Count number cells
                _CalcEngine.prototype._countNumberCells = function (items) {
                    var i = 0, count = 0, item;
                    for (; i < items.length; i++) {
                        item = items[i];
                        if (item != null && wijmo.isNumber(item) && !isNaN(item)) {
                            count++;
                        }
                    }
                    return count;
                };
                // Get the rank for the number in the cell range.
                _CalcEngine.prototype._getRankOfCellRange = function (num, items, order) {
                    if (order === void 0) { order = 0; }
                    var i = 0, rank = 0, item;
                    // Sort the items list
                    if (!order) {
                        items.sort(function (a, b) {
                            if (isNaN(a) || isNaN(b)) {
                                return 1;
                            }
                            return b - a;
                        });
                    }
                    else {
                        items.sort(function (a, b) {
                            if (isNaN(a) || isNaN(b)) {
                                return -1;
                            }
                            return a - b;
                        });
                    }
                    for (; i < items.length; i++) {
                        item = items[i];
                        if (isNaN(item)) {
                            continue;
                        }
                        rank++;
                        if (num === item) {
                            return rank;
                        }
                    }
                    throw num + ' is not in the cell range.';
                };
                // Handles the CountIfs function
                _CalcEngine.prototype._handleCountIfs = function (params, sheet) {
                    var i = 0, itemsList = [], critreiaList = [], list, cellExpr, rowCount, colCount;
                    if (params.length % 2 !== 0) {
                        throw 'Invalid params.';
                    }
                    for (; i < params.length / 2; i++) {
                        cellExpr = params[2 * i];
                        cellExpr = this._ensureNonFunctionExpression(cellExpr);
                        if (cellExpr instanceof sheet_1._CellRangeExpression) {
                            if (i === 0) {
                                if (cellExpr.cells) {
                                    rowCount = cellExpr.cells.rowSpan;
                                    colCount = cellExpr.cells.columnSpan;
                                }
                                else {
                                    throw 'Invalid Cell Reference.';
                                }
                            }
                            else {
                                if (!cellExpr.cells) {
                                    throw 'Invalid Cell Reference.';
                                }
                                else if (cellExpr.cells.rowSpan !== rowCount || cellExpr.cells.columnSpan !== colCount) {
                                    throw 'The row span and column span of each cell range has to be same with each other.';
                                }
                            }
                            list = this._getItemList([cellExpr], sheet, false);
                            itemsList[i] = list.items;
                            critreiaList[i] = params[2 * i + 1];
                        }
                        else {
                            throw 'Invalid Cell Reference.';
                        }
                    }
                    return this._countCellsByCriterias(itemsList, critreiaList, sheet);
                };
                // Count the cells that meet the criteria.
                _CalcEngine.prototype._countCellsByCriterias = function (itemsList, criterias, sheet, countItems) {
                    var i = 0, j = 0, count = 0, rangeLength = itemsList[0].length, parsedRightExprs = [], result, countItem, items, leftExpr, rightExpr;
                    for (; j < criterias.length; j++) {
                        rightExpr = sheet_1._Expression.toString(criterias[j], sheet);
                        if (rightExpr.length === 0) {
                            throw 'Invalid Criteria.';
                        }
                        if (rightExpr === '*') {
                            parsedRightExprs.push(rightExpr);
                        }
                        else {
                            parsedRightExprs.push(this._parseRightExpr(rightExpr));
                        }
                    }
                    for (; i < rangeLength; i++) {
                        result = false;
                        criteriaLoop: for (j = 0; j < itemsList.length; j++) {
                            items = itemsList[j];
                            leftExpr = items[i];
                            rightExpr = parsedRightExprs[j];
                            if (typeof rightExpr === 'string') {
                                if (rightExpr !== '*' && (leftExpr == null || leftExpr === '')) {
                                    result = false;
                                    break criteriaLoop;
                                }
                                result = rightExpr === '*' || this.evaluate(this._combineExpr(leftExpr, rightExpr), null, sheet);
                                if (!result) {
                                    break criteriaLoop;
                                }
                            }
                            else {
                                result = result = rightExpr.reg.test(leftExpr.toString()) === rightExpr.checkMathces;
                                if (!result) {
                                    break criteriaLoop;
                                }
                            }
                        }
                        if (result) {
                            if (countItems) {
                                countItem = countItems[i];
                                if (countItem != null && wijmo.isNumber(countItem) && !isNaN(countItem)) {
                                    count++;
                                }
                            }
                            else {
                                count++;
                            }
                        }
                    }
                    return count;
                };
                // Handles the SumIfs function
                _CalcEngine.prototype._handleSumIfs = function (params, sheet) {
                    var i = 1, itemsList = [], critreiaList = [], list, sumList, sumCellExpr, cellExpr, rowCount, colCount;
                    if (params.length % 2 !== 1) {
                        throw 'Invalid params.';
                    }
                    sumCellExpr = params[0];
                    sumCellExpr = this._ensureNonFunctionExpression(sumCellExpr);
                    if (sumCellExpr instanceof sheet_1._CellRangeExpression) {
                        if (sumCellExpr.cells) {
                            rowCount = sumCellExpr.cells.rowSpan;
                            colCount = sumCellExpr.cells.columnSpan;
                        }
                        else {
                            throw 'Invalid Sum Cell Reference.';
                        }
                        sumList = this._getItemList([sumCellExpr], sheet);
                    }
                    else {
                        throw 'Invalid Sum Cell Reference.';
                    }
                    for (; i < (params.length + 1) / 2; i++) {
                        cellExpr = params[2 * i - 1];
                        cellExpr = this._ensureNonFunctionExpression(cellExpr);
                        if (cellExpr instanceof sheet_1._CellRangeExpression) {
                            if (!cellExpr.cells) {
                                throw 'Invalid Criteria Cell Reference.';
                            }
                            else if (cellExpr.cells.rowSpan !== rowCount || cellExpr.cells.columnSpan !== colCount) {
                                throw 'The row span and column span of each cell range has to be same with each other.';
                            }
                            list = this._getItemList([cellExpr], sheet, false);
                            itemsList[i - 1] = list.items;
                            critreiaList[i - 1] = params[2 * i];
                        }
                        else {
                            throw 'Invalid Criteria Cell Reference.';
                        }
                    }
                    return this._sumCellsByCriterias(itemsList, critreiaList, sumList.items, sheet);
                };
                // Gets the sum of the numeric values in the cells specified by a given criteria.
                _CalcEngine.prototype._sumCellsByCriterias = function (itemsList, criterias, sumItems, sheet) {
                    var i = 0, j = 0, sum = 0, sumItem, rangeLength = itemsList[0].length, parsedRightExprs = [], result, items, leftExpr, rightExpr;
                    if (sumItems == null) {
                        sumItems = itemsList[0];
                    }
                    for (; j < criterias.length; j++) {
                        rightExpr = sheet_1._Expression.toString(criterias[j], sheet);
                        if (rightExpr.length === 0) {
                            throw 'Invalid Criteria.';
                        }
                        if (rightExpr === '*') {
                            parsedRightExprs.push(rightExpr);
                        }
                        else {
                            parsedRightExprs.push(this._parseRightExpr(rightExpr));
                        }
                    }
                    for (; i < rangeLength; i++) {
                        result = false;
                        sumItem = sumItems[i];
                        criteriaLoop: for (j = 0; j < itemsList.length; j++) {
                            items = itemsList[j];
                            leftExpr = items[i];
                            rightExpr = parsedRightExprs[j];
                            if (typeof rightExpr === 'string') {
                                if (rightExpr !== '*' && (leftExpr == null || leftExpr === '')) {
                                    result = false;
                                    break criteriaLoop;
                                }
                                result = rightExpr === '*' || this.evaluate(this._combineExpr(leftExpr, rightExpr), null, sheet);
                                if (!result) {
                                    break criteriaLoop;
                                }
                            }
                            else {
                                result = rightExpr.reg.test(leftExpr.toString()) === rightExpr.checkMathces;
                                if (!result) {
                                    break criteriaLoop;
                                }
                            }
                        }
                        if (result && wijmo.isNumber(sumItem) && !isNaN(sumItem)) {
                            sum += sumItem;
                        }
                    }
                    return sum;
                };
                // Get product for numbers
                _CalcEngine.prototype._getProductOfNumbers = function (items) {
                    var item, i = 0, product = 1, containsValidNum = false;
                    if (items) {
                        for (; i < items.length; i++) {
                            item = items[i];
                            if (wijmo.isNumber(item) && !isNaN(item)) {
                                product *= item;
                                containsValidNum = true;
                            }
                        }
                    }
                    if (containsValidNum) {
                        return product;
                    }
                    return 0;
                };
                //  Handle the subtotal function.
                _CalcEngine.prototype._handleSubtotal = function (params, sheet) {
                    var func, list, aggType, result, isGetHiddenValue = true;
                    func = sheet_1._Expression.toNumber(params[0], sheet);
                    if ((func >= 1 && func <= 11) || (func >= 101 && func <= 111)) {
                        if (func >= 101 && func <= 111) {
                            isGetHiddenValue = false;
                        }
                        func = wijmo.asEnum(func, _SubtotalFunction);
                        list = this._getItemList(params.slice(1), sheet, true, false, isGetHiddenValue);
                        switch (func) {
                            case _SubtotalFunction.Count:
                            case _SubtotalFunction.CountWithoutHidden:
                                return this._countNumberCells(list.items);
                            case _SubtotalFunction.CountA:
                            case _SubtotalFunction.CountAWithoutHidden:
                                return list.items.length;
                            case _SubtotalFunction.Product:
                            case _SubtotalFunction.ProductWithoutHidden:
                                return this._getProductOfNumbers(list.items);
                            case _SubtotalFunction.Average:
                            case _SubtotalFunction.AverageWithoutHidden:
                                aggType = wijmo.Aggregate.Avg;
                                break;
                            case _SubtotalFunction.Max:
                            case _SubtotalFunction.MaxWithoutHidden:
                                aggType = wijmo.Aggregate.Max;
                                break;
                            case _SubtotalFunction.Min:
                            case _SubtotalFunction.MinWithoutHidden:
                                aggType = wijmo.Aggregate.Min;
                                break;
                            case _SubtotalFunction.Std:
                            case _SubtotalFunction.StdWithoutHidden:
                                aggType = wijmo.Aggregate.Std;
                                break;
                            case _SubtotalFunction.StdPop:
                            case _SubtotalFunction.StdPopWithoutHidden:
                                aggType = wijmo.Aggregate.StdPop;
                                break;
                            case _SubtotalFunction.Sum:
                            case _SubtotalFunction.SumWithoutHidden:
                                aggType = wijmo.Aggregate.Sum;
                                break;
                            case _SubtotalFunction.Var:
                            case _SubtotalFunction.VarWithoutHidden:
                                aggType = wijmo.Aggregate.Var;
                                break;
                            case _SubtotalFunction.VarPop:
                            case _SubtotalFunction.VarPopWithoutHidden:
                                aggType = wijmo.Aggregate.VarPop;
                                break;
                        }
                        result = wijmo.getAggregate(aggType, list.items);
                        if (list.isDate) {
                            result = new Date(result);
                        }
                        return result;
                    }
                    throw 'Invalid Subtotal function.';
                };
                // Handle the DCount function.
                _CalcEngine.prototype._handleDCount = function (params, sheet) {
                    var cellExpr = params[0], criteriaCellExpr = params[2], count = 0, field, columnIndex, list;
                    cellExpr = this._ensureNonFunctionExpression(cellExpr);
                    criteriaCellExpr = this._ensureNonFunctionExpression(criteriaCellExpr);
                    if (cellExpr instanceof sheet_1._CellRangeExpression && criteriaCellExpr instanceof sheet_1._CellRangeExpression) {
                        field = params[1].evaluate(sheet);
                        columnIndex = this._getColumnIndexByField(cellExpr, field);
                        list = this._getItemList([cellExpr], sheet, true, false, true, columnIndex);
                        if (list.items && list.items.length > 1) {
                            return this._DCountWithCriterias(list.items.slice(1), cellExpr, criteriaCellExpr);
                        }
                    }
                    throw 'Invalid Count Cell Reference.';
                };
                // Counts the cells by the specified criterias.
                _CalcEngine.prototype._DCountWithCriterias = function (countItems, countRef, criteriaRef) {
                    var criteriaCells = criteriaRef.cells, count = 0, countSheet, criteriaSheet, fieldRowIndex, rowIndex, colIndex, criteriaColIndex, criteria, criteriaField, list, itemsList, criteriaList;
                    countSheet = this._getSheet(countRef.sheetRef);
                    criteriaSheet = this._getSheet(criteriaRef.sheetRef);
                    if (criteriaCells.rowSpan > 1) {
                        fieldRowIndex = criteriaCells.topRow;
                        for (rowIndex = criteriaCells.bottomRow; rowIndex > criteriaCells.topRow; rowIndex--) {
                            itemsList = [];
                            criteriaList = [];
                            for (colIndex = criteriaCells.leftCol; colIndex <= criteriaCells.rightCol; colIndex++) {
                                // Collects the criterias and related cell reference.
                                criteria = this._owner.getCellValue(rowIndex, colIndex, false, criteriaSheet);
                                if (criteria != null && criteria !== '') {
                                    criteriaList.push(new sheet_1._Expression(criteria));
                                    criteriaField = this._owner.getCellValue(fieldRowIndex, colIndex, false, criteriaSheet);
                                    criteriaColIndex = this._getColumnIndexByField(countRef, criteriaField);
                                    list = this._getItemList([countRef], countSheet, false, false, true, criteriaColIndex);
                                    if (list.items != null && list.items.length > 1) {
                                        itemsList.push(list.items.slice(1));
                                    }
                                    else {
                                        throw 'Invalid Count Cell Reference.';
                                    }
                                }
                            }
                            count += this._countCellsByCriterias(itemsList, criteriaList, countSheet, countItems);
                        }
                        return count;
                    }
                    throw 'Invalid Criteria Cell Reference.';
                };
                // Get column index of the count cell range by the field.
                _CalcEngine.prototype._getColumnIndexByField = function (cellExpr, field) {
                    var cells, sheet, columnIndex, value, rowIndex;
                    cells = cellExpr.cells;
                    rowIndex = cells.topRow;
                    if (rowIndex === -1) {
                        throw 'Invalid Count Cell Reference.';
                    }
                    if (wijmo.isInt(field) && !isNaN(field)) {
                        // if the field is integer, we consider the field it the column index of the count cell range.
                        if (field >= 1 && field <= cells.columnSpan) {
                            columnIndex = cells.leftCol + field - 1;
                            return columnIndex;
                        }
                    }
                    else {
                        sheet = this._getSheet(cellExpr.sheetRef);
                        for (columnIndex = cells.leftCol; columnIndex <= cells.rightCol; columnIndex++) {
                            value = this._owner.getCellValue(rowIndex, columnIndex, false, sheet);
                            field = wijmo.isString(field) ? field.toLowerCase() : field;
                            value = wijmo.isString(value) ? value.toLowerCase() : value;
                            if (field === value) {
                                return columnIndex;
                            }
                        }
                    }
                    throw 'Invalid field.';
                };
                // Gets the sheet by the sheetRef.
                _CalcEngine.prototype._getSheet = function (sheetRef) {
                    var i = 0, sheet;
                    if (sheetRef) {
                        for (; i < this._owner.sheets.length; i++) {
                            sheet = this._owner.sheets[i];
                            if (sheet.name === sheetRef) {
                                break;
                            }
                        }
                    }
                    return sheet;
                };
                // Parse the right expression for countif countifs sumif and sumifs function.
                _CalcEngine.prototype._parseRightExpr = function (rightExpr) {
                    var match, matchReg, checkMathces = false;
                    // Match the criteria that contains '?' such as '??match' and etc..
                    if (rightExpr.indexOf('?') > -1 || rightExpr.indexOf('*') > -1) {
                        match = rightExpr.match(/([\?\*]*)(\w+)([\?\*]*)(\w+)([\?\*]*)/);
                        if (match != null && match.length === 6) {
                            matchReg = new RegExp('^' + (match[1].length > 0 ? this._parseRegCriteria(match[1]) : '') + match[2]
                                + (match[3].length > 0 ? this._parseRegCriteria(match[3]) : '') + match[4]
                                + (match[5].length > 0 ? this._parseRegCriteria(match[5]) : '') + '$', 'i');
                        }
                        else {
                            throw 'Invalid Criteria.';
                        }
                        if (/^[<>=]/.test(rightExpr)) {
                            if (rightExpr.trim()[0] === '=') {
                                checkMathces = true;
                            }
                        }
                        else {
                            checkMathces = true;
                        }
                        return {
                            reg: matchReg,
                            checkMathces: checkMathces
                        };
                    }
                    else {
                        if (!isNaN(+rightExpr)) {
                            rightExpr = '=' + rightExpr;
                        }
                        else if (/^\w/.test(rightExpr)) {
                            rightExpr = '="' + rightExpr + '"';
                        }
                        else if (/^[<>=]{1,2}\s*-?\w+$/.test(rightExpr)) {
                            rightExpr = rightExpr.replace(/([<>=]{1,2})\s*(-?\w+)/, '$1"$2"');
                        }
                        else {
                            throw 'Invalid Criteria.';
                        }
                        return rightExpr;
                    }
                };
                // combine the left expression and right expression for countif countifs sumif and sumifs function.
                _CalcEngine.prototype._combineExpr = function (leftExpr, rightExpr) {
                    if (wijmo.isString(leftExpr)) {
                        leftExpr = '"' + leftExpr + '"';
                    }
                    leftExpr = '=' + leftExpr;
                    return leftExpr + rightExpr;
                };
                // Parse regex criteria for '?' and '*'
                _CalcEngine.prototype._parseRegCriteria = function (criteria) {
                    var i = 0, questionMarkCnt = 0, regString = '';
                    for (; i < criteria.length; i++) {
                        if (criteria[i] === '*') {
                            if (questionMarkCnt > 0) {
                                regString += '\\w{' + questionMarkCnt + '}';
                                questionMarkCnt = 0;
                            }
                            regString += '\\w*';
                        }
                        else if (criteria[i] === '?') {
                            questionMarkCnt++;
                        }
                    }
                    if (questionMarkCnt > 0) {
                        regString += '\\w{' + questionMarkCnt + '}';
                    }
                    return regString;
                };
                // Calculate the rate.
                // The algorithm of the rate calculation refers http://stackoverflow.com/questions/3198939/recreate-excel-rate-function-using-newtons-method
                _CalcEngine.prototype._calculateRate = function (params, sheet) {
                    var FINANCIAL_PRECISION = 0.0000001, FINANCIAL_MAX_ITERATIONS = 20, i = 0, x0 = 0, x1, rate, nper, pmt, pv, fv, type, guess, y, f, y0, y1;
                    nper = sheet_1._Expression.toNumber(params[0], sheet);
                    pmt = sheet_1._Expression.toNumber(params[1], sheet);
                    pv = sheet_1._Expression.toNumber(params[2], sheet);
                    fv = params[3] != null ? sheet_1._Expression.toNumber(params[3], sheet) : 0;
                    type = params[4] != null ? sheet_1._Expression.toNumber(params[4], sheet) : 0;
                    guess = params[5] != null ? sheet_1._Expression.toNumber(params[5], sheet) : 0.1;
                    rate = guess;
                    if (Math.abs(rate) < FINANCIAL_PRECISION) {
                        y = pv * (1 + nper * rate) + pmt * (1 + rate * type) * nper + fv;
                    }
                    else {
                        f = Math.exp(nper * Math.log(1 + rate));
                        y = pv * f + pmt * (1 / rate + type) * (f - 1) + fv;
                    }
                    y0 = pv + pmt * nper + fv;
                    y1 = pv * f + pmt * (1 / rate + type) * (f - 1) + fv;
                    // find root by secant method
                    x1 = rate;
                    while ((Math.abs(y0 - y1) > FINANCIAL_PRECISION) && (i < FINANCIAL_MAX_ITERATIONS)) {
                        rate = (y1 * x0 - y0 * x1) / (y1 - y0);
                        x0 = x1;
                        x1 = rate;
                        if (Math.abs(rate) < FINANCIAL_PRECISION) {
                            y = pv * (1 + nper * rate) + pmt * (1 + rate * type) * nper + fv;
                        }
                        else {
                            f = Math.exp(nper * Math.log(1 + rate));
                            y = pv * f + pmt * (1 / rate + type) * (f - 1) + fv;
                        }
                        y0 = y1;
                        y1 = y;
                        ++i;
                    }
                    if (Math.abs(y0 - y1) > FINANCIAL_PRECISION && i === FINANCIAL_MAX_ITERATIONS) {
                        throw 'It is not able to calculate the rate with current parameters.';
                    }
                    return rate;
                };
                // Handle the hlookup function.
                _CalcEngine.prototype._handleHLookup = function (params, sheet) {
                    var lookupVal = params[0].evaluate(sheet), cellExpr = params[1], rowNum = sheet_1._Expression.toNumber(params[2], sheet), approximateMatch = params[3] != null ? sheet_1._Expression.toBoolean(params[3], sheet) : true, cells, colNum;
                    if (lookupVal == null || lookupVal == '') {
                        throw 'Invalid lookup value.';
                    }
                    if (isNaN(rowNum) || rowNum < 0) {
                        throw 'Invalid row index.';
                    }
                    cellExpr = this._ensureNonFunctionExpression(cellExpr);
                    if (cellExpr instanceof sheet_1._CellRangeExpression) {
                        cells = cellExpr.cells;
                        if (rowNum > cells.rowSpan) {
                            throw 'Row index is out of the cell range.';
                        }
                        if (approximateMatch) {
                            colNum = this._exactMatch(lookupVal, cells, sheet, false);
                            if (colNum === -1) {
                                colNum = this._approximateMatch(lookupVal, cells, sheet);
                            }
                        }
                        else {
                            colNum = this._exactMatch(lookupVal, cells, sheet);
                        }
                        if (colNum === -1) {
                            throw 'Lookup Value is not found.';
                        }
                        return this._owner.getCellValue(cells.topRow + rowNum - 1, colNum, false, sheet);
                    }
                    throw 'Invalid Cell Reference.';
                };
                // Handle the exact match for the hlookup.
                _CalcEngine.prototype._exactMatch = function (lookupValue, cells, sheet, needHandleWildCard) {
                    if (needHandleWildCard === void 0) { needHandleWildCard = true; }
                    var rowIndex = cells.topRow, colIndex, value, match, matchReg;
                    if (wijmo.isString(lookupValue)) {
                        lookupValue = lookupValue.toLowerCase();
                    }
                    // handle the wildcard question mark (?) and asterisk (*) for the lookup value.
                    if (needHandleWildCard && wijmo.isString(lookupValue) && (lookupValue.indexOf('?') > -1 || lookupValue.indexOf('*') > -1)) {
                        match = lookupValue.match(/([\?\*]*)(\w+)([\?\*]*)(\w+)([\?\*]*)/);
                        if (match != null && match.length === 6) {
                            matchReg = new RegExp('^' + (match[1].length > 0 ? this._parseRegCriteria(match[1]) : '') + match[2]
                                + (match[3].length > 0 ? this._parseRegCriteria(match[3]) : '') + match[4]
                                + (match[5].length > 0 ? this._parseRegCriteria(match[5]) : '') + '$', 'i');
                        }
                        else {
                            throw 'Invalid lookup value.';
                        }
                    }
                    for (colIndex = cells.leftCol; colIndex <= cells.rightCol; colIndex++) {
                        value = this._owner.getCellValue(rowIndex, colIndex, false, sheet);
                        if (matchReg != null) {
                            if (matchReg.test(value)) {
                                return colIndex;
                            }
                        }
                        else {
                            if (wijmo.isString(value)) {
                                value = value.toLowerCase();
                            }
                            if (lookupValue === value) {
                                return colIndex;
                            }
                        }
                    }
                    return -1;
                };
                // Handle the approximate match for the hlookup.
                _CalcEngine.prototype._approximateMatch = function (lookupValue, cells, sheet) {
                    var val, colIndex, rowIndex = cells.topRow, cellValues = [], i = 0;
                    if (wijmo.isString(lookupValue)) {
                        lookupValue = lookupValue.toLowerCase();
                    }
                    for (colIndex = cells.leftCol; colIndex <= cells.rightCol; colIndex++) {
                        val = this._owner.getCellValue(rowIndex, colIndex, false, sheet);
                        val = isNaN(+val) ? val : +val;
                        cellValues.push({ value: val, index: colIndex });
                    }
                    // Sort the cellValues array with descent order.
                    cellValues.sort(function (a, b) {
                        if (wijmo.isString(a.value)) {
                            a.value = a.value.toLowerCase();
                        }
                        if (wijmo.isString(b.value)) {
                            b.value = b.value.toLowerCase();
                        }
                        if (a.value > b.value) {
                            return -1;
                        }
                        else if (a.value === b.value) {
                            return b.index - a.index;
                        }
                        return 1;
                    });
                    for (; i < cellValues.length; i++) {
                        val = cellValues[i];
                        if (wijmo.isString(val.value)) {
                            val.value = val.value.toLowerCase();
                        }
                        // return the column index of the first value that less than lookup value.
                        if (lookupValue > val.value) {
                            return val.index;
                        }
                    }
                    throw 'Lookup Value is not found.';
                };
                // Check the expression cache.
                _CalcEngine.prototype._checkCache = function (expression) {
                    var expr = this._expressionCache[expression];
                    if (expr) {
                        return expr;
                    }
                    expr = this._parse(expression);
                    // when the size of the expression cache is greater than 10000,
                    // We will release the expression cache.
                    if (this._cacheSize > 10000) {
                        this._clearExpressionCache();
                    }
                    this._expressionCache[expression] = expr;
                    this._cacheSize++;
                    return expr;
                };
                // Ensure current is not function expression.
                _CalcEngine.prototype._ensureNonFunctionExpression = function (expr, sheet) {
                    while (expr instanceof sheet_1._FunctionExpression) {
                        expr = expr.evaluate(sheet);
                    }
                    return expr;
                };
                return _CalcEngine;
            }());
            sheet_1._CalcEngine = _CalcEngine;
            /*
             * Defines the Token class.
             *
             * It assists the expression instance to evaluate value.
             */
            var _Token = (function () {
                /*
                 * Initializes a new instance of the @see:Token class.
                 *
                 * @param val The value of the token.
                 * @param tkID The @see:TokenID value of the token.
                 * @param tkType The @see:TokenType value of the token.
                 */
                function _Token(val, tkID, tkType) {
                    this._value = val;
                    this._tokenID = tkID;
                    this._tokenType = tkType;
                }
                Object.defineProperty(_Token.prototype, "value", {
                    /*
                     * Gets the value of the token instance.
                     */
                    get: function () {
                        return this._value;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(_Token.prototype, "tokenID", {
                    /*
                     * Gets the token ID of the token instance.
                     */
                    get: function () {
                        return this._tokenID;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(_Token.prototype, "tokenType", {
                    /*
                     * Gets the token type of the token instance.
                     */
                    get: function () {
                        return this._tokenType;
                    },
                    enumerable: true,
                    configurable: true
                });
                return _Token;
            }());
            sheet_1._Token = _Token;
            /*
             * Function definition class (keeps function name, parameter counts, and function).
             */
            var _FunctionDefinition = (function () {
                /*
                 * Initializes a new instance of the @see:FunctionDefinition class.
                 *
                 * @param func The function will be invoked by the CalcEngine.
                 * @param paramMax The maximum count of the parameter that the function need.
                 * @param paramMin The minimum count of the parameter that the function need.
                 */
                function _FunctionDefinition(func, paramMax, paramMin) {
                    this._paramMax = Number.MAX_VALUE;
                    this._paramMin = Number.MIN_VALUE;
                    this._func = func;
                    if (wijmo.isNumber(paramMax) && !isNaN(paramMax)) {
                        this._paramMax = paramMax;
                    }
                    if (wijmo.isNumber(paramMin) && !isNaN(paramMin)) {
                        this._paramMin = paramMin;
                    }
                }
                Object.defineProperty(_FunctionDefinition.prototype, "paramMax", {
                    /*
                     * Gets the paramMax of the FunctionDefinition instance.
                     */
                    get: function () {
                        return this._paramMax;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(_FunctionDefinition.prototype, "paramMin", {
                    /*
                     * Gets the paramMin of the FunctionDefinition instance.
                     */
                    get: function () {
                        return this._paramMin;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(_FunctionDefinition.prototype, "func", {
                    /*
                     * Gets the func of the FunctionDefinition instance.
                     */
                    get: function () {
                        return this._func;
                    },
                    enumerable: true,
                    configurable: true
                });
                return _FunctionDefinition;
            }());
            sheet_1._FunctionDefinition = _FunctionDefinition;
            /*
             * Token types (used when building expressions, sequence defines operator priority)
             */
            (function (_TokenType) {
                /*
                 * This token type includes '<', '>', '=', '<=', '>=' and '<>'.
                 */
                _TokenType[_TokenType["COMPARE"] = 0] = "COMPARE";
                /*
                 * This token type includes '+' and '-'.
                 */
                _TokenType[_TokenType["ADDSUB"] = 1] = "ADDSUB";
                /*
                 * This token type includes '*' and '/'.
                 */
                _TokenType[_TokenType["MULDIV"] = 2] = "MULDIV";
                /*
                 * This token type includes '^'.
                 */
                _TokenType[_TokenType["POWER"] = 3] = "POWER";
                /*
                 * This token type includes '&'.
                 */
                _TokenType[_TokenType["CONCAT"] = 4] = "CONCAT";
                /*
                 * This token type includes '(' and ')'.
                 */
                _TokenType[_TokenType["GROUP"] = 5] = "GROUP";
                /*
                 * This token type includes number value, string value and etc..
                 */
                _TokenType[_TokenType["LITERAL"] = 6] = "LITERAL";
                /*
                 * This token type includes function.
                 */
                _TokenType[_TokenType["IDENTIFIER"] = 7] = "IDENTIFIER";
            })(sheet_1._TokenType || (sheet_1._TokenType = {}));
            var _TokenType = sheet_1._TokenType;
            /*
             * Token ID (used when evaluating expressions)
             */
            (function (_TokenID) {
                /*
                 * Greater than.
                 */
                _TokenID[_TokenID["GT"] = 0] = "GT";
                /*
                 * Less than.
                 */
                _TokenID[_TokenID["LT"] = 1] = "LT";
                /*
                 * Greater than or equal to.
                 */
                _TokenID[_TokenID["GE"] = 2] = "GE";
                /*
                 * Less than or equal to.
                 */
                _TokenID[_TokenID["LE"] = 3] = "LE";
                /*
                 * Equal to.
                 */
                _TokenID[_TokenID["EQ"] = 4] = "EQ";
                /*
                 * Not equal to.
                 */
                _TokenID[_TokenID["NE"] = 5] = "NE";
                /*
                 * Addition.
                 */
                _TokenID[_TokenID["ADD"] = 6] = "ADD";
                /*
                 * Subtraction.
                 */
                _TokenID[_TokenID["SUB"] = 7] = "SUB";
                /*
                 * Multiplication.
                 */
                _TokenID[_TokenID["MUL"] = 8] = "MUL";
                /*
                 * Division.
                 */
                _TokenID[_TokenID["DIV"] = 9] = "DIV";
                /*
                 * Gets quotient of division.
                 */
                _TokenID[_TokenID["DIVINT"] = 10] = "DIVINT";
                /*
                 * Gets remainder of division.
                 */
                _TokenID[_TokenID["MOD"] = 11] = "MOD";
                /*
                 * Power.
                 */
                _TokenID[_TokenID["POWER"] = 12] = "POWER";
                /*
                 * String concat.
                 */
                _TokenID[_TokenID["CONCAT"] = 13] = "CONCAT";
                /*
                 * Opening bracket.
                 */
                _TokenID[_TokenID["OPEN"] = 14] = "OPEN";
                /*
                 * Closing bracket.
                 */
                _TokenID[_TokenID["CLOSE"] = 15] = "CLOSE";
                /*
                 * Group end.
                 */
                _TokenID[_TokenID["END"] = 16] = "END";
                /*
                 * Comma.
                 */
                _TokenID[_TokenID["COMMA"] = 17] = "COMMA";
                /*
                 * Period.
                 */
                _TokenID[_TokenID["PERIOD"] = 18] = "PERIOD";
                /*
                 * Literal token
                 */
                _TokenID[_TokenID["ATOM"] = 19] = "ATOM";
            })(sheet_1._TokenID || (sheet_1._TokenID = {}));
            var _TokenID = sheet_1._TokenID;
            /*
             * Specifies the type of aggregate for flexsheet.
             */
            var _FlexSheetAggregate;
            (function (_FlexSheetAggregate) {
                /*
                 * Counts the number of cells that contain numbers, and counts numbers within the list of arguments.
                 */
                _FlexSheetAggregate[_FlexSheetAggregate["Count"] = 0] = "Count";
                /*
                 * Returns the number of cells that are not empty in a range.
                 */
                _FlexSheetAggregate[_FlexSheetAggregate["CountA"] = 1] = "CountA";
                /*
                 * Returns the number of empty cells in a specified range of cells.
                 */
                _FlexSheetAggregate[_FlexSheetAggregate["ConutBlank"] = 2] = "ConutBlank";
                /*
                 * Returns the number of the cells that meet the criteria you specify in the argument.
                 */
                _FlexSheetAggregate[_FlexSheetAggregate["CountIf"] = 3] = "CountIf";
                /*
                 * Returns the number of the cells that meet multiple criteria.
                 */
                _FlexSheetAggregate[_FlexSheetAggregate["CountIfs"] = 4] = "CountIfs";
                /*
                 * Returns the rank of a number in a list of numbers.
                 */
                _FlexSheetAggregate[_FlexSheetAggregate["Rank"] = 5] = "Rank";
                /*
                 * Returns the sum of the numeric values in the cells specified by a given criteria.
                 */
                _FlexSheetAggregate[_FlexSheetAggregate["SumIf"] = 6] = "SumIf";
                /*
                 * Returns the sum of the numeric values in the cells specified by a multiple criteria.
                 */
                _FlexSheetAggregate[_FlexSheetAggregate["SumIfs"] = 7] = "SumIfs";
                /*
                 * Multiplies all the numbers given as arguments and returns the product.
                 */
                _FlexSheetAggregate[_FlexSheetAggregate["Product"] = 8] = "Product";
            })(_FlexSheetAggregate || (_FlexSheetAggregate = {}));
            /*
             * Specifies the type of subtotal f to calculate over a group of values.
             */
            var _SubtotalFunction;
            (function (_SubtotalFunction) {
                /*
                 * Returns the average value of the numeric values in the group.
                 */
                _SubtotalFunction[_SubtotalFunction["Average"] = 1] = "Average";
                /*
                 * Counts the number of cells that contain numbers, and counts numbers within the list of arguments.
                 */
                _SubtotalFunction[_SubtotalFunction["Count"] = 2] = "Count";
                /*
                 * Counts the number of cells that are not empty in a range.
                 */
                _SubtotalFunction[_SubtotalFunction["CountA"] = 3] = "CountA";
                /*
                 * Returns the maximum value in the group.
                 */
                _SubtotalFunction[_SubtotalFunction["Max"] = 4] = "Max";
                /*
                 * Returns the minimum value in the group.
                 */
                _SubtotalFunction[_SubtotalFunction["Min"] = 5] = "Min";
                /*
                 * Multiplies all the numbers given as arguments and returns the product.
                 */
                _SubtotalFunction[_SubtotalFunction["Product"] = 6] = "Product";
                /*
                 *Returns the sample standard deviation of the numeric values in the group
                 * (uses the formula based on n-1).
                 */
                _SubtotalFunction[_SubtotalFunction["Std"] = 7] = "Std";
                /*
                 *Returns the population standard deviation of the values in the group
                 * (uses the formula based on n).
                 */
                _SubtotalFunction[_SubtotalFunction["StdPop"] = 8] = "StdPop";
                /*
                 * Returns the sum of the numeric values in the group.
                 */
                _SubtotalFunction[_SubtotalFunction["Sum"] = 9] = "Sum";
                /*
                 * Returns the sample variance of the numeric values in the group
                 * (uses the formula based on n-1).
                 */
                _SubtotalFunction[_SubtotalFunction["Var"] = 10] = "Var";
                /*
                 * Returns the population variance of the values in the group
                 * (uses the formula based on n).
                 */
                _SubtotalFunction[_SubtotalFunction["VarPop"] = 11] = "VarPop";
                /*
                 * Returns the average value of the numeric values in the group and ignores the hidden rows and columns.
                 */
                _SubtotalFunction[_SubtotalFunction["AverageWithoutHidden"] = 101] = "AverageWithoutHidden";
                /*
                 * Counts the number of cells that contain numbers, and counts numbers within the list of arguments and ignores the hidden rows and columns.
                 */
                _SubtotalFunction[_SubtotalFunction["CountWithoutHidden"] = 102] = "CountWithoutHidden";
                /*
                 * Counts the number of cells that are not empty in a range and ignores the hidden rows and columns.
                 */
                _SubtotalFunction[_SubtotalFunction["CountAWithoutHidden"] = 103] = "CountAWithoutHidden";
                /*
                 * Returns the maximum value in the group and ignores the hidden rows and columns.
                 */
                _SubtotalFunction[_SubtotalFunction["MaxWithoutHidden"] = 104] = "MaxWithoutHidden";
                /*
                 * Multiplies all the numbers given as arguments and returns the product and ignores the hidden rows and columns.
                 */
                _SubtotalFunction[_SubtotalFunction["MinWithoutHidden"] = 105] = "MinWithoutHidden";
                /*
                 * Multiplies all the numbers given as arguments and returns the product and ignores the hidden rows and columns.
                 */
                _SubtotalFunction[_SubtotalFunction["ProductWithoutHidden"] = 106] = "ProductWithoutHidden";
                /*
                 *Returns the sample standard deviation of the numeric values in the group
                 * (uses the formula based on n-1) and ignores the hidden rows and columns.
                 */
                _SubtotalFunction[_SubtotalFunction["StdWithoutHidden"] = 107] = "StdWithoutHidden";
                /*
                 *Returns the population standard deviation of the values in the group
                 * (uses the formula based on n) and ignores the hidden rows and columns.
                 */
                _SubtotalFunction[_SubtotalFunction["StdPopWithoutHidden"] = 108] = "StdPopWithoutHidden";
                /*
                 * Returns the sum of the numeric values in the group and ignores the hidden rows and columns.
                 */
                _SubtotalFunction[_SubtotalFunction["SumWithoutHidden"] = 109] = "SumWithoutHidden";
                /*
                 * Returns the sample variance of the numeric values in the group
                 * (uses the formula based on n-1) and ignores the hidden rows and columns.
                 */
                _SubtotalFunction[_SubtotalFunction["VarWithoutHidden"] = 110] = "VarWithoutHidden";
                /*
                 * Returns the population variance of the values in the group
                 * (uses the formula based on n) and ignores the hidden rows and columns.
                 */
                _SubtotalFunction[_SubtotalFunction["VarPopWithoutHidden"] = 111] = "VarPopWithoutHidden";
            })(_SubtotalFunction || (_SubtotalFunction = {}));
        })(sheet = grid.sheet || (grid.sheet = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_CalcEngine.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var sheet;
        (function (sheet_1) {
            'use strict';
            /*
             * Defines the base class that represents parsed expressions.
             */
            var _Expression = (function () {
                /*
                 * Initializes a new instance of the @see:Expression class.
                 *
                 * @param arg This parameter is used to build the token for the expression.
                 */
                function _Expression(arg) {
                    if (arg) {
                        if (arg instanceof sheet_1._Token) {
                            this._token = arg;
                        }
                        else {
                            this._token = new sheet_1._Token(arg, sheet_1._TokenID.ATOM, sheet_1._TokenType.LITERAL);
                        }
                    }
                    else {
                        this._token = new sheet_1._Token(null, sheet_1._TokenID.ATOM, sheet_1._TokenType.IDENTIFIER);
                    }
                }
                Object.defineProperty(_Expression.prototype, "token", {
                    /*
                     * Gets the token of the expression.
                     */
                    get: function () {
                        return this._token;
                    },
                    enumerable: true,
                    configurable: true
                });
                /*
                 * Evaluates the expression.
                 *
                 * @param sheet The @see:Sheet is referenced by the @see:Expression.
                 * @param rowIndex The row index of the cell where the expression located in.
                 * @param columnIndex The column index of the cell where the expression located in.
                 */
                _Expression.prototype.evaluate = function (sheet, rowIndex, columnIndex) {
                    if (this._token.tokenType !== sheet_1._TokenType.LITERAL) {
                        throw 'Bad expression.';
                    }
                    return this._token.value;
                };
                /*
                 * Parse the expression to a string value.
                 *
                 * @param x The @see:Expression need be parsed to string value.
                 * @param sheet The @see:Sheet is referenced by the @see:Expression.
                 */
                _Expression.toString = function (x, sheet) {
                    var v = x.evaluate(sheet);
                    if (!wijmo.isPrimitive(v)) {
                        v = v.value;
                    }
                    return v != null ? v.toString() : '';
                };
                /*
                 * Parse the expression to a number value.
                 *
                 * @param x The @see:Expression need be parsed to number value.
                 * @param sheet The @see:Sheet is referenced by the @see:Expression.
                 */
                _Expression.toNumber = function (x, sheet) {
                    // evaluate
                    var v = x.evaluate(sheet);
                    if (!wijmo.isPrimitive(v)) {
                        v = v.value;
                    }
                    // handle numbers
                    if (wijmo.isNumber(v)) {
                        return v;
                    }
                    // handle booleans
                    if (wijmo.isBoolean(v)) {
                        return v ? 1 : 0;
                    }
                    // handle dates
                    if (wijmo.isDate(v)) {
                        return this._toOADate(v);
                    }
                    // handle strings
                    if (wijmo.isString(v)) {
                        if (v) {
                            return +v;
                        }
                        else {
                            return 0;
                        }
                    }
                    // handle everything else
                    return wijmo.changeType(v, wijmo.DataType.Number, '');
                };
                /*
                 * Parse the expression to a boolean value.
                 *
                 * @param x The @see:Expression need be parsed to boolean value.
                 * @param sheet The @see:Sheet is referenced by the @see:Expression.
                 */
                _Expression.toBoolean = function (x, sheet) {
                    // evaluate
                    var v = x.evaluate(sheet);
                    if (!wijmo.isPrimitive(v)) {
                        v = v.value;
                    }
                    // handle booleans
                    if (wijmo.isBoolean(v)) {
                        return v;
                    }
                    // handle numbers
                    if (wijmo.isNumber(v)) {
                        return v === 0 ? false : true;
                    }
                    // handle everything else
                    return wijmo.changeType(v, wijmo.DataType.Boolean, '');
                };
                /*
                 * Parse the expression to a date value.
                 *
                 * @param x The @see:Expression need be parsed to date value.
                 * @param sheet The @see:Sheet is referenced by the @see:Expression.
                 */
                _Expression.toDate = function (x, sheet) {
                    // evaluate
                    var v = x.evaluate(sheet);
                    if (!wijmo.isPrimitive(v)) {
                        v = v.value;
                    }
                    // handle dates
                    if (wijmo.isDate(v)) {
                        return v;
                    }
                    // handle numbers
                    if (wijmo.isNumber(v)) {
                        return this._fromOADate(v);
                    }
                    // handle everything else
                    return wijmo.changeType(v, wijmo.DataType.Date, '');
                };
                // convert the common date to OLE Automation date.
                _Expression._toOADate = function (val) {
                    var epoch = Date.UTC(1899, 11, 30), // 1899-12-30T00:00:00
                    currentUTC = Date.UTC(val.getFullYear(), val.getMonth(), val.getDate(), val.getHours(), val.getMinutes(), val.getSeconds(), val.getMilliseconds());
                    return (currentUTC - epoch) / 8.64e7;
                };
                // convert the OLE Automation date to common date.
                _Expression._fromOADate = function (oADate) {
                    var epoch = Date.UTC(1899, 11, 30);
                    return new Date(oADate * 8.64e7 + epoch);
                };
                return _Expression;
            }());
            sheet_1._Expression = _Expression;
            /*
             * Defines the unary expression class.
             * For e.g. -1.23.
             */
            var _UnaryExpression = (function (_super) {
                __extends(_UnaryExpression, _super);
                /*
                 * Initializes a new instance of the @see:UnaryExpression class.
                 *
                 * @param arg This parameter is used to build the token for the expression.
                 * @param expr The @see:Expression instance for evaluating the UnaryExpression.
                 */
                function _UnaryExpression(arg, expr) {
                    _super.call(this, arg);
                    this._expr = expr;
                }
                /*
                 * Overrides the evaluate function of base class.
                 *
                 * @param sheet The @see:Sheet is referenced by the @see:Expression.
                 */
                _UnaryExpression.prototype.evaluate = function (sheet) {
                    if (this.token.tokenID === sheet_1._TokenID.SUB) {
                        if (this._evaluatedValue == null) {
                            this._evaluatedValue = -_Expression.toNumber(this._expr, sheet);
                        }
                        return this._evaluatedValue;
                    }
                    if (this.token.tokenID === sheet_1._TokenID.ADD) {
                        if (this._evaluatedValue == null) {
                            this._evaluatedValue = +_Expression.toNumber(this._expr, sheet);
                        }
                        return this._evaluatedValue;
                    }
                    throw 'Bad expression.';
                };
                return _UnaryExpression;
            }(_Expression));
            sheet_1._UnaryExpression = _UnaryExpression;
            /*
             * Defines the binary expression class.
             * For e.g. 1 + 1.
             */
            var _BinaryExpression = (function (_super) {
                __extends(_BinaryExpression, _super);
                /*
                 * Initializes a new instance of the @see:BinaryExpression class.
                 *
                 * @param arg This parameter is used to build the token for the expression.
                 * @param leftExpr The @see:Expression instance for evaluating the BinaryExpression.
                 * @param rightExpr The @see:Expression instance for evaluating the BinaryExpression.
                 */
                function _BinaryExpression(arg, leftExpr, rightExpr) {
                    _super.call(this, arg);
                    this._leftExpr = leftExpr;
                    this._rightExpr = rightExpr;
                }
                /*
                 * Overrides the evaluate function of base class.
                 *
                 * @param sheet The @see:Sheet is referenced by the @see:Expression.
                 */
                _BinaryExpression.prototype.evaluate = function (sheet) {
                    var strLeftVal, strRightVal, leftValue, rightValue, compareVal;
                    if (this._evaluatedValue != null) {
                        return this._evaluatedValue;
                    }
                    strLeftVal = _Expression.toString(this._leftExpr, sheet);
                    strRightVal = _Expression.toString(this._rightExpr, sheet);
                    if (this.token.tokenType === sheet_1._TokenType.CONCAT) {
                        this._evaluatedValue = strLeftVal + strRightVal;
                        return this._evaluatedValue;
                    }
                    leftValue = _Expression.toNumber(this._leftExpr, sheet);
                    rightValue = _Expression.toNumber(this._rightExpr, sheet);
                    compareVal = leftValue - rightValue;
                    // handle comparisons
                    if (this.token.tokenType === sheet_1._TokenType.COMPARE) {
                        switch (this.token.tokenID) {
                            case sheet_1._TokenID.GT: return compareVal > 0;
                            case sheet_1._TokenID.LT: return compareVal < 0;
                            case sheet_1._TokenID.GE: return compareVal >= 0;
                            case sheet_1._TokenID.LE: return compareVal <= 0;
                            case sheet_1._TokenID.EQ:
                                if (isNaN(compareVal)) {
                                    this._evaluatedValue = strLeftVal.toLowerCase() === strRightVal.toLowerCase();
                                    return this._evaluatedValue;
                                }
                                else {
                                    this._evaluatedValue = compareVal === 0;
                                    return this._evaluatedValue;
                                }
                            case sheet_1._TokenID.NE:
                                if (isNaN(compareVal)) {
                                    this._evaluatedValue = strLeftVal.toLowerCase() !== strRightVal.toLowerCase();
                                    return this._evaluatedValue;
                                }
                                else {
                                    this._evaluatedValue = compareVal !== 0;
                                    return this._evaluatedValue;
                                }
                        }
                    }
                    // handle everything else
                    switch (this.token.tokenID) {
                        case sheet_1._TokenID.ADD:
                            this._evaluatedValue = leftValue + rightValue;
                            break;
                        case sheet_1._TokenID.SUB:
                            this._evaluatedValue = leftValue - rightValue;
                            break;
                        case sheet_1._TokenID.MUL:
                            this._evaluatedValue = leftValue * rightValue;
                            break;
                        case sheet_1._TokenID.DIV:
                            this._evaluatedValue = leftValue / rightValue;
                            break;
                        case sheet_1._TokenID.DIVINT:
                            this._evaluatedValue = Math.floor(leftValue / rightValue);
                            break;
                        case sheet_1._TokenID.MOD:
                            this._evaluatedValue = Math.floor(leftValue % rightValue);
                            break;
                        case sheet_1._TokenID.POWER:
                            if (rightValue === 0.0) {
                                this._evaluatedValue = 1.0;
                            }
                            if (rightValue === 0.5) {
                                this._evaluatedValue = Math.sqrt(leftValue);
                            }
                            if (rightValue === 1.0) {
                                this._evaluatedValue = leftValue;
                            }
                            if (rightValue === 2.0) {
                                this._evaluatedValue = leftValue * leftValue;
                            }
                            if (rightValue === 3.0) {
                                this._evaluatedValue = leftValue * leftValue * leftValue;
                            }
                            if (rightValue === 4.0) {
                                this._evaluatedValue = leftValue * leftValue * leftValue * leftValue;
                            }
                            this._evaluatedValue = Math.pow(leftValue, rightValue);
                            break;
                        default:
                            this._evaluatedValue = NaN;
                            break;
                    }
                    if (!isNaN(this._evaluatedValue)) {
                        return this._evaluatedValue;
                    }
                    throw 'Bad expression.';
                };
                return _BinaryExpression;
            }(_Expression));
            sheet_1._BinaryExpression = _BinaryExpression;
            /*
             * Defines the cell range expression class.
             * For e.g. A1 or A1:B2.
             */
            var _CellRangeExpression = (function (_super) {
                __extends(_CellRangeExpression, _super);
                /*
                 * Initializes a new instance of the @see:CellRangeExpression class.
                 *
                 * @param cells The @see:CellRange instance represents the cell range for the CellRangeExpression.
                 * @param sheetRef The sheet name of the sheet which the cells range refers.
                 * @param flex The @see:FlexSheet instance for evaluating the value for the CellRangeExpression.
                 */
                function _CellRangeExpression(cells, sheetRef, flex) {
                    _super.call(this);
                    this._cells = cells;
                    this._sheetRef = sheetRef;
                    this._flex = flex;
                    this._evalutingRange = {};
                }
                /*
                 * Overrides the evaluate function of base class.
                 *
                 * @param sheet The @see:Sheet is referenced by the @see:Expression.
                 */
                _CellRangeExpression.prototype.evaluate = function (sheet) {
                    if (this._evaluatedValue == null) {
                        this._evaluatedValue = this._getCellValue(this._cells, sheet);
                    }
                    return this._evaluatedValue;
                };
                /*
                 * Gets the value list for each cell inside the cell range.
                 *
                 * @param isGetHiddenValue indicates whether get the cell value of the hidden row or hidden column.
                 * @param columnIndex indicates which column of the cell range need be get.
                 * @param sheet The @see:Sheet whose value to evaluate. If not specified then the data from current sheet
                 */
                _CellRangeExpression.prototype.getValues = function (isGetHiddenValue, columnIndex, sheet) {
                    if (isGetHiddenValue === void 0) { isGetHiddenValue = true; }
                    var cellValue, vals = [], valIndex = 0, rowIndex, columnIndex, startColumnIndex, endColumnIndex;
                    startColumnIndex = columnIndex != null && !isNaN(+columnIndex) ? columnIndex : this._cells.leftCol;
                    endColumnIndex = columnIndex != null && !isNaN(+columnIndex) ? columnIndex : this._cells.rightCol;
                    sheet = this._getSheet() || sheet || this._flex.selectedSheet;
                    if (!sheet) {
                        return null;
                    }
                    for (rowIndex = this._cells.topRow; rowIndex <= this._cells.bottomRow; rowIndex++) {
                        if (rowIndex >= sheet.grid.rows.length) {
                            throw 'The cell reference is out of the cell range of the flexsheet.';
                        }
                        if (!isGetHiddenValue && sheet.grid.rows[rowIndex].isVisible === false) {
                            continue;
                        }
                        for (columnIndex = startColumnIndex; columnIndex <= endColumnIndex; columnIndex++) {
                            if (columnIndex >= sheet.grid.columns.length) {
                                throw 'The cell reference is out of the cell range of the flexsheet.';
                            }
                            if (!isGetHiddenValue && sheet.grid.columns[columnIndex].isVisible === false) {
                                continue;
                            }
                            cellValue = this._getCellValue(new grid.CellRange(rowIndex, columnIndex), sheet);
                            if (!wijmo.isPrimitive(cellValue)) {
                                cellValue = cellValue.value;
                            }
                            vals[valIndex] = cellValue;
                            valIndex++;
                        }
                    }
                    return vals;
                };
                Object.defineProperty(_CellRangeExpression.prototype, "cells", {
                    /*
                     * Gets the cell range of the CellRangeExpression.
                     */
                    get: function () {
                        return this._cells;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(_CellRangeExpression.prototype, "sheetRef", {
                    /*
                     * Gets the sheet reference of the CellRangeExpression.
                     */
                    get: function () {
                        return this._sheetRef;
                    },
                    enumerable: true,
                    configurable: true
                });
                // Get cell value for a cell.
                _CellRangeExpression.prototype._getCellValue = function (cell, sheet) {
                    var sheet, cellKey;
                    sheet = this._getSheet() || sheet || this._flex.selectedSheet;
                    if (!sheet) {
                        return null;
                    }
                    cellKey = sheet.name + ':' + cell.row + ',' + cell.col + '-' + cell.row2 + ',' + cell.col2;
                    if (this._evalutingRange[cellKey]) {
                        throw 'Circular Reference';
                    }
                    try {
                        if (this._flex) {
                            this._evalutingRange[cellKey] = true;
                            return this._flex.getCellValue(cell.row, cell.col, false, sheet);
                        }
                    }
                    finally {
                        delete this._evalutingRange[cellKey];
                    }
                };
                // Gets the sheet by the sheetRef.
                _CellRangeExpression.prototype._getSheet = function () {
                    var i = 0, sheet;
                    if (!this._sheetRef) {
                        return null;
                    }
                    for (; i < this._flex.sheets.length; i++) {
                        sheet = this._flex.sheets[i];
                        if (sheet.name === this._sheetRef) {
                            return sheet;
                        }
                    }
                    throw 'Invalid sheet reference';
                };
                return _CellRangeExpression;
            }(_Expression));
            sheet_1._CellRangeExpression = _CellRangeExpression;
            /*
             * Defines the function expression class.
             * For e.g. sum(1,2,3).
             */
            var _FunctionExpression = (function (_super) {
                __extends(_FunctionExpression, _super);
                /*
                 * Initializes a new instance of the @see:FunctionExpression class.
                 *
                 * @param func The @see:FunctionDefinition instance keeps function name, parameter counts, and function.
                 * @param params The parameter list that the function of the @see:FunctionDefinition instance needs.
                 */
                function _FunctionExpression(func, params) {
                    _super.call(this);
                    this._funcDefinition = func;
                    this._params = params;
                }
                /*
                 * Overrides the evaluate function of base class.
                 *
                 * @param sheet The @see:Sheet is referenced by the @see:Expression.
                 * @param rowIndex The row index of the cell where the expression located in.
                 * @param columnIndex The column index of the cell where the expression located in.
                 */
                _FunctionExpression.prototype.evaluate = function (sheet, rowIndex, columnIndex) {
                    if (this._evaluatedValue == null) {
                        this._evaluatedValue = this._funcDefinition.func(this._params, sheet, rowIndex, columnIndex);
                    }
                    return this._evaluatedValue;
                };
                return _FunctionExpression;
            }(_Expression));
            sheet_1._FunctionExpression = _FunctionExpression;
        })(sheet = grid.sheet || (grid.sheet = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_Expression.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var sheet;
        (function (sheet) {
            'use strict';
            /*
             * Base class for Flexsheet undo/redo actions.
             */
            var _UndoAction = (function () {
                /*
                 * Initializes a new instance of the @see:_UndoAction class.
                 *
                 * @param owner The @see: FlexSheet control that the @see:_UndoAction works for.
                 */
                function _UndoAction(owner) {
                    this._owner = owner;
                    this._sheetIndex = owner.selectedSheetIndex;
                }
                Object.defineProperty(_UndoAction.prototype, "sheetIndex", {
                    /*
                     * Gets the index of the sheet that the undo action wokrs for.
                     */
                    get: function () {
                        return this._sheetIndex;
                    },
                    enumerable: true,
                    configurable: true
                });
                /*
                 * Executes undo of the undo action
                 */
                _UndoAction.prototype.undo = function () {
                    throw 'This abstract method must be overrided.';
                };
                /*
                 * Executes redo of the undo action
                 */
                _UndoAction.prototype.redo = function () {
                    throw 'This abstract method must be overrided.';
                };
                /*
                 * Saves the current flexsheet state.
                 */
                _UndoAction.prototype.saveNewState = function () {
                    throw 'This abstract method must be overrided.';
                };
                return _UndoAction;
            }());
            sheet._UndoAction = _UndoAction;
            /*
             * Defines the _EditAction class.
             *
             * It deals with the undo\redo for editing value of the flexsheet cells.
             */
            var _EditAction = (function (_super) {
                __extends(_EditAction, _super);
                /*
                 * Initializes a new instance of the @see:_EditAction class.
                 *
                 * @param owner The @see: FlexSheet control that the _EditAction works for.
                 * @param selection The @CellRange of current editing cell.
                 */
                function _EditAction(owner, selection) {
                    var index, selection, rowIndex, colIndex, val;
                    _super.call(this, owner);
                    this._isPaste = false;
                    this._selections = selection ? [selection] : owner.selectedSheet.selectionRanges.slice();
                    this._oldValues = [];
                    for (index = 0; index < this._selections.length; index++) {
                        selection = this._selections[index];
                        for (rowIndex = selection.topRow; rowIndex <= selection.bottomRow; rowIndex++) {
                            for (colIndex = selection.leftCol; colIndex <= selection.rightCol; colIndex++) {
                                val = owner.getCellData(rowIndex, colIndex, !!owner.columns[colIndex].dataMap);
                                val = val == undefined ? '' : val;
                                this._oldValues.push(val);
                            }
                        }
                    }
                }
                Object.defineProperty(_EditAction.prototype, "isPaste", {
                    /*
                     * Gets the isPaste state to indicate the edit action works for edit cell or copy/paste.
                     */
                    get: function () {
                        return this._isPaste;
                    },
                    enumerable: true,
                    configurable: true
                });
                /*
                 * Overrides the undo method of its base class @see:_UndoAction.
                 */
                _EditAction.prototype.undo = function () {
                    var i = 0, index, selection, rowIndex, colIndex;
                    this._owner._clearCalcEngine();
                    this._owner.selectedSheet.selectionRanges.clear();
                    for (index = 0; index < this._selections.length; index++) {
                        selection = this._selections[index];
                        for (rowIndex = selection.topRow; rowIndex <= selection.bottomRow; rowIndex++) {
                            for (colIndex = selection.leftCol; colIndex <= selection.rightCol; colIndex++) {
                                this._owner.setCellData(rowIndex, colIndex, this._oldValues[i]);
                                i++;
                            }
                        }
                        this._owner.selectedSheet.selectionRanges.push(selection);
                    }
                    this._owner.refresh(false);
                };
                /*
                 * Overrides the redo method of its base class @see:_UndoAction.
                 */
                _EditAction.prototype.redo = function () {
                    var i = 0, index, selection, rowIndex, colIndex;
                    this._owner._clearCalcEngine();
                    this._owner.selectedSheet.selectionRanges.clear();
                    for (index = 0; index < this._selections.length; index++) {
                        selection = this._selections[index];
                        for (rowIndex = selection.topRow; rowIndex <= selection.bottomRow; rowIndex++) {
                            for (colIndex = selection.leftCol; colIndex <= selection.rightCol; colIndex++) {
                                this._owner.setCellData(rowIndex, colIndex, this._newValues[i]);
                                i++;
                            }
                        }
                        this._owner.selectedSheet.selectionRanges.push(selection);
                    }
                    this._owner.refresh(false);
                };
                /*
                 * Overrides the saveNewState of its base class @see:_UndoAction.
                 */
                _EditAction.prototype.saveNewState = function () {
                    var index, selection, rowIndex, currentCol, rowIndex, colIndex, val;
                    this._newValues = [];
                    for (index = 0; index < this._selections.length; index++) {
                        selection = this._selections[index];
                        for (rowIndex = selection.topRow; rowIndex <= selection.bottomRow; rowIndex++) {
                            for (colIndex = selection.leftCol; colIndex <= selection.rightCol; colIndex++) {
                                currentCol = this._owner.columns[colIndex];
                                if (!currentCol) {
                                    return false;
                                }
                                val = this._owner.getCellData(rowIndex, colIndex, !!currentCol.dataMap);
                                val = val == undefined ? '' : val;
                                this._newValues.push(val);
                            }
                        }
                    }
                    return !this._checkActionState();
                };
                /*
                 * Mark the cell edit action works for paste action.
                 */
                _EditAction.prototype.markIsPaste = function () {
                    this._isPaste = true;
                };
                /*
                 * Update the edit action for pasting.
                 *
                 * @param rng the @see:CellRange used to update the edit action
                 */
                _EditAction.prototype.updateForPasting = function (rng) {
                    var selection = this._selections[this._selections.length - 1], val = this._owner.getCellData(rng.row, rng.col, !!this._owner.columns[rng.col].dataMap);
                    if (!this._addingValue) {
                        this._addingValue = true;
                        this._oldValues = [];
                    }
                    val = val == undefined ? '' : val;
                    this._oldValues.push(val);
                    selection.row2 = Math.max(selection.row2, rng.row2);
                    selection.col2 = Math.max(selection.col2, rng.col2);
                };
                // Check whether the values changed after editing.
                _EditAction.prototype._checkActionState = function () {
                    var i;
                    if (this._oldValues.length !== this._newValues.length) {
                        return false;
                    }
                    for (i = 0; i < this._oldValues.length; i++) {
                        if (this._oldValues[i] !== this._newValues[i]) {
                            return false;
                        }
                    }
                    return true;
                };
                return _EditAction;
            }(_UndoAction));
            sheet._EditAction = _EditAction;
            /*
             * Defines the _ColumnResizeAction class.
             *
             * It deals with the undo/redo for resize the column of the flexsheet.
             */
            var _ColumnResizeAction = (function (_super) {
                __extends(_ColumnResizeAction, _super);
                /*
                 * Initializes a new instance of the @see:_ColumnResizeAction class.
                 *
                 * @param owner The @see: FlexSheet control that the _ColumnResizeAction works for.
                 * @param panel The @see: GridPanel indicates the resizing column belongs to which part of the FlexSheet.
                 * @param colIndex it indicates which column is resizing.
                 */
                function _ColumnResizeAction(owner, panel, colIndex) {
                    _super.call(this, owner);
                    this._panel = panel;
                    this._colIndex = colIndex;
                    this._oldColWidth = panel.columns[colIndex].width;
                }
                /*
                 * Overrides the undo method of its base class @see:_UndoAction.
                 */
                _ColumnResizeAction.prototype.undo = function () {
                    this._panel.columns[this._colIndex].width = this._oldColWidth;
                };
                /*
                 * Overrides the redo method of its base class @see:_UndoAction.
                 */
                _ColumnResizeAction.prototype.redo = function () {
                    this._panel.columns[this._colIndex].width = this._newColWidth;
                };
                /*
                 * Overrides the saveNewState method of its base class @see:_UndoAction.
                 */
                _ColumnResizeAction.prototype.saveNewState = function () {
                    this._newColWidth = this._panel.columns[this._colIndex].width;
                    if (this._oldColWidth === this._newColWidth) {
                        return false;
                    }
                    return true;
                };
                return _ColumnResizeAction;
            }(_UndoAction));
            sheet._ColumnResizeAction = _ColumnResizeAction;
            /*
             * Defines the _RowResizeAction class.
             *
             * It deals with the undo\redo for resize the row of the flexsheet.
             */
            var _RowResizeAction = (function (_super) {
                __extends(_RowResizeAction, _super);
                /*
                 * Initializes a new instance of the @see:_RowResizeAction class.
                 *
                 * @param owner The @see: FlexSheet control that the _RowResizeAction works for.
                 * @param panel The @see: GridPanel indicates the resizing row belongs to which part of the FlexSheet.
                 * @param rowIndex it indicates which row is resizing.
                 */
                function _RowResizeAction(owner, panel, rowIndex) {
                    _super.call(this, owner);
                    this._panel = panel;
                    this._rowIndex = rowIndex;
                    this._oldRowHeight = panel.rows[rowIndex].height;
                }
                /*
                 * Overrides the undo method of its base class @see:_UndoAction.
                 */
                _RowResizeAction.prototype.undo = function () {
                    this._panel.rows[this._rowIndex].height = this._oldRowHeight;
                };
                /*
                 * Overrides the redo method of its base class @see:_UndoAction.
                 */
                _RowResizeAction.prototype.redo = function () {
                    this._panel.rows[this._rowIndex].height = this._newRowHeight;
                };
                /*
                 * Overrides the saveNewState method of its base class @see:_UndoAction.
                 */
                _RowResizeAction.prototype.saveNewState = function () {
                    this._newRowHeight = this._panel.rows[this._rowIndex].height;
                    if (this._oldRowHeight === this._newRowHeight) {
                        return false;
                    }
                    return true;
                };
                return _RowResizeAction;
            }(_UndoAction));
            sheet._RowResizeAction = _RowResizeAction;
            /*
             * Defines the _InsertDeleteColumnAction class.
             *
             * It deals with the undo\redo for insert or delete column of the flexsheet.
             */
            var _ColumnsChangedAction = (function (_super) {
                __extends(_ColumnsChangedAction, _super);
                /*
                 * Initializes a new instance of the @see:_InsertDeleteColumnAction class.
                 *
                 * @param owner The @see: FlexSheet control that the _InsertDeleteColumnAction works for.
                 */
                function _ColumnsChangedAction(owner) {
                    var colIndex, columns = [];
                    _super.call(this, owner);
                    this._selection = owner.selection;
                    for (colIndex = 0; colIndex < owner.columns.length; colIndex++) {
                        columns.push(owner.columns[colIndex]);
                    }
                    this._oldValue = {
                        columns: columns,
                        sortList: owner.sortManager._committedList.slice(),
                        styledCells: owner.selectedSheet ? JSON.parse(JSON.stringify(owner.selectedSheet._styledCells)) : null,
                        mergedCells: owner._cloneMergedCells()
                    };
                }
                /*
                 * Overrides the undo method of its base class @see:_UndoAction.
                 */
                _ColumnsChangedAction.prototype.undo = function () {
                    var colIndex, i, formulaObj, oldFormulas, self = this;
                    if (!self._owner.selectedSheet) {
                        return;
                    }
                    self._owner._clearCalcEngine();
                    self._owner.finishEditing();
                    self._owner.columns.clear();
                    self._owner.selectedSheet._styledCells = undefined;
                    self._owner.selectedSheet._mergedRanges = undefined;
                    self._owner.columns.beginUpdate();
                    for (colIndex = 0; colIndex < self._oldValue.columns.length; colIndex++) {
                        self._owner.columns.push(self._oldValue.columns[colIndex]);
                    }
                    self._owner.columns.endUpdate();
                    self._owner.selectedSheet._styledCells = self._oldValue.styledCells;
                    self._owner.selectedSheet._mergedRanges = self._oldValue.mergedCells;
                    if (self._affectedFormulas) {
                        oldFormulas = self._affectedFormulas.oldFormulas;
                    }
                    self._owner.deferUpdate(function () {
                        self._owner.selection = self._selection;
                        // Set the 'old' formulas for redo.
                        if (!!oldFormulas && oldFormulas.length > 0) {
                            for (i = 0; i < oldFormulas.length; i++) {
                                formulaObj = oldFormulas[i];
                                self._owner.setCellData(formulaObj.point.x, formulaObj.point.y, formulaObj.formula);
                            }
                        }
                        // Synch with current sheet.
                        self._owner._copyTo(self._owner.selectedSheet);
                        self._owner._copyFrom(self._owner.selectedSheet);
                    });
                    // Synch the cell style for current sheet.
                    self._owner.selectedSheet.grid['wj_sheetInfo'].styledCells = self._owner.selectedSheet._styledCells;
                    // Synch the merged range for current sheet.
                    self._owner.selectedSheet.grid['wj_sheetInfo'].mergedRanges = self._owner.selectedSheet._mergedRanges;
                    self._owner.sortManager.sortDescriptions.sourceCollection = self._oldValue.sortList.slice();
                    self._owner.sortManager.commitSort(false);
                    self._owner.sortManager._refresh();
                    self._owner.selection = self._selection;
                    self._owner.refresh(true);
                    setTimeout(function () {
                        self._owner.rows._dirty = true;
                        self._owner.columns._dirty = true;
                        self._owner.refresh(true);
                    }, 10);
                };
                /*
                 * Overrides the redo method of its base class @see:_UndoAction.
                 */
                _ColumnsChangedAction.prototype.redo = function () {
                    var colIndex, i, formulaObj, newFormulas, self = this;
                    if (!self._owner.selectedSheet) {
                        return;
                    }
                    self._owner._clearCalcEngine();
                    self._owner.finishEditing();
                    self._owner.columns.clear();
                    self._owner.selectedSheet._styledCells = undefined;
                    self._owner.selectedSheet._mergedRanges = undefined;
                    self._owner.columns.beginUpdate();
                    for (colIndex = 0; colIndex < self._newValue.columns.length; colIndex++) {
                        self._owner.columns.push(self._newValue.columns[colIndex]);
                    }
                    self._owner.columns.endUpdate();
                    self._owner.selectedSheet._styledCells = self._newValue.styledCells;
                    self._owner.selectedSheet._mergedRanges = self._newValue.mergedCells;
                    if (self._affectedFormulas) {
                        newFormulas = self._affectedFormulas.newFormulas;
                    }
                    self._owner.deferUpdate(function () {
                        self._owner.selection = self._selection;
                        // Set the 'new' formulas for redo.
                        if (!!newFormulas && newFormulas.length > 0) {
                            for (i = 0; i < newFormulas.length; i++) {
                                formulaObj = newFormulas[i];
                                self._owner.setCellData(formulaObj.point.x, formulaObj.point.y, formulaObj.formula);
                            }
                        }
                        // Synch with current sheet.
                        self._owner._copyTo(self._owner.selectedSheet);
                        self._owner._copyFrom(self._owner.selectedSheet);
                    });
                    // Synch the cell style for current sheet.
                    self._owner.selectedSheet.grid['wj_sheetInfo'].styledCells = self._owner.selectedSheet._styledCells;
                    // Synch the merged range for current sheet.
                    self._owner.selectedSheet.grid['wj_sheetInfo'].mergedRanges = self._owner.selectedSheet._mergedRanges;
                    self._owner.sortManager.sortDescriptions.sourceCollection = self._newValue.sortList.slice();
                    self._owner.sortManager.commitSort(false);
                    self._owner.sortManager._refresh();
                    self._owner.selection = self._selection;
                    self._owner.refresh(true);
                    setTimeout(function () {
                        self._owner.rows._dirty = true;
                        self._owner.columns._dirty = true;
                        self._owner.refresh(true);
                    }, 10);
                };
                /*
                 * Overrides the saveNewState method of its base class @see:_UndoAction.
                 */
                _ColumnsChangedAction.prototype.saveNewState = function () {
                    var colIndex, columns = [];
                    for (colIndex = 0; colIndex < this._owner.columns.length; colIndex++) {
                        columns.push(this._owner.columns[colIndex]);
                    }
                    this._newValue = {
                        columns: columns,
                        sortList: this._owner.sortManager._committedList.slice(),
                        styledCells: this._owner.selectedSheet ? JSON.parse(JSON.stringify(this._owner.selectedSheet._styledCells)) : null,
                        mergedCells: this._owner._cloneMergedCells()
                    };
                    return true;
                };
                return _ColumnsChangedAction;
            }(_UndoAction));
            sheet._ColumnsChangedAction = _ColumnsChangedAction;
            /*
             * Defines the _InsertDeleteRowAction class.
             *
             * It deals with the undo\redo for insert or delete row of the flexsheet.
             */
            var _RowsChangedAction = (function (_super) {
                __extends(_RowsChangedAction, _super);
                /*
                 * Initializes a new instance of the @see:_InsertDeleteRowAction class.
                 *
                 * @param owner The @see: FlexSheet control that the _InsertDeleteRowAction works for.
                 */
                function _RowsChangedAction(owner) {
                    var rowIndex, colIndex, rows = [], columns = [];
                    _super.call(this, owner);
                    this._selection = owner.selection;
                    for (rowIndex = 0; rowIndex < owner.rows.length; rowIndex++) {
                        rows.push(owner.rows[rowIndex]);
                    }
                    for (colIndex = 0; colIndex < owner.columns.length; colIndex++) {
                        columns.push(owner.columns[colIndex]);
                    }
                    this._oldValue = {
                        rows: rows,
                        columns: columns,
                        itemsSource: owner.itemsSource ? owner.itemsSource.slice() : undefined,
                        styledCells: owner.selectedSheet ? JSON.parse(JSON.stringify(owner.selectedSheet._styledCells)) : null,
                        mergedCells: owner._cloneMergedCells()
                    };
                }
                /*
                 * Overrides the undo method of its base class @see:_UndoAction.
                 */
                _RowsChangedAction.prototype.undo = function () {
                    var rowIndex, colIndex, i, processingRow, formulaObj, oldFormulas, self = this, dataSourceBinding = !!self._oldValue.itemsSource;
                    if (!self._owner.selectedSheet) {
                        return;
                    }
                    self._owner._clearCalcEngine();
                    self._owner.finishEditing();
                    self._owner.columns.clear();
                    self._owner.rows.clear();
                    self._owner.selectedSheet._styledCells = undefined;
                    self._owner.selectedSheet._mergedRanges = undefined;
                    if (dataSourceBinding) {
                        self._owner.autoGenerateColumns = false;
                        self._owner.itemsSource = self._oldValue.itemsSource.slice();
                    }
                    self._owner.rows.beginUpdate();
                    for (rowIndex = 0; rowIndex < self._oldValue.rows.length; rowIndex++) {
                        processingRow = self._oldValue.rows[rowIndex];
                        if (dataSourceBinding) {
                            if (!processingRow.dataItem && !(processingRow instanceof sheet.HeaderRow)) {
                                self._owner.rows.splice(rowIndex, 0, processingRow);
                            }
                        }
                        else {
                            self._owner.rows.push(processingRow);
                        }
                    }
                    for (colIndex = 0; colIndex < self._oldValue.columns.length; colIndex++) {
                        self._owner.columns.push(self._oldValue.columns[colIndex]);
                    }
                    self._owner.rows.endUpdate();
                    self._owner.selectedSheet._styledCells = self._oldValue.styledCells;
                    self._owner.selectedSheet._mergedRanges = self._oldValue.mergedCells;
                    if (self._affecedFormulas) {
                        oldFormulas = self._affecedFormulas.oldFormulas;
                    }
                    self._owner.deferUpdate(function () {
                        self._owner.selection = self._selection;
                        // Set the 'old' formulas for redo.
                        if (!!oldFormulas && oldFormulas.length > 0) {
                            for (i = 0; i < oldFormulas.length; i++) {
                                formulaObj = oldFormulas[i];
                                self._owner.setCellData(formulaObj.point.x, formulaObj.point.y, formulaObj.formula);
                            }
                        }
                        // Synch with current sheet.
                        self._owner._copyTo(self._owner.selectedSheet);
                        self._owner._copyFrom(self._owner.selectedSheet);
                    });
                    // Synch the cell style for current sheet.
                    self._owner.selectedSheet.grid['wj_sheetInfo'].styledCells = self._owner.selectedSheet._styledCells;
                    // Synch the merged range for current sheet.
                    self._owner.selectedSheet.grid['wj_sheetInfo'].mergedRanges = self._owner.selectedSheet._mergedRanges;
                    self._owner.selection = self._selection;
                    self._owner.refresh(true);
                    setTimeout(function () {
                        self._owner.rows._dirty = true;
                        self._owner.columns._dirty = true;
                        self._owner.refresh(true);
                    }, 10);
                };
                /*
                 * Overrides the redo method of its base class @see:_UndoAction.
                 */
                _RowsChangedAction.prototype.redo = function () {
                    var rowIndex, colIndex, i, processingRow, formulaObj, newFormulas, self = this, dataSourceBinding = !!self._newValue.itemsSource;
                    if (!self._owner.selectedSheet) {
                        return;
                    }
                    self._owner._clearCalcEngine();
                    self._owner.finishEditing();
                    self._owner.columns.clear();
                    self._owner.rows.clear();
                    self._owner.selectedSheet._styledCells = undefined;
                    self._owner.selectedSheet._mergedRanges = undefined;
                    if (dataSourceBinding) {
                        self._owner.autoGenerateColumns = false;
                        self._owner.itemsSource = self._newValue.itemsSource.slice();
                    }
                    self._owner.rows.beginUpdate();
                    for (rowIndex = 0; rowIndex < self._newValue.rows.length; rowIndex++) {
                        processingRow = self._newValue.rows[rowIndex];
                        if (dataSourceBinding) {
                            if (!processingRow.dataItem && !(processingRow instanceof sheet.HeaderRow)) {
                                self._owner.rows.splice(rowIndex, 0, processingRow);
                            }
                        }
                        else {
                            self._owner.rows.push(processingRow);
                        }
                    }
                    for (colIndex = 0; colIndex < self._newValue.columns.length; colIndex++) {
                        self._owner.columns.push(self._newValue.columns[colIndex]);
                    }
                    self._owner.rows.endUpdate();
                    self._owner.selectedSheet._styledCells = self._newValue.styledCells;
                    self._owner.selectedSheet._mergedRanges = self._newValue.mergedCells;
                    if (self._affecedFormulas) {
                        newFormulas = self._affecedFormulas.newFormulas;
                    }
                    self._owner.deferUpdate(function () {
                        // Set the 'new' formulas for redo.
                        if (!!newFormulas && newFormulas.length > 0) {
                            for (i = 0; i < newFormulas.length; i++) {
                                formulaObj = newFormulas[i];
                                self._owner.setCellData(formulaObj.point.x, formulaObj.point.y, formulaObj.formula);
                            }
                        }
                        // Synch with current sheet.
                        self._owner._copyTo(self._owner.selectedSheet);
                        self._owner._copyFrom(self._owner.selectedSheet);
                    });
                    // Synch the cell style for current sheet.
                    self._owner.selectedSheet.grid['wj_sheetInfo'].styledCells = self._owner.selectedSheet._styledCells;
                    // Synch the merged range for current sheet.
                    self._owner.selectedSheet.grid['wj_sheetInfo'].mergedRanges = self._owner.selectedSheet._mergedRanges;
                    self._owner.selection = self._selection;
                    self._owner.refresh(true);
                    setTimeout(function () {
                        self._owner.rows._dirty = true;
                        self._owner.columns._dirty = true;
                        self._owner.refresh(true);
                    }, 10);
                };
                /*
                 * Overrides the saveNewState method of its base class @see:_UndoAction.
                 */
                _RowsChangedAction.prototype.saveNewState = function () {
                    var rowIndex, colIndex, rows = [], columns = [];
                    for (rowIndex = 0; rowIndex < this._owner.rows.length; rowIndex++) {
                        rows.push(this._owner.rows[rowIndex]);
                    }
                    for (colIndex = 0; colIndex < this._owner.columns.length; colIndex++) {
                        columns.push(this._owner.columns[colIndex]);
                    }
                    this._newValue = {
                        rows: rows,
                        columns: columns,
                        itemsSource: this._owner.itemsSource ? this._owner.itemsSource.slice() : undefined,
                        styledCells: this._owner.selectedSheet ? JSON.parse(JSON.stringify(this._owner.selectedSheet._styledCells)) : null,
                        mergedCells: this._owner._cloneMergedCells()
                    };
                    return true;
                };
                return _RowsChangedAction;
            }(_UndoAction));
            sheet._RowsChangedAction = _RowsChangedAction;
            /*
             * Defines the _CellStyleAction class.
             *
             * It deals with the undo\redo for applying style for the cells of the flexsheet.
             */
            var _CellStyleAction = (function (_super) {
                __extends(_CellStyleAction, _super);
                /*
                 * Initializes a new instance of the @see:_CellStyleAction class.
                 *
                 * @param owner The @see: FlexSheet control that the _CellStyleAction works for.
                 * @param styledCells Current styled cells of the @see: FlexSheet control.
                 */
                function _CellStyleAction(owner, styledCells) {
                    _super.call(this, owner);
                    this._oldStyledCells = styledCells ? JSON.parse(JSON.stringify(styledCells)) : (owner.selectedSheet ? JSON.parse(JSON.stringify(owner.selectedSheet._styledCells)) : null);
                }
                /*
                 * Overrides the undo method of its base class @see:_UndoAction.
                 */
                _CellStyleAction.prototype.undo = function () {
                    if (!this._owner.selectedSheet) {
                        return;
                    }
                    this._owner.selectedSheet._styledCells = JSON.parse(JSON.stringify(this._oldStyledCells));
                    this._owner.selectedSheet.grid['wj_sheetInfo'].styledCells = this._owner.selectedSheet._styledCells;
                    this._owner.refresh(false);
                };
                /*
                 * Overrides the redo method of its base class @see:_UndoAction.
                 */
                _CellStyleAction.prototype.redo = function () {
                    if (!this._owner.selectedSheet) {
                        return;
                    }
                    this._owner.selectedSheet._styledCells = JSON.parse(JSON.stringify(this._newStyledCells));
                    this._owner.selectedSheet.grid['wj_sheetInfo'].styledCells = this._owner.selectedSheet._styledCells;
                    this._owner.refresh(false);
                };
                /*
                 * Overrides the saveNewState method of its base class @see:_UndoAction.
                 */
                _CellStyleAction.prototype.saveNewState = function () {
                    this._newStyledCells = this._owner.selectedSheet ? JSON.parse(JSON.stringify(this._owner.selectedSheet._styledCells)) : null;
                    return true;
                };
                return _CellStyleAction;
            }(_UndoAction));
            sheet._CellStyleAction = _CellStyleAction;
            /*
             * Defines the _CellMergeAction class.
             *
             * It deals with the undo\redo for merging the cells of the flexsheet.
             */
            var _CellMergeAction = (function (_super) {
                __extends(_CellMergeAction, _super);
                /*
                 * Initializes a new instance of the @see:_CellMergeAction class.
                 *
                 * @param owner The @see: FlexSheet control that the _CellMergeAction works for.
                 */
                function _CellMergeAction(owner) {
                    _super.call(this, owner);
                    this._oldMergedCells = owner._cloneMergedCells();
                }
                /*
                 * Overrides the undo method of its base class @see:_UndoAction.
                 */
                _CellMergeAction.prototype.undo = function () {
                    if (!this._owner.selectedSheet) {
                        return;
                    }
                    this._owner._clearCalcEngine();
                    this._owner.selectedSheet._mergedRanges = this._oldMergedCells;
                    this._owner.selectedSheet.grid['wj_sheetInfo'].mergedRanges = this._owner.selectedSheet._mergedRanges;
                    this._owner.refresh(true);
                };
                /*
                 * Overrides the redo method of its base class @see:_UndoAction.
                 */
                _CellMergeAction.prototype.redo = function () {
                    if (!this._owner.selectedSheet) {
                        return;
                    }
                    this._owner._clearCalcEngine();
                    this._owner.selectedSheet._mergedRanges = this._newMergedCells;
                    this._owner.selectedSheet.grid['wj_sheetInfo'].mergedRanges = this._owner.selectedSheet._mergedRanges;
                    this._owner.refresh(true);
                };
                /*
                 * Overrides the saveNewState method of its base class @see:_UndoAction.
                 */
                _CellMergeAction.prototype.saveNewState = function () {
                    this._newMergedCells = this._owner._cloneMergedCells();
                    return true;
                };
                return _CellMergeAction;
            }(_UndoAction));
            sheet._CellMergeAction = _CellMergeAction;
            /*
             * Defines the _SortColumnAction class.
             *
             * It deals with the undo\redo for sort columns of the flexsheet.
             */
            var _SortColumnAction = (function (_super) {
                __extends(_SortColumnAction, _super);
                /*
                 * Initializes a new instance of the @see:_CellMergeAction class.
                 *
                 * @param owner The @see: FlexSheet control that the @see:_CellMergeAction works for.
                 */
                function _SortColumnAction(owner) {
                    var rowIndex, colIndex, columns = [], rows = [];
                    _super.call(this, owner);
                    if (!owner.itemsSource) {
                        for (rowIndex = 0; rowIndex < owner.rows.length; rowIndex++) {
                            rows.push(owner.rows[rowIndex]);
                        }
                        for (colIndex = 0; colIndex < owner.columns.length; colIndex++) {
                            columns.push(owner.columns[colIndex]);
                        }
                    }
                    this._oldValue = {
                        sortList: owner.sortManager._committedList.slice(),
                        rows: rows,
                        columns: columns
                    };
                }
                /*
                 * Overrides the undo method of its base class @see:_UndoAction.
                 */
                _SortColumnAction.prototype.undo = function () {
                    var _this = this;
                    var rowIndex, colIndex;
                    if (!this._owner.selectedSheet) {
                        return;
                    }
                    this._owner._clearCalcEngine();
                    this._owner.sortManager.sortDescriptions.sourceCollection = this._oldValue.sortList.slice();
                    this._owner.sortManager.commitSort(false);
                    this._owner.sortManager._refresh();
                    if (!this._owner.itemsSource) {
                        this._owner._isCopyingOrUndoing = true;
                        this._owner.rows.clear();
                        this._owner.columns.clear();
                        this._owner.selectedSheet.grid.rows.clear();
                        this._owner.selectedSheet.grid.columns.clear();
                        for (rowIndex = 0; rowIndex < this._oldValue.rows.length; rowIndex++) {
                            this._owner.rows.push(this._oldValue.rows[rowIndex]);
                            // Synch the rows of the grid for current sheet.
                            this._owner.selectedSheet.grid.rows.push(this._oldValue.rows[rowIndex]);
                        }
                        for (colIndex = 0; colIndex < this._oldValue.columns.length; colIndex++) {
                            this._owner.columns.push(this._oldValue.columns[colIndex]);
                            // Synch the columns of the grid for current sheet.
                            this._owner.selectedSheet.grid.columns.push(this._oldValue.columns[colIndex]);
                        }
                        this._owner._isCopyingOrUndoing = false;
                        setTimeout(function () {
                            _this._owner.rows._dirty = true;
                            _this._owner.columns._dirty = true;
                            _this._owner.refresh(true);
                        }, 10);
                    }
                };
                /*
                 * Overrides the redo method of its base class @see:_UndoAction.
                 */
                _SortColumnAction.prototype.redo = function () {
                    var _this = this;
                    var rowIndex, colIndex;
                    if (!this._owner.selectedSheet) {
                        return;
                    }
                    this._owner._clearCalcEngine();
                    this._owner.sortManager.sortDescriptions.sourceCollection = this._newValue.sortList.slice();
                    this._owner.sortManager.commitSort(false);
                    this._owner.sortManager._refresh();
                    if (!this._owner.itemsSource) {
                        this._owner._isCopyingOrUndoing = true;
                        this._owner.rows.clear();
                        this._owner.columns.clear();
                        this._owner.selectedSheet.grid.rows.clear();
                        this._owner.selectedSheet.grid.columns.clear();
                        for (rowIndex = 0; rowIndex < this._newValue.rows.length; rowIndex++) {
                            this._owner.rows.push(this._newValue.rows[rowIndex]);
                            // Synch the rows of the grid for current sheet.
                            this._owner.selectedSheet.grid.rows.push(this._newValue.rows[rowIndex]);
                        }
                        for (colIndex = 0; colIndex < this._newValue.columns.length; colIndex++) {
                            this._owner.columns.push(this._newValue.columns[colIndex]);
                            // Synch the columns of the grid for current sheet.
                            this._owner.selectedSheet.grid.columns.push(this._newValue.columns[colIndex]);
                        }
                        this._owner._isCopyingOrUndoing = false;
                        setTimeout(function () {
                            _this._owner.rows._dirty = true;
                            _this._owner.columns._dirty = true;
                            _this._owner.refresh(true);
                        }, 10);
                    }
                };
                /*
                 * Overrides the saveNewState method of its base class @see:_UndoAction.
                 */
                _SortColumnAction.prototype.saveNewState = function () {
                    var rowIndex, colIndex, columns = [], rows = [];
                    if (!this._owner.itemsSource) {
                        for (rowIndex = 0; rowIndex < this._owner.rows.length; rowIndex++) {
                            rows.push(this._owner.rows[rowIndex]);
                        }
                        for (colIndex = 0; colIndex < this._owner.columns.length; colIndex++) {
                            columns.push(this._owner.columns[colIndex]);
                        }
                    }
                    this._newValue = {
                        sortList: this._owner.sortManager._committedList.slice(),
                        rows: rows,
                        columns: columns
                    };
                    return true;
                };
                return _SortColumnAction;
            }(_UndoAction));
            sheet._SortColumnAction = _SortColumnAction;
            /*
             * Defines the _MoveCellsAction class.
             *
             * It deals with drag & drop the rows or columns to move or copy the cells action.
             */
            var _MoveCellsAction = (function (_super) {
                __extends(_MoveCellsAction, _super);
                /*
                 * Initializes a new instance of the @see:_MoveCellsAction class.
                 *
                 * @param owner The @see: FlexSheet control that the @see:_MoveCellsAction works for.
                 * @param draggingCells The @see: CellRange contains dragging target cells.
                 * @param droppingCells The @see: CellRange contains the dropping target cells.
                 * @param isCopyCells Indicates whether the action is moving or copying the cells.
                 */
                function _MoveCellsAction(owner, draggingCells, droppingCells, isCopyCells) {
                    var rowIndex, colIndex, cellIndex, val, cellStyle;
                    _super.call(this, owner);
                    if (!owner.selectedSheet) {
                        return;
                    }
                    if (draggingCells.topRow === 0 && draggingCells.bottomRow === owner.rows.length - 1) {
                        this._isDraggingColumns = true;
                    }
                    else {
                        this._isDraggingColumns = false;
                    }
                    this._isCopyCells = isCopyCells;
                    this._dragRange = draggingCells;
                    this._dropRange = droppingCells;
                    this._oldDroppingCells = [];
                    this._oldDroppingColumnSetting = {};
                    for (rowIndex = droppingCells.topRow; rowIndex <= droppingCells.bottomRow; rowIndex++) {
                        for (colIndex = droppingCells.leftCol; colIndex <= droppingCells.rightCol; colIndex++) {
                            if (this._isDraggingColumns) {
                                if (!this._oldDroppingColumnSetting[colIndex]) {
                                    this._oldDroppingColumnSetting[colIndex] = {
                                        dataType: owner.columns[colIndex].dataType,
                                        align: owner.columns[colIndex].align,
                                        format: owner.columns[colIndex].format
                                    };
                                }
                            }
                            cellIndex = rowIndex * this._owner.columns.length + colIndex;
                            if (this._owner.selectedSheet._styledCells[cellIndex]) {
                                cellStyle = JSON.parse(JSON.stringify(this._owner.selectedSheet._styledCells[cellIndex]));
                            }
                            else {
                                cellStyle = undefined;
                            }
                            val = this._owner.getCellData(rowIndex, colIndex, false);
                            this._oldDroppingCells.push({
                                rowIndex: rowIndex,
                                columnIndex: colIndex,
                                cellContent: val,
                                cellStyle: cellStyle
                            });
                        }
                    }
                    if (!isCopyCells) {
                        this._draggingCells = [];
                        this._draggingColumnSetting = {};
                        for (rowIndex = draggingCells.topRow; rowIndex <= draggingCells.bottomRow; rowIndex++) {
                            for (colIndex = draggingCells.leftCol; colIndex <= draggingCells.rightCol; colIndex++) {
                                if (this._isDraggingColumns) {
                                    if (!this._draggingColumnSetting[colIndex]) {
                                        this._draggingColumnSetting[colIndex] = {
                                            dataType: owner.columns[colIndex].dataType,
                                            align: owner.columns[colIndex].align,
                                            format: owner.columns[colIndex].format
                                        };
                                    }
                                }
                                cellIndex = rowIndex * this._owner.columns.length + colIndex;
                                if (this._owner.selectedSheet._styledCells[cellIndex]) {
                                    cellStyle = JSON.parse(JSON.stringify(this._owner.selectedSheet._styledCells[cellIndex]));
                                }
                                else {
                                    cellStyle = undefined;
                                }
                                val = this._owner.getCellData(rowIndex, colIndex, false);
                                this._draggingCells.push({
                                    rowIndex: rowIndex,
                                    columnIndex: colIndex,
                                    cellContent: val,
                                    cellStyle: cellStyle
                                });
                            }
                        }
                    }
                }
                /*
                 * Overrides the undo method of its base class @see:_UndoAction.
                 */
                _MoveCellsAction.prototype.undo = function () {
                    var self = this, index, moveCellActionValue, cellIndex, val, cellStyle, srcColIndex, descColIndex;
                    if (!self._owner.selectedSheet) {
                        return;
                    }
                    self._owner._clearCalcEngine();
                    for (index = 0; index < self._oldDroppingCells.length; index++) {
                        moveCellActionValue = self._oldDroppingCells[index];
                        self._owner.setCellData(moveCellActionValue.rowIndex, moveCellActionValue.columnIndex, moveCellActionValue.cellContent);
                        cellIndex = moveCellActionValue.rowIndex * self._owner.columns.length + moveCellActionValue.columnIndex;
                        if (moveCellActionValue.cellStyle) {
                            self._owner.selectedSheet._styledCells[cellIndex] = moveCellActionValue.cellStyle;
                        }
                        else {
                            delete self._owner.selectedSheet._styledCells[cellIndex];
                        }
                    }
                    if (self._isDraggingColumns && !!self._oldDroppingColumnSetting) {
                        Object.keys(self._oldDroppingColumnSetting).forEach(function (key) {
                            self._owner.columns[+key].dataType = self._oldDroppingColumnSetting[+key].dataType ? self._oldDroppingColumnSetting[+key].dataType : wijmo.DataType.Object;
                            self._owner.columns[+key].align = self._oldDroppingColumnSetting[+key].align;
                            self._owner.columns[+key].format = self._oldDroppingColumnSetting[+key].format;
                        });
                    }
                    if (!self._isCopyCells) {
                        for (index = 0; index < self._draggingCells.length; index++) {
                            moveCellActionValue = self._draggingCells[index];
                            self._owner.setCellData(moveCellActionValue.rowIndex, moveCellActionValue.columnIndex, moveCellActionValue.cellContent);
                            cellIndex = moveCellActionValue.rowIndex * self._owner.columns.length + moveCellActionValue.columnIndex;
                            if (moveCellActionValue.cellStyle) {
                                self._owner.selectedSheet._styledCells[cellIndex] = moveCellActionValue.cellStyle;
                            }
                        }
                        if (self._isDraggingColumns && !!self._draggingColumnSetting) {
                            Object.keys(self._draggingColumnSetting).forEach(function (key) {
                                self._owner.columns[+key].dataType = self._draggingColumnSetting[+key].dataType ? self._draggingColumnSetting[+key].dataType : wijmo.DataType.Object;
                                self._owner.columns[+key].align = self._draggingColumnSetting[+key].align;
                                self._owner.columns[+key].format = self._draggingColumnSetting[+key].format;
                            });
                        }
                        if (self._isDraggingColumns) {
                            if (self._dragRange.leftCol < self._dropRange.leftCol) {
                                descColIndex = self._dragRange.leftCol;
                                for (srcColIndex = self._dropRange.leftCol; srcColIndex <= self._dropRange.rightCol; srcColIndex++) {
                                    self._owner._updateColumnFiler(srcColIndex, descColIndex);
                                    descColIndex++;
                                }
                            }
                            else {
                                descColIndex = self._dragRange.rightCol;
                                for (srcColIndex = self._dropRange.rightCol; srcColIndex >= self._dropRange.leftCol; srcColIndex--) {
                                    self._owner._updateColumnFiler(srcColIndex, descColIndex);
                                    descColIndex--;
                                }
                            }
                        }
                    }
                };
                /*
                 * Overrides the redo method of its base class @see:_UndoAction.
                 */
                _MoveCellsAction.prototype.redo = function () {
                    var self = this, index, moveCellActionValue, cellIndex, val, cellStyle, srcColIndex, descColIndex;
                    if (!self._owner.selectedSheet) {
                        return;
                    }
                    self._owner._clearCalcEngine();
                    if (!self._isCopyCells) {
                        for (index = 0; index < self._draggingCells.length; index++) {
                            moveCellActionValue = self._draggingCells[index];
                            self._owner.setCellData(moveCellActionValue.rowIndex, moveCellActionValue.columnIndex, null);
                            cellIndex = moveCellActionValue.rowIndex * self._owner.columns.length + moveCellActionValue.columnIndex;
                            if (self._owner.selectedSheet._styledCells[cellIndex]) {
                                delete self._owner.selectedSheet._styledCells[cellIndex];
                            }
                        }
                        if (self._isDraggingColumns && !!self._draggingColumnSetting) {
                            Object.keys(self._draggingColumnSetting).forEach(function (key) {
                                self._owner.columns[+key].dataType = wijmo.DataType.Object;
                                self._owner.columns[+key].align = null;
                                self._owner.columns[+key].format = null;
                            });
                        }
                    }
                    for (index = 0; index < self._newDroppingCells.length; index++) {
                        moveCellActionValue = self._newDroppingCells[index];
                        self._owner.setCellData(moveCellActionValue.rowIndex, moveCellActionValue.columnIndex, moveCellActionValue.cellContent);
                        cellIndex = moveCellActionValue.rowIndex * self._owner.columns.length + moveCellActionValue.columnIndex;
                        if (moveCellActionValue.cellStyle) {
                            self._owner.selectedSheet._styledCells[cellIndex] = moveCellActionValue.cellStyle;
                        }
                        else {
                            delete self._owner.selectedSheet._styledCells[cellIndex];
                        }
                    }
                    if (self._isDraggingColumns && !!self._newDroppingColumnSetting) {
                        Object.keys(self._newDroppingColumnSetting).forEach(function (key) {
                            self._owner.columns[+key].dataType = self._newDroppingColumnSetting[+key].dataType ? self._newDroppingColumnSetting[+key].dataType : wijmo.DataType.Object;
                            self._owner.columns[+key].align = self._newDroppingColumnSetting[+key].align;
                            self._owner.columns[+key].format = self._newDroppingColumnSetting[+key].format;
                        });
                    }
                    if (self._isDraggingColumns && !self._isCopyCells) {
                        if (self._dragRange.leftCol > self._dropRange.leftCol) {
                            descColIndex = self._dropRange.leftCol;
                            for (srcColIndex = self._dragRange.leftCol; srcColIndex <= self._dragRange.rightCol; srcColIndex++) {
                                self._owner._updateColumnFiler(srcColIndex, descColIndex);
                                descColIndex++;
                            }
                        }
                        else {
                            descColIndex = self._dropRange.rightCol;
                            for (srcColIndex = self._dragRange.rightCol; srcColIndex >= self._dragRange.leftCol; srcColIndex--) {
                                self._owner._updateColumnFiler(srcColIndex, descColIndex);
                                descColIndex--;
                            }
                        }
                    }
                };
                /*
                 * Overrides the saveNewState method of its base class @see:_UndoAction.
                 */
                _MoveCellsAction.prototype.saveNewState = function () {
                    var rowIndex, colIndex, cellIndex, val, cellStyle;
                    if (!this._owner.selectedSheet) {
                        return false;
                    }
                    if (this._dropRange) {
                        this._newDroppingCells = [];
                        this._newDroppingColumnSetting = {};
                        for (rowIndex = this._dropRange.topRow; rowIndex <= this._dropRange.bottomRow; rowIndex++) {
                            for (colIndex = this._dropRange.leftCol; colIndex <= this._dropRange.rightCol; colIndex++) {
                                if (this._isDraggingColumns) {
                                    if (!this._newDroppingColumnSetting[colIndex]) {
                                        this._newDroppingColumnSetting[colIndex] = {
                                            dataType: this._owner.columns[colIndex].dataType,
                                            align: this._owner.columns[colIndex].align,
                                            format: this._owner.columns[colIndex].format
                                        };
                                    }
                                }
                                cellIndex = rowIndex * this._owner.columns.length + colIndex;
                                if (this._owner.selectedSheet._styledCells[cellIndex]) {
                                    cellStyle = JSON.parse(JSON.stringify(this._owner.selectedSheet._styledCells[cellIndex]));
                                }
                                else {
                                    cellStyle = undefined;
                                }
                                val = this._owner.getCellData(rowIndex, colIndex, false);
                                this._newDroppingCells.push({
                                    rowIndex: rowIndex,
                                    columnIndex: colIndex,
                                    cellContent: val,
                                    cellStyle: cellStyle
                                });
                            }
                        }
                        return true;
                    }
                    return false;
                };
                return _MoveCellsAction;
            }(_UndoAction));
            sheet._MoveCellsAction = _MoveCellsAction;
        })(sheet = grid.sheet || (grid.sheet = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_UndoAction.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var sheet;
        (function (sheet) {
            'use strict';
            /*
             * Defines the ContextMenu for a @see:FlexSheet control.
             */
            var _ContextMenu = (function (_super) {
                __extends(_ContextMenu, _super);
                /*
                 * Initializes a new instance of the _ContextMenu class.
                 *
                 * @param element The DOM element that will host the control, or a jQuery selector (e.g. '#theCtrl').
                 * @param owner The @see: FlexSheet control what the ContextMenu works with.
                 */
                function _ContextMenu(element, owner) {
                    _super.call(this, element);
                    this._owner = owner;
                    this.applyTemplate('', this.getTemplate(), {
                        _insRows: 'insert-rows',
                        _delRows: 'delete-rows',
                        _insCols: 'insert-columns',
                        _delCols: 'delete-columns',
                    });
                    this._init();
                }
                /*
                 * Show the context menu.
                 *
                 * @param e The mouse event.
                 * @param point The point indicates the position for the context menu.
                 */
                _ContextMenu.prototype.show = function (e, point) {
                    var posX = (point ? point.x : e.clientX) + (e ? window.pageXOffset : 0), //Left Position of Mouse Pointer
                    posY = (point ? point.y : e.clientY) + (e ? window.pageYOffset : 0); //Top Position of Mouse Pointer
                    this.hostElement.style.position = 'absolute';
                    this.hostElement.style.display = 'inline';
                    if (posY + this.hostElement.clientHeight > window.innerHeight) {
                        posY -= this.hostElement.clientHeight;
                    }
                    if (posX + this.hostElement.clientWidth > window.innerWidth) {
                        posX -= this.hostElement.clientWidth;
                    }
                    this.hostElement.style.top = posY + 'px';
                    this.hostElement.style.left = posX + 'px';
                };
                /*
                 * Hide the context menu.
                 */
                _ContextMenu.prototype.hide = function () {
                    this.hostElement.style.display = 'none';
                };
                // Initialize the context menu.
                _ContextMenu.prototype._init = function () {
                    var self = this;
                    self.hostElement.style.zIndex = '9999';
                    document.querySelector('body').appendChild(self.hostElement);
                    self.addEventListener(self.hostElement, 'contextmenu', function (e) {
                        e.preventDefault();
                    });
                    self.addEventListener(self._insRows, 'click', function (e) {
                        self._owner.insertRows();
                        self.hide();
                        self._owner.hostElement.focus();
                    });
                    self.addEventListener(self._delRows, 'click', function (e) {
                        self._owner.deleteRows();
                        self.hide();
                        self._owner.hostElement.focus();
                    });
                    self.addEventListener(self._insCols, 'click', function (e) {
                        self._owner.insertColumns();
                        self.hide();
                        self._owner.hostElement.focus();
                    });
                    self.addEventListener(self._delCols, 'click', function (e) {
                        self._owner.deleteColumns();
                        self.hide();
                        self._owner.hostElement.focus();
                    });
                };
                _ContextMenu.controlTemplate = '<div class="wj-context-menu" width="150px">' +
                    '<div class="wj-context-menu-item" wj-part="insert-rows">Insert Row</div>' +
                    '<div class="wj-context-menu-item" wj-part="delete-rows">Delete Rows</div>' +
                    '<div class="wj-context-menu-item" wj-part="insert-columns">Insert Column</div>' +
                    '<div class="wj-context-menu-item" wj-part="delete-columns">Delete Columns</div>' +
                    '</div>';
                return _ContextMenu;
            }(wijmo.Control));
            sheet._ContextMenu = _ContextMenu;
        })(sheet = grid.sheet || (grid.sheet = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_ContextMenu.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var sheet;
        (function (sheet) {
            'use strict';
            /*
             * Defines the _TabHolder control.
             */
            var _TabHolder = (function (_super) {
                __extends(_TabHolder, _super);
                /*
                 * Initializes a new instance of the _TabHolder class.
                 *
                 * @param element The DOM element that will host the control, or a jQuery selector (e.g. '#theCtrl').
                 * @param owner The @see: FlexSheet control that the _TabHolder control is associated to.
                 */
                function _TabHolder(element, owner) {
                    _super.call(this, element);
                    this._splitterMousedownHdl = this._splitterMousedownHandler.bind(this);
                    this._owner = owner;
                    if (this.hostElement.attributes['tabindex']) {
                        this.hostElement.attributes.removeNamedItem('tabindex');
                    }
                    // instantiate and apply template
                    this.applyTemplate('', this.getTemplate(), {
                        _divSheet: 'left',
                        _divSplitter: 'splitter',
                        _divRight: 'right'
                    });
                    this._init();
                }
                Object.defineProperty(_TabHolder.prototype, "sheetControl", {
                    /*
                     * Gets the SheetTabs control
                     */
                    get: function () {
                        return this._sheetControl;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(_TabHolder.prototype, "visible", {
                    //get scrollBar(): ScrollBar {
                    //	return this._hScrollbar;
                    //}
                    /*
                     * Gets or sets the visibility of the TabHolder control
                     */
                    get: function () {
                        return this.hostElement.style.display !== 'none';
                    },
                    set: function (value) {
                        this.hostElement.style.display = value ? 'block' : 'none';
                        this._divSheet.style.display = value ? 'block' : 'none';
                    },
                    enumerable: true,
                    configurable: true
                });
                /*
                 * Gets the Blanket size for the TabHolder control.
                 */
                _TabHolder.prototype.getSheetBlanketSize = function () {
                    //var scrollBarSize = ScrollBar.getSize();
                    //return (scrollBarSize === 0 ? 20 : scrollBarSize + 3);
                    return 20;
                };
                /*
                 * Adjust the size of the TabHolder control
                 */
                _TabHolder.prototype.adjustSize = function () {
                    var hScrollDis = this._owner.scrollSize.width - this._owner.clientSize.width, vScrollDis = this._owner.scrollSize.height - this._owner.clientSize.height, eParent = this._divSplitter.parentElement, 
                    //totalWidth: number,
                    leftWidth;
                    if (hScrollDis <= 0) {
                        eParent.style.minWidth = '100px';
                        this._divSplitter.style.display = 'none';
                        this._divRight.style.display = 'none';
                        this._divSheet.style.width = '100%';
                        this._divSplitter.removeEventListener('mousedown', this._splitterMousedownHdl, true);
                    }
                    else {
                        eParent.style.minWidth = '300px';
                        this._divSplitter.style.display = 'none';
                        this._divRight.style.display = 'none';
                        //totalWidth = eParent.clientWidth - this._divSplitter.offsetWidth;
                        this._divSheet.style.width = '100%';
                        //leftWidth = Math.ceil(totalWidth / 2);
                        //this._divSheet.style.width = leftWidth + 'px';
                        //this._divRight.style.width = (totalWidth - leftWidth) + 'px';
                        //if (vScrollDis <= 0) {
                        //	this._divHScrollbar.style.marginRight = '0px';
                        //} else {
                        //	this._divHScrollbar.style.marginRight = '20px';
                        //}
                        //this._hScrollbar.scrollDistance = hScrollDis;
                        //this._hScrollbar.scrollValue = -this._owner.scrollPosition.x;
                        this._divSplitter.removeEventListener('mousedown', this._splitterMousedownHdl, true);
                        this._divSplitter.addEventListener('mousedown', this._splitterMousedownHdl, true);
                    }
                    this._sheetControl._adjustSize();
                };
                // Init the size of the splitter.
                // And init the ScrollBar, SheetTabs control 
                _TabHolder.prototype._init = function () {
                    var self = this;
                    self._funSplitterMousedown = function (e) {
                        self._splitterMouseupHandler(e);
                    };
                    self._divSplitter.parentElement.style.height = self.getSheetBlanketSize() + 'px';
                    //init scrollbar
                    //self._hScrollbar = new ScrollBar(self._divHScrollbar);
                    //init sheet
                    self._sheetControl = new sheet._SheetTabs(self._divSheet, this._owner);
                    //self._owner.scrollPositionChanged.addHandler(() => {
                    //	self._hScrollbar.scrollValue = -self._owner.scrollPosition.x;
                    //});
                };
                // Mousedown event handler for the splitter
                _TabHolder.prototype._splitterMousedownHandler = function (e) {
                    this._startPos = e.pageX;
                    document.addEventListener('mousemove', this._splitterMousemoveHandler.bind(this), true);
                    document.addEventListener('mouseup', this._funSplitterMousedown, true);
                    e.preventDefault();
                };
                // Mousemove event handler for the splitter
                _TabHolder.prototype._splitterMousemoveHandler = function (e) {
                    if (this._startPos === null || typeof (this._startPos) === 'undefined') {
                        return;
                    }
                    this._adjustDis(e.pageX - this._startPos);
                };
                // Mouseup event handler for the splitter
                _TabHolder.prototype._splitterMouseupHandler = function (e) {
                    document.removeEventListener('mousemove', this._splitterMousemoveHandler, true);
                    document.removeEventListener('mouseup', this._funSplitterMousedown, true);
                    this._adjustDis(e.pageX - this._startPos);
                    this._startPos = null;
                };
                // Adjust the distance for the splitter
                _TabHolder.prototype._adjustDis = function (dis) {
                    var rightWidth = this._divRight.offsetWidth - dis, leftWidth = this._divSheet.offsetWidth + dis;
                    if (rightWidth <= 100) {
                        rightWidth = 100;
                        dis = this._divRight.offsetWidth - rightWidth;
                        leftWidth = this._divSheet.offsetWidth + dis;
                    }
                    else if (leftWidth <= 100) {
                        leftWidth = 100;
                        dis = leftWidth - this._divSheet.offsetWidth;
                        rightWidth = this._divRight.offsetWidth - dis;
                    }
                    if (dis == 0) {
                        return;
                    }
                    this._divRight.style.width = rightWidth + 'px';
                    this._divSheet.style.width = leftWidth + 'px';
                    this._startPos = this._startPos + dis;
                    //this._hScrollbar.invalidate(false);
                };
                _TabHolder.controlTemplate = '<div>' +
                    '<div wj-part="left" style ="float:left;height:100%;overflow:hidden"></div>' +
                    '<div wj-part="splitter" style="float:left;height:100%;width:6px;background-color:#e9eaee;padding:2px;cursor:e-resize"><div style="background-color:#8a9eb2;height:100%"></div></div>' +
                    '<div wj-part="right" style="float:left;height:100%;background-color:#e9eaee">' +
                    // We will use the native scrollbar of the flexgrid instead of the custom scrollbar of flexsheet (TFS 121971)
                    //'<div wj-part="hscrollbar" style="float:none;height:100%;border-left:1px solid #8a9eb2; padding-top:1px; display: none;"></div>' + // right scrollbar
                    '</div>' +
                    '</div>';
                return _TabHolder;
            }(wijmo.Control));
            sheet._TabHolder = _TabHolder;
        })(sheet = grid.sheet || (grid.sheet = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_TabHolder.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var sheet;
        (function (sheet) {
            'use strict';
            /*
             * Defines the _FlexSheetCellFactory class.
             *
             * This class extends the CellFactory of FlexGrid control.
             * It updates the content of the row/column header for the FlexSheet control.
             */
            var _FlexSheetCellFactory = (function (_super) {
                __extends(_FlexSheetCellFactory, _super);
                function _FlexSheetCellFactory() {
                    _super.apply(this, arguments);
                }
                /*
                 * Overrides the updateCell function of the CellFactory class.
                 *
                 * @param panel Part of the grid that owns this cell.
                 * @param r Index of this cell's row.
                 * @param c Index of this cell's column.
                 * @param cell Element that represents the cell.
                 * @param rng @see:CellRange that contains the cell's merged range, or null if the cell is not merged.
                 */
                _FlexSheetCellFactory.prototype.updateCell = function (panel, r, c, cell, rng) {
                    var g = panel.grid, r2 = r, c2 = c, content, cellIndex, flex, fc, val, data, isFormula, styleInfo, checkBox, input, bcol, format;
                    // We shall reset the styles of current cell before updating current cell.
                    if (panel.cellType === wijmo.grid.CellType.Cell) {
                        this._resetCellStyle(panel.columns[c], cell);
                    }
                    _super.prototype.updateCell.call(this, panel, r, c, cell, rng);
                    // adjust for merged ranges
                    if (rng && !rng.isSingleCell) {
                        r = rng.row;
                        c = rng.col;
                        r2 = rng.row2;
                        c2 = rng.col2;
                    }
                    bcol = g._getBindingColumn(panel, r, panel.columns[c]);
                    switch (panel.cellType) {
                        case wijmo.grid.CellType.RowHeader:
                            cell.textContent = (r + 1) + '';
                            break;
                        case wijmo.grid.CellType.ColumnHeader:
                            content = sheet.FlexSheet.convertNumberToAlpha(c);
                            cell.innerHTML = cell.innerHTML.replace(cell.textContent, '') + content;
                            cell.style.textAlign = 'center';
                            break;
                        case wijmo.grid.CellType.Cell:
                            flex = panel.grid;
                            cellIndex = r * flex.columns.length + c;
                            styleInfo = flex.selectedSheet && flex.selectedSheet._styledCells ? flex.selectedSheet._styledCells[cellIndex] : null;
                            //process the header row with binding
                            if (panel.rows[r] instanceof sheet.HeaderRow) {
                                cell.innerHTML = wijmo.escapeHtml(panel.columns[c].header);
                                wijmo.addClass(cell, 'wj-header-row');
                            }
                            else {
                                val = flex.getCellValue(r, c, false);
                                if (flex.editRange && flex.editRange.contains(r, c)) {
                                    data = flex.getCellData(r, c, false);
                                    isFormula = data != null && typeof data === 'string' && data[0] === '=';
                                    if (wijmo.isNumber(val) && !bcol.dataMap && !isFormula) {
                                        format = (styleInfo ? styleInfo.format : '') || bcol.format || 'n';
                                        val = this._getFormattedValue(val, format);
                                        input = cell.querySelector('input');
                                        if (input) {
                                            input.value = val;
                                        }
                                    }
                                }
                                else {
                                    if (panel.columns[c].dataType === wijmo.DataType.Boolean) {
                                        checkBox = cell.querySelector('[type="checkbox"]');
                                        if (checkBox) {
                                            checkBox.checked = flex.getCellValue(r, c);
                                        }
                                    }
                                    else if (bcol.dataMap) {
                                        val = flex.getCellValue(r, c, true);
                                        fc = cell.firstChild;
                                        if (fc && fc.nodeType === 3 && fc.nodeValue !== val) {
                                            fc.nodeValue = val;
                                        }
                                    }
                                    else {
                                        if (wijmo.isNumber(val)) {
                                            format = (styleInfo ? styleInfo.format : '') || bcol.format;
                                            if (!format) {
                                                val = this._getFormattedValue(val, 'n');
                                            }
                                            else {
                                                val = flex.getCellValue(r, c, true);
                                            }
                                        }
                                        else {
                                            val = flex.getCellValue(r, c, true);
                                        }
                                        cell.innerHTML = val;
                                    }
                                }
                                if (styleInfo) {
                                    var st = cell.style, styleInfoVal;
                                    for (var styleProp in styleInfo) {
                                        if (styleProp === 'className') {
                                            if (styleInfo.className) {
                                                wijmo.addClass(cell, styleInfo.className + '-style');
                                            }
                                        }
                                        else if (styleProp !== 'format' && (styleInfoVal = styleInfo[styleProp])) {
                                            if ((wijmo.hasClass(cell, 'wj-state-selected') || wijmo.hasClass(cell, 'wj-state-multi-selected'))
                                                && (styleProp === 'color' || styleProp === 'backgroundColor')) {
                                                st[styleProp] = '';
                                            }
                                            else {
                                                st[styleProp] = styleInfoVal;
                                            }
                                        }
                                    }
                                }
                            }
                            // customize the cell
                            if (g.itemFormatter) {
                                g.itemFormatter(panel, r, c, cell);
                            }
                            if (g.formatItem.hasHandlers) {
                                var rng = grid.CellFactory._fmtRng;
                                if (!rng) {
                                    rng = grid.CellFactory._fmtRng = new grid.CellRange(r, c, r2, c2);
                                }
                                else {
                                    rng.setRange(r, c, r2, c2);
                                }
                                var e = new grid.FormatItemEventArgs(panel, rng, cell);
                                g.onFormatItem(e);
                            }
                            if (!!cell.style.backgroundColor || !!cell.style.color) {
                                if (!styleInfo) {
                                    flex.selectedSheet._styledCells[cellIndex] = styleInfo = {};
                                }
                                if (!!cell.style.backgroundColor) {
                                    styleInfo.backgroundColor = cell.style.backgroundColor;
                                }
                                if (!!cell.style.color) {
                                    styleInfo.color = cell.style.color;
                                }
                            }
                            break;
                    }
                };
                // Reset the styles of the cell.
                _FlexSheetCellFactory.prototype._resetCellStyle = function (column, cell) {
                    ['fontFamily', 'fontSize', 'fontStyle', 'fontWeight', 'textDecoration', 'textAlign', 'verticalAlign', 'backgroundColor', 'color'].forEach(function (val) {
                        if (val === 'textAlign') {
                            cell.style.textAlign = column.getAlignment();
                        }
                        else {
                            cell.style[val] = '';
                        }
                    });
                };
                // Get the formatted value.
                _FlexSheetCellFactory.prototype._getFormattedValue = function (value, format) {
                    var val;
                    if (value !== Math.round(value)) {
                        format = format.replace(/([a-z])(\d*)(.*)/ig, '$0112$3');
                    }
                    val = wijmo.Globalize.formatNumber(value, format, true);
                    return val;
                };
                return _FlexSheetCellFactory;
            }(grid.CellFactory));
            sheet._FlexSheetCellFactory = _FlexSheetCellFactory;
        })(sheet = grid.sheet || (grid.sheet = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_FlexSheetCellFactory.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
/**
* Defines the @see:FlexSheet control and associated classes.
*
* The @see:FlexSheet control extends the @see:FlexGrid control to provide Excel-like
* features.
*/
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid_1) {
        var sheet;
        (function (sheet_1) {
            'use strict';
            var FlexSheetFunctions = [
                { name: 'abs', description: 'Returns the absolute value of a number.' },
                { name: 'acos', description: 'Returns the arccosine of a number.' },
                { name: 'and', description: 'Returns TRUE if all of its arguments are TRUE.' },
                { name: 'asin', description: 'Returns the arcsine of a number.' },
                { name: 'atan', description: 'Returns the arctangent of a number.' },
                { name: 'atan2', description: 'Returns the arctangent from x- and y-coordinates.' },
                { name: 'average', description: 'Returns the average of its arguments.' },
                { name: 'ceiling', description: 'Rounds a number to the nearest integer or to the nearest multiple of significance.' },
                { name: 'char', description: 'Returns the character specified by the code number.' },
                { name: 'choose', description: 'Chooses a value from a list of values.' },
                { name: 'code', description: 'Returns a numeric code for the first character in a text string.' },
                { name: 'column', description: 'Returns the column number of a reference.' },
                { name: 'columns', description: 'Returns the number of columns in a reference.' },
                { name: 'concatenate', description: 'Joins several text items into one text item.' },
                { name: 'cos', description: 'Returns the cosine of a number.' },
                { name: 'count', description: 'Counts how many numbers are in the list of arguments.' },
                { name: 'counta', description: 'Counts how many values are in the list of arguments.' },
                { name: 'countblank', description: 'Counts the number of blank cells within a range.' },
                { name: 'countif', description: 'Counts the number of cells within a range that meet the given criteria.' },
                { name: 'countifs', description: 'Counts the number of cells within a range that meet multiple criteria.' },
                { name: 'date', description: 'Returns the serial number of a particular date.' },
                { name: 'datedif', description: 'Calculates the number of days, months, or years between two dates.' },
                { name: 'day', description: 'Converts a serial number to a day of the month.' },
                { name: 'dcount', description: 'Counts the cells that contain numbers in a database.' },
                { name: 'exp', description: 'Returns e raised to the power of a given number.' },
                { name: 'false', description: 'Returns the logical value FALSE.' },
                { name: 'find', description: 'Finds one text value within another (case-sensitive).' },
                { name: 'floor', description: 'Rounds a number down, toward zero.' },
                { name: 'hlookup', description: 'Looks in the top row of an array and returns the value of the indicated cell.' },
                { name: 'hour', description: 'Converts a serial number to an hour.' },
                { name: 'if', description: 'Specifies a logical test to perform.' },
                { name: 'index', description: 'Uses an index to choose a value from a reference.' },
                { name: 'left', description: 'Returns the leftmost characters from a text value.' },
                { name: 'len', description: 'Returns the number of characters in a text string.' },
                { name: 'ln', description: 'Returns the natural logarithm of a number.' },
                { name: 'lower', description: 'Converts text to lowercase.' },
                { name: 'max', description: 'Returns the maximum value in a list of arguments.' },
                { name: 'mid', description: 'Returns a specific number of characters from a text string starting at the position you specify.' },
                { name: 'min', description: 'Returns the minimum value in a list of arguments.' },
                { name: 'mod', description: 'Returns the remainder from division.' },
                { name: 'month', description: 'Converts a serial number to a month.' },
                { name: 'not', description: 'Reverses the logic of its argument.' },
                { name: 'now', description: 'Returns the serial number of the current date and time.' },
                { name: 'or', description: 'Returns TRUE if any argument is TRUE.' },
                { name: 'pi', description: 'Returns the value of pi.' },
                { name: 'power', description: 'Returns the result of a number raised to a power.' },
                { name: 'product', description: 'Multiplies its arguments.' },
                { name: 'proper', description: 'Capitalizes the first letter in each word of a text value.' },
                { name: 'rand', description: 'Returns a random number between 0 and 1.' },
                { name: 'rank', description: 'Returns the rank of a number in a list of numbers.' },
                { name: 'rate', description: 'Returns the interest rate per period of an annuity.' },
                { name: 'replace', description: 'Replaces characters within text.' },
                { name: 'rept', description: 'Repeats text a given number of times.' },
                { name: 'right', description: 'Returns the rightmost characters from a text value.' },
                { name: 'round', description: 'Rounds a number to a specified number of digits.' },
                { name: 'rounddown', description: 'Rounds a number down, toward zero.' },
                { name: 'roundup', description: 'Rounds a number up, away from zero.' },
                { name: 'row', description: 'Returns the row number of a reference.' },
                { name: 'rows', description: 'Returns the number of rows in a reference.' },
                { name: 'search', description: 'Finds one text value within another (not case-sensitive).' },
                { name: 'sin', description: 'Returns the sine of the given angle.' },
                { name: 'sqrt', description: 'Returns a positive square root.' },
                { name: 'stdev', description: 'Estimates standard deviation based on a sample.' },
                { name: 'stdevp', description: 'Calculates standard deviation based on the entire population.' },
                { name: 'substitute', description: 'Substitutes new text for old text in a text string.' },
                { name: 'subtotal', description: 'Returns a subtotal in a list or database.' },
                { name: 'sum', description: 'Adds its arguments.' },
                { name: 'sumif', description: 'Adds the cells specified by a given criteria.' },
                { name: 'sumifs', description: 'Adds the cells in a range that meet multiple criteria.' },
                { name: 'tan', description: 'Returns the tangent of a number.' },
                { name: 'text', description: 'Formats a number and converts it to text.' },
                { name: 'time', description: 'Returns the serial number of a particular time.' },
                { name: 'today', description: 'Returns the serial number of today\'s date.' },
                { name: 'trim', description: 'Removes spaces from text.' },
                { name: 'true', description: 'Returns the logical value TRUE.' },
                { name: 'trunc', description: 'Truncates a number to an integer.' },
                { name: 'upper', description: 'Converts text to uppercase.' },
                { name: 'value', description: 'Converts a text argument to a number.' },
                { name: 'var', description: 'Estimates variance based on a sample.' },
                { name: 'varp', description: 'Calculates variance based on the entire population.' },
                { name: 'year', description: 'Converts a serial number to a year.' },
            ];
            /**
             * Defines the @see:FlexSheet control.
             *
             * The @see:FlexSheet control extends the @see:FlexGrid control to provide Excel-like
             * features such as a calculation engine, multiple sheets, undo/redo, and
             * XLSX import/export.
             */
            var FlexSheet = (function (_super) {
                __extends(FlexSheet, _super);
                /**
                 * Initializes a new instance of the @see:FlexSheet class.
                 *
                 * @param element The DOM element that will host the control, or a jQuery selector (e.g. '#theCtrl').
                 * @param options JavaScript object containing initialization data for the control.
                 */
                function FlexSheet(element, options) {
                    _super.call(this, element, options);
                    this._selectedSheetIndex = -1;
                    this._columnHeaderClicked = false;
                    this._addingSheet = false;
                    this._mouseMoveHdl = this._mouseMove.bind(this);
                    this._clickHdl = this._click.bind(this);
                    this._touchStartHdl = this._touchStart.bind(this);
                    this._touchEndHdl = this._touchEnd.bind(this);
                    this._isClicking = false;
                    /**
                     * Occurs when current sheet index changed.
                     */
                    this.selectedSheetChanged = new wijmo.Event();
                    /**
                     * Occurs when dragging the rows or the columns of the <b>FlexSheet</b>.
                     */
                    this.draggingRowColumn = new wijmo.Event();
                    /**
                     * Occurs when dropping the rows or the columns of the <b>FlexSheet</b>.
                     */
                    this.droppingRowColumn = new wijmo.Event();
                    /**
                     * Occurs after the @see:FlexSheet loads the @see:Workbook instance
                     */
                    this.loaded = new wijmo.Event();
                    /**
                     * Occurs when the @see:FlexSheet meets the unknown formula.
                     */
                    this.unknownFunction = new wijmo.Event();
                    /**
                     * Occurs when the @see:FlexSheet is cleared.
                     */
                    this.sheetCleared = new wijmo.Event();
                    this['_eCt'].style.backgroundColor = 'white';
                    // We will use the native scrollbar of the flexgrid instead of the custom scrollbar of flexsheet (TFS 121971)
                    //this['_root'].style.overflowX = 'hidden';
                    wijmo.addClass(this.hostElement, 'wj-flexsheet');
                    // Set the default font to Arial of the FlexSheet control (TFS 127769) 
                    wijmo.setCss(this.hostElement, {
                        fontFamily: 'Arial'
                    });
                    this['_cf'] = new sheet_1._FlexSheetCellFactory();
                    // initialize the splitter, the sheet tab and the hscrollbar.
                    this._init();
                    this.showSort = false;
                    this.allowSorting = false;
                    this.showGroups = false;
                    this.showMarquee = true;
                    this.showSelectedHeaders = grid_1.HeadersVisibility.All;
                    this.allowResizing = grid_1.AllowResizing.Both;
                    this.allowDragging = grid_1.AllowDragging.None;
                }
                Object.defineProperty(FlexSheet.prototype, "sheets", {
                    /**
                     * Gets the collection of @see:Sheet objects representing workbook sheets.
                     */
                    get: function () {
                        if (!this._sheets) {
                            this._sheets = new sheet_1.SheetCollection();
                            this._sheets.selectedSheetChanged.addHandler(this._selectedSheetChange, this);
                            this._sheets.collectionChanged.addHandler(this._sourceChange, this);
                            this._sheets.sheetVisibleChanged.addHandler(this._sheetVisibleChange, this);
                            this._sheets.sheetCleared.addHandler(this.onSheetCleared, this);
                        }
                        return this._sheets;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(FlexSheet.prototype, "selectedSheetIndex", {
                    /**
                     * Gets or sets the index of the current sheet in the @see:FlexSheet.
                     */
                    get: function () {
                        return this._selectedSheetIndex;
                    },
                    set: function (value) {
                        if (value !== this._selectedSheetIndex) {
                            this._showSheet(value);
                            this._sheets.selectedIndex = value;
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(FlexSheet.prototype, "selectedSheet", {
                    /**
                     * Gets the current @see:Sheet in the <b>FlexSheet</b>.
                     */
                    get: function () {
                        return this._sheets[this._selectedSheetIndex];
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(FlexSheet.prototype, "isFunctionListOpen", {
                    /**
                     * Gets a value indicating whether the function list is opened.
                     */
                    get: function () {
                        return this._functionListHost && this._functionListHost.style.display !== 'none';
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(FlexSheet.prototype, "isTabHolderVisible", {
                    /**
                     * Gets or sets a value indicating whether the TabHolder is visible.
                     */
                    get: function () {
                        return this._tabHolder.visible;
                    },
                    set: function (value) {
                        if (value !== this._tabHolder.visible) {
                            if (value) {
                                this._divContainer.style.height = (this._divContainer.parentElement.clientHeight - this._tabHolder.getSheetBlanketSize()) + 'px';
                            }
                            else {
                                this._divContainer.style.height = this._divContainer.parentElement.clientHeight + 'px';
                            }
                            this._tabHolder.visible = value;
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(FlexSheet.prototype, "undoStack", {
                    /**
                     * Gets the @see:UndoStack instance that controls undo and redo operations of the <b>FlexSheet</b>.
                     */
                    get: function () {
                        return this._undoStack;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(FlexSheet.prototype, "sortManager", {
                    /**
                     * Gets the @see:SortManager instance that controls <b>FlexSheet</b> sorting.
                     */
                    get: function () {
                        return this._sortManager;
                    },
                    enumerable: true,
                    configurable: true
                });
                /**
                 * Raises the currentSheetChanged event.
                 *
                 * @param e @see:PropertyChangedEventArgs that contains the event data.
                 */
                FlexSheet.prototype.onSelectedSheetChanged = function (e) {
                    this._sortManager._refresh();
                    this.selectedSheetChanged.raise(this, e);
                };
                /**
                 * Raises the draggingRowColumn event.
                 */
                FlexSheet.prototype.onDraggingRowColumn = function (e) {
                    this.draggingRowColumn.raise(this, e);
                };
                /**
                 * Raises the droppingRowColumn event.
                 */
                FlexSheet.prototype.onDroppingRowColumn = function () {
                    this.droppingRowColumn.raise(this, new wijmo.EventArgs());
                };
                /**
                 * Raises the loaded event.
                 */
                FlexSheet.prototype.onLoaded = function () {
                    var self = this;
                    if (self._toRefresh) {
                        clearTimeout(self._toRefresh);
                        self._toRefresh = null;
                    }
                    self._toRefresh = setTimeout(function () {
                        self.rows._dirty = true;
                        self.columns._dirty = true;
                        self.invalidate();
                    }, 10);
                    self.loaded.raise(this, new wijmo.EventArgs());
                };
                /**
                 * Raises the unknownFunction event.
                 */
                FlexSheet.prototype.onUnknownFunction = function (e) {
                    this.unknownFunction.raise(this, e);
                };
                /**
                 * Raises the sheetCleared event.
                 */
                FlexSheet.prototype.onSheetCleared = function () {
                    this.sheetCleared.raise(this, new wijmo.EventArgs());
                };
                /**
                 * Overridden to refresh the sheet and the TabHolder.
                 *
                 * @param fullUpdate Whether to update the control layout as well as the content.
                 */
                FlexSheet.prototype.refresh = function (fullUpdate) {
                    if (fullUpdate === void 0) { fullUpdate = true; }
                    this._divContainer.style.height = (this._divContainer.parentElement.clientHeight - (this.isTabHolderVisible ? this._tabHolder.getSheetBlanketSize() : 0)) + 'px';
                    if (!this.preserveSelectedState && !!this.selectedSheet) {
                        this.selectedSheet.selectionRanges.clear();
                        this.selectedSheet.selectionRanges.push(this.selection);
                    }
                    _super.prototype.refresh.call(this, fullUpdate);
                    this._tabHolder.adjustSize();
                };
                /**
                 * Overrides the setCellData function of the base class.
                 *
                 * @param r Index of the row that contains the cell.
                 * @param c Index, name, or binding of the column that contains the cell.
                 * @param value Value to store in the cell.
                 * @param coerce Whether to change the value automatically to match the column's data type.
                 * @return True if the value was stored successfully, false otherwise.
                 */
                FlexSheet.prototype.setCellData = function (r, c, value, coerce) {
                    if (coerce === void 0) { coerce = false; }
                    var isFormula = wijmo.isString(value) && value.length > 1 && value[0] === '=';
                    this._calcEngine._clearExpressionCache();
                    return this.cells.setCellData(r, c, value, coerce && !isFormula);
                };
                /**
                 * Overrides the base class method to take into account the function list.
                 */
                FlexSheet.prototype.containsFocus = function () {
                    return this.isFunctionListOpen || _super.prototype.containsFocus.call(this);
                };
                /**
                 * Add an unbound @see:Sheet to the <b>FlexSheet</b>.
                 *
                 * @param sheetName The name of the Sheet.
                 * @param rows The row count of the Sheet.
                 * @param cols The column count of the Sheet.
                 * @param pos The position in the <b>sheets</b> collection.
                 * @param grid The @see:FlexGrid instance associated with the @see:Sheet. If not specified then new @see:FlexGrid instance
                 * will be created.
                 */
                FlexSheet.prototype.addUnboundSheet = function (sheetName, rows, cols, pos, grid) {
                    var sheet = this._addSheet(sheetName, rows, cols, pos, grid);
                    if (sheet.selectionRanges.length === 0) {
                        // Store current selection in the selection array for multiple selection.
                        sheet.selectionRanges.push(this.selection);
                    }
                    return sheet;
                };
                /**
                 * Add a bound @see:Sheet to the <b>FlexSheet</b>.
                 *
                 * @param sheetName The name of the @see:Sheet.
                 * @param source The items source for the @see:Sheet.
                 * @param pos The position in the <b>sheets</b> collection.
                 * @param grid The @see:FlexGrid instance associated with the @see:Sheet. If not specified then new @see:FlexGrid instance
                 * will be created.
                 */
                FlexSheet.prototype.addBoundSheet = function (sheetName, source, pos, grid) {
                    var sheet = this._addSheet(sheetName, 0, 0, pos, grid);
                    if (source) {
                        sheet.itemsSource = source;
                    }
                    if (sheet.selectionRanges.length === 0) {
                        // Store current selection in the selection array for multiple selection.
                        sheet.selectionRanges.push(this.selection);
                    }
                    return sheet;
                };
                /**
                 * Apply the style to a range of cells.
                 *
                 * @param cellStyle The @see:ICellStyle object to apply.
                 * @param cells An array of @see:CellRange objects to apply the style to. If not specified then
                 * style is applied to the currently selected cells.
                 * @param isPreview Indicates whether the applied style is just for preview.
                 */
                FlexSheet.prototype.applyCellsStyle = function (cellStyle, cells, isPreview) {
                    if (isPreview === void 0) { isPreview = false; }
                    var rowIndex, colIndex, ranges = cells || [this.selection], range, index, cellStyleAction;
                    if (!this.selectedSheet) {
                        return;
                    }
                    // Cancel current applied style.
                    if (!cellStyle && this._cloneStyle) {
                        this.selectedSheet._styledCells = JSON.parse(JSON.stringify(this._cloneStyle));
                        this._cloneStyle = null;
                        this.refresh(false);
                        return;
                    }
                    // Apply cells style for the cell range of the FlexSheet control.
                    if (ranges) {
                        if (!cells && !isPreview) {
                            cellStyleAction = new sheet_1._CellStyleAction(this, this._cloneStyle);
                            this._cloneStyle = null;
                        }
                        else if (isPreview && !this._cloneStyle) {
                            this._cloneStyle = JSON.parse(JSON.stringify(this.selectedSheet._styledCells));
                        }
                        for (index = 0; index < ranges.length; index++) {
                            range = ranges[index];
                            for (rowIndex = range.topRow; rowIndex <= range.bottomRow; rowIndex++) {
                                for (colIndex = range.leftCol; colIndex <= range.rightCol; colIndex++) {
                                    this._applyStyleForCell(rowIndex, colIndex, cellStyle);
                                }
                            }
                        }
                        if (!cells && !isPreview) {
                            cellStyleAction.saveNewState();
                            this._undoStack._addAction(cellStyleAction);
                        }
                    }
                    if (!cells) {
                        this.refresh(false);
                    }
                };
                /**
                 * Freeze or unfreeze the columns and rows of the <b>FlexSheet</b> control.
                 */
                FlexSheet.prototype.freezeAtCursor = function () {
                    var self = this, rowIndex, colIndex, frozenColumns, frozenRows, row, column;
                    if (!self.selectedSheet) {
                        return;
                    }
                    if (self.selection && self.frozenRows === 0 && self.frozenColumns === 0) {
                        // hide rows\cols scrolled above and scrolled left of the view range
                        // so the user can freeze arbitrary parts of the grid 
                        // (not necessarily starting with the first row/column)
                        if (self._ptScrl.y < 0) {
                            for (rowIndex = 0; rowIndex < self.selection.topRow - 1; rowIndex++) {
                                row = self.rows[rowIndex];
                                if (!(row instanceof HeaderRow)) {
                                    if (row._pos + self._ptScrl.y < 0) {
                                        row.visible = false;
                                    }
                                    else {
                                        self.selectedSheet._freezeHiddenRowCnt = rowIndex;
                                        break;
                                    }
                                }
                            }
                        }
                        if (self._ptScrl.x < 0) {
                            for (colIndex = 0; colIndex < self.selection.leftCol - 1; colIndex++) {
                                column = self.columns[colIndex];
                                if (column._pos + self._ptScrl.x < 0) {
                                    self.columns[colIndex].visible = false;
                                }
                                else {
                                    self.selectedSheet._freezeHiddenColumnCnt = colIndex;
                                    break;
                                }
                            }
                        }
                        // freeze
                        frozenColumns = self.selection.leftCol > 0 ? self.selection.leftCol : 0;
                        frozenRows = self.selection.topRow > 0 ? self.selection.topRow : 0;
                    }
                    else {
                        // unhide
                        for (rowIndex = 0; rowIndex < self.frozenRows - 1; rowIndex++) {
                            self.rows[rowIndex].visible = true;
                        }
                        for (colIndex = 0; colIndex < self.frozenColumns - 1; colIndex++) {
                            self.columns[colIndex].visible = true;
                        }
                        // Apply the filter of the FlexSheet again after resetting the visible of the rows. (TFS 204887)
                        self._filter.apply();
                        // unfreeze
                        frozenColumns = 0;
                        frozenRows = 0;
                        self.selectedSheet._freezeHiddenRowCnt = 0;
                        self.selectedSheet._freezeHiddenColumnCnt = 0;
                    }
                    // Synch to the grid of current sheet.
                    self.frozenRows = self.selectedSheet.grid.frozenRows = frozenRows;
                    self.frozenColumns = self.selectedSheet.grid.frozenColumns = frozenColumns;
                    setTimeout(function () {
                        self.rows._dirty = true;
                        self.columns._dirty = true;
                        self.invalidate();
                        self.scrollIntoView(self.selection.topRow, self.selection.leftCol);
                    }, 10);
                };
                /**
                 * Show the filter editor.
                 */
                FlexSheet.prototype.showColumnFilter = function () {
                    var selectedCol = this.selection.col > 0 ? this.selection.col : 0;
                    if (this.columns.length > 0) {
                        this._filter.editColumnFilter(this.columns[selectedCol]);
                    }
                };
                /**
                 * Clears the content of the <b>FlexSheet</b> control.
                 */
                FlexSheet.prototype.clear = function () {
                    this.selection = new grid_1.CellRange();
                    this.sheets.clear();
                    this._selectedSheetIndex = -1;
                    this.columns.clear();
                    this.rows.clear();
                    this.columnHeaders.columns.clear();
                    this.rowHeaders.rows.clear();
                    this._undoStack.clear();
                    this._ptScrl = new wijmo.Point();
                    this._clearCalcEngine();
                    this.addUnboundSheet();
                };
                /**
                 * Gets the @see:IFormatState object describing formatting of the selected cells.
                 *
                 * @return The @see:IFormatState object containing formatting properties.
                 */
                FlexSheet.prototype.getSelectionFormatState = function () {
                    var rowIndex, colIndex, rowCount = this.rows.length, columnCount = this.columns.length, formatState = {
                        isBold: false,
                        isItalic: false,
                        isUnderline: false,
                        textAlign: 'left',
                        isMergedCell: false
                    };
                    // If there is no rows or columns in the flexsheet, we should return the default format state (TFS 122628)
                    if (rowCount === 0 || columnCount === 0) {
                        return formatState;
                    }
                    // Check the selected cells
                    if (this.selection) {
                        if (this.selection.row >= rowCount || this.selection.row2 >= rowCount
                            || this.selection.col >= columnCount || this.selection.col2 >= columnCount) {
                            return formatState;
                        }
                        for (rowIndex = this.selection.topRow; rowIndex <= this.selection.bottomRow; rowIndex++) {
                            for (colIndex = this.selection.leftCol; colIndex <= this.selection.rightCol; colIndex++) {
                                this._checkCellFormat(rowIndex, colIndex, formatState);
                            }
                        }
                    }
                    return formatState;
                };
                /**
                 * Inserts rows in the current @see:Sheet of the <b>FlexSheet</b> control.
                 *
                 * @param index The position where new rows should be added. If not specified then rows will be added
                 * before the first row of the current selection.
                 * @param count The numbers of rows to add. If not specified then one row will be added.
                 */
                FlexSheet.prototype.insertRows = function (index, count) {
                    var rowIndex = wijmo.isNumber(index) && index >= 0 ? index :
                        (this.selection && this.selection.topRow > -1) ? this.selection.topRow : 0, rowCount = wijmo.isNumber(count) ? count : 1, insRowAction = new sheet_1._RowsChangedAction(this), currentRow = this.rows[rowIndex], i;
                    if (!this.selectedSheet) {
                        return;
                    }
                    // We disable inserting rows manually for the bound sheet.
                    // Because it will cause the synch issue between the itemsSource and the sheet.
                    if (this.itemsSource) {
                        return;
                    }
                    this._clearCalcEngine();
                    this.finishEditing();
                    // The header row of the bound sheet should always in the top of the flexsheet.
                    // The new should be added below the header row. (TFS #124391.)
                    if (rowIndex === 0 && currentRow && currentRow.constructor === HeaderRow) {
                        rowIndex = 1;
                    }
                    // We should update styled cells hash before adding rows.
                    this._updateCellsForUpdatingRow(this.rows.length, rowIndex, rowCount);
                    // Update the affected formulas.
                    insRowAction._affecedFormulas = this._updateAffectedFormula(rowIndex, rowCount, true, true);
                    this.rows.beginUpdate();
                    for (i = 0; i < rowCount; i++) {
                        this.rows.insert(rowIndex, new grid_1.Row());
                    }
                    this.rows.endUpdate();
                    if (!this.selection || this.selection.row === -1 || this.selection.col === -1) {
                        this.selection = new grid_1.CellRange(0, 0);
                    }
                    // Synch with current sheet.
                    this._copyTo(this.selectedSheet);
                    insRowAction.saveNewState();
                    this._undoStack._addAction(insRowAction);
                };
                /**
                 * Deletes rows from the current @see:Sheet of the <b>FlexSheet</b> control.
                 *
                 * @param index The starting index of the deleting rows. If not specified then rows will be deleted
                 * starting from the first row of the current selection.
                 * @param count The numbers of rows to delete. If not specified then one row will be deleted.
                 */
                FlexSheet.prototype.deleteRows = function (index, count) {
                    var rowCount = wijmo.isNumber(count) && count >= 0 ? count :
                        (this.selection && this.selection.topRow > -1) ? this.selection.bottomRow - this.selection.topRow + 1 : 1, firstRowIndex = wijmo.isNumber(index) && index >= 0 ? index :
                        (this.selection && this.selection.topRow > -1) ? this.selection.topRow : -1, lastRowIndex = wijmo.isNumber(index) && index >= 0 ? index + rowCount - 1 :
                        (this.selection && this.selection.topRow > -1) ? this.selection.bottomRow : -1, delRowAction = new sheet_1._RowsChangedAction(this), rowDeleted = false, deletingRow, deletingRowIndex, currentRowsLength;
                    if (!this.selectedSheet) {
                        return;
                    }
                    // We disable deleting rows manually for the bound sheet.
                    // Because it will cause the synch issue between the itemsSource and the sheet.
                    if (this.itemsSource) {
                        return;
                    }
                    this._clearCalcEngine();
                    this.finishEditing();
                    if (firstRowIndex > -1 && lastRowIndex > -1) {
                        // We should update styled cells hash before deleting rows.
                        this._updateCellsForUpdatingRow(this.rows.length, firstRowIndex, rowCount, true);
                        // Update the affected formulas.
                        delRowAction._affecedFormulas = this._updateAffectedFormula(lastRowIndex, lastRowIndex - firstRowIndex + 1, false, true);
                        this.rows.beginUpdate();
                        for (; lastRowIndex >= firstRowIndex; lastRowIndex--) {
                            deletingRow = this.rows[lastRowIndex];
                            // The header row of the bound sheet is a specific row.
                            // So it hasn't to be deleted manually.
                            if (deletingRow && deletingRow.constructor === HeaderRow) {
                                continue;
                            }
                            // if we remove the rows in the bound sheet,
                            // we need remove the row related item in the itemsSource of the flexsheet. (TFS 121651)
                            if (deletingRow.dataItem && this.collectionView) {
                                this.collectionView.beginUpdate();
                                deletingRowIndex = this._getCvIndex(lastRowIndex);
                                if (deletingRowIndex > -1) {
                                    this.itemsSource.splice(lastRowIndex - 1, 1);
                                }
                                this.collectionView.endUpdate();
                            }
                            else {
                                this.rows.removeAt(lastRowIndex);
                            }
                            rowDeleted = true;
                        }
                        this.rows.endUpdate();
                        currentRowsLength = this.rows.length;
                        if (currentRowsLength === 0) {
                            this.selectedSheet.selectionRanges.clear();
                            this.select(new grid_1.CellRange());
                        }
                        else if (lastRowIndex === currentRowsLength - 1) {
                            this.select(new grid_1.CellRange(lastRowIndex, 0, lastRowIndex, this.columns.length - 1));
                        }
                        else {
                            this.select(new grid_1.CellRange(this.selection.topRow, this.selection.col, this.selection.topRow, this.selection.col2));
                        }
                        // Synch with current sheet.
                        this._copyTo(this.selectedSheet);
                        if (rowDeleted) {
                            delRowAction.saveNewState();
                            this._undoStack._addAction(delRowAction);
                        }
                    }
                };
                /**
                 * Inserts columns in the current @see:Sheet of the <b>FlexSheet</b> control.
                 *
                 * @param index The position where new columns should be added. If not specified then columns will be added
                 * before the left column of the current selection.
                 * @param count The numbers of columns to add. If not specified then one column will be added.
                 */
                FlexSheet.prototype.insertColumns = function (index, count) {
                    var columnIndex = wijmo.isNumber(index) && index >= 0 ? index :
                        this.selection && this.selection.leftCol > -1 ? this.selection.leftCol : 0, colCount = wijmo.isNumber(count) ? count : 1, insColumnAction = new sheet_1._ColumnsChangedAction(this), i;
                    if (!this.selectedSheet) {
                        return;
                    }
                    // We disable inserting columns manually for the bound sheet.
                    // Because it will cause the synch issue between the itemsSource and the sheet.
                    if (this.itemsSource) {
                        return;
                    }
                    this._clearCalcEngine();
                    this.finishEditing();
                    // We should update styled cells hash before adding columns.
                    this._updateCellsForUpdatingColumn(this.columns.length, columnIndex, colCount);
                    // Update the affected formulas.
                    insColumnAction._affectedFormulas = this._updateAffectedFormula(columnIndex, colCount, true, false);
                    this.columns.beginUpdate();
                    for (i = 0; i < colCount; i++) {
                        this.columns.insert(columnIndex, new grid_1.Column());
                    }
                    this.columns.endUpdate();
                    if (!this.selection || this.selection.row === -1 || this.selection.col === -1) {
                        this.selection = new grid_1.CellRange(0, 0);
                    }
                    // Synch with current sheet.
                    this._copyTo(this.selectedSheet);
                    insColumnAction.saveNewState();
                    this._undoStack._addAction(insColumnAction);
                };
                /**
                 * Deletes columns from the current @see:Sheet of the <b>FlexSheet</b> control.
                 *
                 * @param index The starting index of the deleting columns. If not specified then columns will be deleted
                 * starting from the first column of the current selection.
                 * @param count The numbers of columns to delete. If not specified then one column will be deleted.
                 */
                FlexSheet.prototype.deleteColumns = function (index, count) {
                    var currentColumnLength, colCount = wijmo.isNumber(count) && count >= 0 ? count :
                        (this.selection && this.selection.leftCol > -1) ? this.selection.rightCol - this.selection.leftCol + 1 : 1, firstColIndex = wijmo.isNumber(index) && index >= 0 ? index :
                        (this.selection && this.selection.leftCol > -1) ? this.selection.leftCol : -1, lastColIndex = wijmo.isNumber(index) && index >= 0 ? index + colCount - 1 :
                        (this.selection && this.selection.leftCol > -1) ? this.selection.rightCol : -1, delColumnAction = new sheet_1._ColumnsChangedAction(this);
                    if (!this.selectedSheet) {
                        return;
                    }
                    // We disable deleting columns manually for the bound sheet.
                    // Because it will cause the synch issue between the itemsSource and the sheet.
                    if (this.itemsSource) {
                        return;
                    }
                    this._clearCalcEngine();
                    this.finishEditing();
                    if (firstColIndex > -1 && lastColIndex > -1) {
                        // We should update styled cells hash before deleting columns.
                        this._updateCellsForUpdatingColumn(this.columns.length, firstColIndex, colCount, true);
                        // Update the affected formulas.
                        delColumnAction._affectedFormulas = this._updateAffectedFormula(lastColIndex, lastColIndex - firstColIndex + 1, false, false);
                        this.columns.beginUpdate();
                        for (; lastColIndex >= firstColIndex; lastColIndex--) {
                            this.columns.removeAt(lastColIndex);
                            this._sortManager.deleteSortLevel(lastColIndex);
                        }
                        this.columns.endUpdate();
                        this._sortManager.commitSort(false);
                        currentColumnLength = this.columns.length;
                        if (currentColumnLength === 0) {
                            this.selectedSheet.selectionRanges.clear();
                            this.select(new grid_1.CellRange());
                        }
                        else if (lastColIndex === currentColumnLength - 1) {
                            this.select(new grid_1.CellRange(0, lastColIndex, this.rows.length - 1, lastColIndex));
                        }
                        else {
                            this.select(new grid_1.CellRange(this.selection.row, this.selection.leftCol, this.selection.row2, this.selection.leftCol));
                        }
                        // Synch with current sheet.
                        this._copyTo(this.selectedSheet);
                        delColumnAction.saveNewState();
                        this._undoStack._addAction(delColumnAction);
                    }
                };
                /**
                 * Merges the selected @see:CellRange into one cell.
                 *
                 * @param cells The @see:CellRange to merge.
                 */
                FlexSheet.prototype.mergeRange = function (cells) {
                    var rowIndex, colIndex, cellIndex, mergedRange, range = cells || this.selection, mergedCellExists = false, cellMergeAction;
                    if (!this.selectedSheet) {
                        return;
                    }
                    if (range) {
                        if (range.rowSpan === 1 && range.columnSpan === 1) {
                            return;
                        }
                        if (!cells) {
                            cellMergeAction = new sheet_1._CellMergeAction(this);
                        }
                        if (!this._resetMergedRange(range)) {
                            for (rowIndex = range.topRow; rowIndex <= range.bottomRow; rowIndex++) {
                                for (colIndex = range.leftCol; colIndex <= range.rightCol; colIndex++) {
                                    cellIndex = rowIndex * this.columns.length + colIndex;
                                    this.selectedSheet._mergedRanges[cellIndex] = new grid_1.CellRange(range.topRow, range.leftCol, range.bottomRow, range.rightCol);
                                }
                            }
                        }
                        if (!cells) {
                            cellMergeAction.saveNewState();
                            this._undoStack._addAction(cellMergeAction);
                        }
                    }
                    if (!cells) {
                        this.refresh();
                    }
                };
                /**
                 * Gets a @see:CellRange that specifies the merged extent of a cell
                 * in a @see:GridPanel.
                 * This method overrides the getMergedRange method of its parent class FlexGrid
                 *
                 * @param panel @see:GridPanel that contains the range.
                 * @param r Index of the row that contains the cell.
                 * @param c Index of the column that contains the cell.
                 * @param clip Whether to clip the merged range to the grid's current view range.
                 * @return A @see:CellRange that specifies the merged range, or null if the cell is not merged.
                 */
                FlexSheet.prototype.getMergedRange = function (panel, r, c, clip) {
                    if (clip === void 0) { clip = true; }
                    var cellIndex = r * this.columns.length + c, mergedRange = this.selectedSheet ? this.selectedSheet._mergedRanges[cellIndex] : null, topRow, bottonRow, leftCol, rightCol;
                    if (panel === this.cells && mergedRange) {
                        // Adjust the merged cell with the frozen pane.
                        if (!mergedRange.isSingleCell && (this.frozenRows > 0 || this.frozenColumns > 0)
                            && ((mergedRange.topRow < this.frozenRows && mergedRange.bottomRow >= this.frozenRows)
                                || (mergedRange.leftCol < this.frozenColumns && mergedRange.rightCol >= this.frozenColumns))) {
                            topRow = mergedRange.topRow;
                            bottonRow = mergedRange.bottomRow;
                            leftCol = mergedRange.leftCol;
                            rightCol = mergedRange.rightCol;
                            if (r >= this.frozenRows && mergedRange.topRow < this.frozenRows) {
                                topRow = this.frozenRows;
                            }
                            if (r < this.frozenRows && mergedRange.bottomRow >= this.frozenRows) {
                                bottonRow = this.frozenRows - 1;
                            }
                            if (bottonRow >= this.rows.length) {
                                bottonRow = this.rows.length - 1;
                            }
                            if (c >= this.frozenColumns && mergedRange.leftCol < this.frozenColumns) {
                                leftCol = this.frozenColumns;
                            }
                            if (c < this.frozenColumns && mergedRange.rightCol >= this.frozenColumns) {
                                rightCol = this.frozenColumns - 1;
                            }
                            if (rightCol >= this.columns.length) {
                                rightCol = this.columns.length - 1;
                            }
                            return new grid_1.CellRange(topRow, leftCol, bottonRow, rightCol);
                        }
                        if (mergedRange.bottomRow >= this.rows.length) {
                            return new grid_1.CellRange(mergedRange.topRow, mergedRange.leftCol, this.rows.length - 1, mergedRange.rightCol);
                        }
                        if (mergedRange.rightCol >= this.columns.length) {
                            return new grid_1.CellRange(mergedRange.topRow, mergedRange.leftCol, mergedRange.bottomRow, this.columns.length - 1);
                        }
                        return mergedRange.clone();
                    }
                    // Only when there are columns in current sheet, it will get the merge range from parent flexgrid. (TFS #142348, #143544)
                    if (c >= 0 && this.columns && this.columns.length > c && r >= 0 && this.rows && this.rows.length > c) {
                        return _super.prototype.getMergedRange.call(this, panel, r, c, clip);
                    }
                    return null;
                };
                /**
                 * Evaluates a formula.
                 *
                 * @see:FlexSheet formulas follow the Excel syntax, including a large subset of the
                 * functions supported by Excel. A complete list of the functions supported by
                 * @see:FlexSheet can be found here:
                 * <a href="static/FlexSheetFunctions.html">FlexSheet Functions</a>.
                 *
                 * @param formula The formula to evaluate. The formula may start with an optional equals sign ('=').
                 * @param format If specified, defines the .Net format that will be applied to the evaluated value.
                 * @param sheet The @see:Sheet whose data will be used for evaluation.
                 *              If not specified then the current sheet is used.
                 */
                FlexSheet.prototype.evaluate = function (formula, format, sheet) {
                    return this._evaluate(formula, format, sheet);
                };
                /**
                 * Gets the evaluated cell value.
                 *
                 * Unlike the <b>getCellData</b> method that returns a raw data that can be a value or a formula, the <b>getCellValue</b>
                 * method always returns an evaluated value, that is if the cell contains a formula then it will be evaluated first and the
                 * resulting value will be returned.
                 *
                 * @param rowIndex The row index of the cell.
                 * @param colIndex The column index of the cell.
                 * @param formatted Indicates whether to return an original or a formatted value of the cell.
                 * @param sheet The @see:Sheet whose value to evaluate. If not specified then the data from current sheet
                 * is used.
                 */
                FlexSheet.prototype.getCellValue = function (rowIndex, colIndex, formatted, sheet) {
                    if (formatted === void 0) { formatted = false; }
                    var col = this.columns[colIndex], cellIndex = rowIndex * this.columns.length + colIndex, styleInfo, format, cellVal;
                    styleInfo = sheet ? sheet._styledCells[cellIndex] : (this.selectedSheet ? this.selectedSheet._styledCells[cellIndex] : null);
                    format = styleInfo && styleInfo.format ? styleInfo.format : '';
                    cellVal = sheet ? sheet.grid.getCellData(rowIndex, colIndex, false) : this.getCellData(rowIndex, colIndex, false);
                    if (wijmo.isString(cellVal) && cellVal[0] === '=') {
                        cellVal = this._evaluate(cellVal, formatted ? format : '', sheet, rowIndex, colIndex);
                    }
                    if (wijmo.isPrimitive(cellVal)) {
                        if (formatted) {
                            if (col.dataMap) {
                                cellVal = col.dataMap.getDisplayValue(cellVal);
                            }
                            cellVal = cellVal != null ? wijmo.Globalize.format(cellVal, format || col.format) : '';
                        }
                    }
                    else if (cellVal) {
                        if (formatted) {
                            cellVal = wijmo.Globalize.format(cellVal.value, format || cellVal.format || col.format);
                        }
                        else {
                            cellVal = cellVal.value;
                        }
                    }
                    return cellVal == null ? '' : cellVal;
                };
                /**
                 * Open the function list.
                 *
                 * @param target The DOM element that toggle the function list.
                 */
                FlexSheet.prototype.showFunctionList = function (target) {
                    var self = this, functionOffset = self._cumulativeOffset(target), rootOffset = self._cumulativeOffset(self['_root']), offsetTop, offsetLeft;
                    self._functionTarget = wijmo.tryCast(target, HTMLInputElement);
                    if (self._functionTarget && self._functionTarget.value && self._functionTarget.value[0] === '=') {
                        self._functionList._cv.filter = function (item) {
                            var text = item['actualvalue'].toLowerCase(), searchIndex = self._getCurrentFormulaIndex(self._functionTarget.value), searchText;
                            if (searchIndex === -1) {
                                searchIndex = 0;
                            }
                            searchText = self._functionTarget.value.substr(searchIndex + 1).trim().toLowerCase();
                            if ((searchText.length > 0 && text.indexOf(searchText) === 0) || self._functionTarget.value === '=') {
                                return true;
                            }
                            return false;
                        };
                        self._functionList.selectedIndex = 0;
                        offsetTop = functionOffset.y + target.clientHeight + 2 + (wijmo.hasClass(target, 'wj-grid-editor') ? this._ptScrl.y : 0);
                        offsetLeft = functionOffset.x + (wijmo.hasClass(target, 'wj-grid-editor') ? this._ptScrl.x : 0);
                        wijmo.setCss(self._functionListHost, {
                            height: self._functionList._cv.items.length > 5 ? '218px' : 'auto',
                            display: self._functionList._cv.items.length > 0 ? 'block' : 'none',
                            top: '',
                            left: ''
                        });
                        self._functionListHost.scrollTop = 0;
                        if (self._functionListHost.offsetHeight + offsetTop > rootOffset.y + self['_root'].offsetHeight) {
                            offsetTop = offsetTop - target.clientHeight - self._functionListHost.offsetHeight - 5;
                        }
                        else {
                            offsetTop += 5;
                        }
                        if (self._functionListHost.offsetWidth + offsetLeft > rootOffset.x + self['_root'].offsetWidth) {
                            offsetLeft = rootOffset.x + self['_root'].offsetWidth - self._functionListHost.offsetWidth;
                        }
                        wijmo.setCss(self._functionListHost, {
                            top: offsetTop,
                            left: offsetLeft
                        });
                    }
                    else {
                        self.hideFunctionList();
                    }
                };
                /**
                 * Close the function list.
                 */
                FlexSheet.prototype.hideFunctionList = function () {
                    this._functionListHost.style.display = 'none';
                };
                /**
                 * Select previous function in the function list.
                 */
                FlexSheet.prototype.selectPreviousFunction = function () {
                    var index = this._functionList.selectedIndex;
                    if (index > 0) {
                        this._functionList.selectedIndex--;
                    }
                };
                /**
                 * Select next function in the function list.
                 */
                FlexSheet.prototype.selectNextFunction = function () {
                    var index = this._functionList.selectedIndex;
                    if (index < this._functionList.itemsSource.length) {
                        this._functionList.selectedIndex++;
                    }
                };
                /**
                 * Inserts the selected function from the function list to the cell value editor.
                 */
                FlexSheet.prototype.applyFunctionToCell = function () {
                    var self = this, currentFormulaIndex;
                    if (self._functionTarget) {
                        currentFormulaIndex = self._getCurrentFormulaIndex(self._functionTarget.value);
                        if (currentFormulaIndex === -1) {
                            currentFormulaIndex = self._functionTarget.value.indexOf('=');
                        }
                        else {
                            currentFormulaIndex += 1;
                        }
                        self._functionTarget.value = self._functionTarget.value.substring(0, currentFormulaIndex) + self._functionList.selectedValue + '(';
                        if (self._functionTarget.value[0] !== '=') {
                            self._functionTarget.value = '=' + self._functionTarget.value;
                        }
                        self._functionTarget.focus();
                        self.hideFunctionList();
                    }
                };
                /**
                 * Saves the <b>FlexSheet</b> to xlsx file.
                 *
                 * For example:
                 * <pre>// This sample exports FlexSheet content to an xlsx
                 * // click.
                 * &nbsp;
                 * // HTML
                 * &lt;button
                 *     onclick="saveXlsx('FlexSheet.xlsx')"&gt;
                 *     Save
                 * &lt;/button&gt;
                 * &nbsp;
                 * // JavaScript
                 * function saveXlsx(fileName) {
                 *     // Save the flexGrid to xlsx file.
                 *     flexsheet.save(fileName);
                 * }</pre>
                 *
                 * @param fileName Name of the file that will be generated.
                 * @return A workbook instance containing the generated xlsx file content.
                 */
                FlexSheet.prototype.save = function (fileName) {
                    var workbook = this._saveToWorkbook();
                    if (fileName) {
                        workbook.save(fileName);
                    }
                    return workbook;
                };
                /*
                 * Save the <b>FlexSheet</b> to Workbook Object Model represented by the @see:IWorkbook interface.
                 *
                 * @return The @see:IWorkbook instance representing export results.
                 */
                FlexSheet.prototype.saveToWorkbookOM = function () {
                    var workbook = this._saveToWorkbook();
                    return workbook._serialize();
                };
                /**
                 * Loads the workbook into the <b>FlexSheet</b>.
                 *
                 * For example:
                 * <pre>// This sample opens an xlsx file chosen via Open File
                 * // dialog and fills FlexSheet
                 * &nbsp;
                 * // HTML
                 * &lt;input type="file"
                 *     id="importFile"
                 *     accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                 * /&gt;
                 * &lt;div id="flexHost"&gt;&lt;/&gt;
                 * &nbsp;
                 * // JavaScript
                 * var flexSheet = new wijmo.grid.FlexSheet("#flexHost"),
                 *     importFile = document.getElementById('importFile');
                 * &nbsp;
                 * importFile.addEventListener('change', function () {
                 *     loadWorkbook();
                 * });
                 * &nbsp;
                 * function loadWorkbook() {
                 *     var reader,
                 *         file = importFile.files[0];
                 *     if (file) {
                 *         reader = new FileReader();
                 *         reader.onload = function (e) {
                 *             flexSheet.load(reader.result);
                 *         };
                 *         reader.readAsArrayBuffer(file);
                 *     }
                 * }</pre>
                 *
                 * @param workbook An Workbook instance or a Blob instance or a base 64 stirng or an ArrayBuffer containing xlsx file content.
                 */
                FlexSheet.prototype.load = function (workbook) {
                    var workbookInstance, reader, self = this;
                    if (workbook instanceof Blob) {
                        reader = new FileReader();
                        reader.onload = function () {
                            var fileContent = reader.result;
                            fileContent = wijmo.xlsx.Workbook._base64EncArr(new Uint8Array(fileContent));
                            workbookInstance = new wijmo.xlsx.Workbook();
                            workbookInstance.load(fileContent);
                            self._loadFromWorkbook(workbookInstance);
                        };
                        reader.readAsArrayBuffer(workbook);
                    }
                    else if (workbook instanceof wijmo.xlsx.Workbook) {
                        self._loadFromWorkbook(workbook);
                    }
                    else {
                        if (workbook instanceof ArrayBuffer) {
                            workbook = wijmo.xlsx.Workbook._base64EncArr(new Uint8Array(workbook));
                        }
                        else if (!wijmo.isString(workbook)) {
                            throw 'Invalid workbook.';
                        }
                        workbookInstance = new wijmo.xlsx.Workbook();
                        workbookInstance.load(workbook);
                        self._loadFromWorkbook(workbookInstance);
                    }
                };
                /*
                 * Load the Workbook Object Model instance into the <b>FlexSheet</b>.
                 *
                 * @param workbook The Workbook Object Model instance to load data from.
                 */
                FlexSheet.prototype.loadFromWorkbookOM = function (workbook) {
                    var grids = [], workbookInstance;
                    if (workbook instanceof wijmo.xlsx.Workbook) {
                        workbookInstance = workbook;
                    }
                    else {
                        workbookInstance = new wijmo.xlsx.Workbook();
                        workbookInstance._deserialize(workbook);
                    }
                    this._loadFromWorkbook(workbookInstance);
                };
                /**
                 * Undo the last user action.
                 */
                FlexSheet.prototype.undo = function () {
                    var self = this;
                    // The undo should wait until other operations have done. (TFS 189582) 
                    setTimeout(function () {
                        self._undoStack.undo();
                    }, 100);
                };
                /**
                 * Redo the last user action.
                 */
                FlexSheet.prototype.redo = function () {
                    var self = this;
                    // The redo should wait until other operations have done. (TFS 189582) 
                    setTimeout(function () {
                        self._undoStack.redo();
                    }, 100);
                };
                /**
                 * Selects a cell range and optionally scrolls it into view.
                 *
                 * @see:FlexSheet overrides this method to adjust the selection cell range for the merged cells in the @see:FlexSheet.
                 *
                 * @param rng The cell range to select.
                 * @param show Indicates whether to scroll the new selection into view.
                 */
                FlexSheet.prototype.select = function (rng, show) {
                    if (show === void 0) { show = true; }
                    var mergedRange, rowIndex, colIndex;
                    if (rng.rowSpan !== this.rows.length && rng.columnSpan !== this.columns.length) {
                        for (rowIndex = rng.topRow; rowIndex <= rng.bottomRow; rowIndex++) {
                            for (colIndex = rng.leftCol; colIndex <= rng.rightCol; colIndex++) {
                                mergedRange = this.getMergedRange(this.cells, rowIndex, colIndex);
                                if (mergedRange && !rng.equals(mergedRange)) {
                                    if (rng.row <= rng.row2) {
                                        rng.row = Math.min(rng.topRow, mergedRange.topRow);
                                        rng.row2 = Math.max(rng.bottomRow, mergedRange.bottomRow);
                                    }
                                    else {
                                        rng.row = Math.max(rng.bottomRow, mergedRange.bottomRow);
                                        rng.row2 = Math.min(rng.topRow, mergedRange.topRow);
                                    }
                                    if (rng.col <= rng.col2) {
                                        rng.col = Math.min(rng.leftCol, mergedRange.leftCol);
                                        rng.col2 = Math.max(rng.rightCol, mergedRange.rightCol);
                                    }
                                    else {
                                        rng.col = Math.max(rng.rightCol, mergedRange.rightCol);
                                        rng.col2 = Math.min(rng.leftCol, mergedRange.leftCol);
                                    }
                                }
                            }
                        }
                    }
                    if (this.collectionView) {
                        // When select all cells in the bound sheet, we should ignore the header row of the bound sheet.
                        // This updating is for TFS issue #128358
                        if (rng.topRow === 0 && rng.bottomRow === this.rows.length - 1
                            && rng.leftCol === 0 && rng.rightCol === this.columns.length - 1) {
                            rng.row = 1;
                            rng.row2 = this.rows.length - 1;
                        }
                    }
                    _super.prototype.select.call(this, rng, show);
                };
                /**
                 * Add custom function in @see:FlexSheet.
                 * @param name the name of the custom function.
                 * @param func the custom function.
                 * @param description the description of the custom function, it will be shown in the function autocompletion of the @see:FlexSheet.
                 * @param minParamsCount the minimum count of the parameter that the function need.
                 * @param maxParamsCount the maximum count of the parameter that the function need.
                 *        If the count of the parameters in the custom function is arbitrary, the minParamsCount and maxParamsCount should be set to null.
                 */
                FlexSheet.prototype.addCustomFunction = function (name, func, description, minParamsCount, maxParamsCount) {
                    this._calcEngine.addCustomFunction(name, func, minParamsCount, maxParamsCount);
                    this._addCustomFunctionDescription(name, description);
                };
                /**
                 * Disposes of the control by removing its association with the host element.
                 */
                FlexSheet.prototype.dispose = function () {
                    var userAgent = window.navigator.userAgent;
                    document.removeEventListener('mousemove', this._mouseMoveHdl);
                    document.body.removeEventListener('click', this._clickHdl);
                    if (userAgent.match(/iPad/i) || userAgent.match(/iPhone/i)) {
                        document.body.removeEventListener('touchstart', this._touchStartHdl);
                        document.body.removeEventListener('touchend', this._touchEndHdl);
                    }
                    _super.prototype.dispose.call(this);
                };
                /**
                 * Parses a string into rows and columns and applies the content to a given range.
                 *
                 * Override the <b>setClipString</b> method of @see:FlexGrid.
                 *
                 * @param text Tab and newline delimited text to parse into the grid.
                 * @param rng @see:CellRange to copy. If omitted, the current selection is used.
                 */
                FlexSheet.prototype.setClipString = function (text, rng) {
                    var autoRange = rng == null, pasted = false, rngPaste, row, copiedRow, copiedCol, col, lines, cells, cellData, matches, isUpdated, i, cellRefIndex, cellRef, cellAddress, updatedCellRef, rowDiff, colDiff, e;
                    if (!this._copiedRange) {
                        _super.prototype.setClipString.call(this, text, rng);
                        return;
                    }
                    rng = rng ? wijmo.asType(rng, grid_1.CellRange) : this.selection;
                    // normalize text
                    text = wijmo.asString(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
                    if (text && text[text.length - 1] == '\n') {
                        text = text.substring(0, text.length - 1);
                    }
                    if (autoRange && !rng.isSingleCell) {
                        text = this._expandClipString(text, rng);
                    }
                    // keep track of paste range to select later
                    rngPaste = new grid_1.CellRange(rng.topRow, rng.leftCol);
                    // copy lines to rows
                    this.beginUpdate();
                    row = rng.topRow;
                    copiedRow = this._copiedRange.topRow;
                    rowDiff = row - copiedRow;
                    lines = text.split('\n');
                    for (i = 0; i < lines.length && row < this.rows.length; i++, row++) {
                        // skip invisible row, keep clip line
                        if (!this.rows[row].isVisible) {
                            i--;
                            continue;
                        }
                        // copy cells to columns
                        cells = lines[i].split('\t');
                        copiedCol = this._copiedRange.leftCol;
                        col = rng.leftCol;
                        colDiff = col - copiedCol;
                        for (var j = 0; j < cells.length && col < this.columns.length; j++, col++) {
                            // skip invisible column, keep clip cell
                            if (!this.columns[col].isVisible) {
                                j--;
                                continue;
                            }
                            // assign cell
                            if (!this.columns[col].isReadOnly && !this.rows[row].isReadOnly) {
                                cellData = cells[j];
                                if (!!cellData && typeof cellData === 'string' && cellData[0] === '=' && (rowDiff !== 0 || colDiff !== 0)) {
                                    matches = cellData.match(/(?=\b\D)\$?[A-Za-z]+\$?\d+/g);
                                    if (!!matches && matches.length > 0) {
                                        for (cellRefIndex = 0; cellRefIndex < matches.length; cellRefIndex++) {
                                            cellRef = matches[cellRefIndex];
                                            if (cellRef.toLowerCase() !== 'atan2') {
                                                cellAddress = wijmo.xlsx.Workbook.tableAddress(cellRef);
                                                cellAddress.row += rowDiff;
                                                cellAddress.col += colDiff;
                                                updatedCellRef = wijmo.xlsx.Workbook.xlsxAddress(cellAddress.row, cellAddress.col, cellAddress.absRow, cellAddress.absCol);
                                                cellData = cellData.replace(cellRef, updatedCellRef);
                                            }
                                        }
                                    }
                                }
                                // raise events so user can cancel the paste
                                e = new grid_1.CellRangeEventArgs(this.cells, new grid_1.CellRange(row, col), cellData);
                                if (this.onPastingCell(e)) {
                                    if (this.cells.setCellData(row, col, cellData)) {
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
                    // done, refresh view to update sorting/filtering 
                    if (this.collectionView && pasted) {
                        this.collectionView.refresh();
                    }
                    // select pasted range
                    this.select(rngPaste);
                };
                // Override the getCvIndex method of its parent class FlexGrid
                FlexSheet.prototype._getCvIndex = function (index) {
                    var row;
                    if (index > -1 && this.collectionView) {
                        row = this.rows[index];
                        if (row instanceof HeaderRow) {
                            return index;
                        }
                        if (row.dataItem) {
                            return _super.prototype._getCvIndex.call(this, index);
                        }
                        return this.collectionView.currentPosition;
                    }
                    return -1;
                };
                // Initialize the FlexSheet control
                FlexSheet.prototype._init = function () {
                    var _this = this;
                    var self = this, userAgent = window.navigator.userAgent, mouseUp = function (e) {
                        document.removeEventListener('mouseup', mouseUp);
                        self._mouseUp(e);
                    };
                    self._divContainer = self.hostElement.querySelector('[wj-part="container"]');
                    self._tabHolder = new sheet_1._TabHolder(self.hostElement.querySelector('[wj-part="tab-holder"]'), self);
                    self._contextMenu = new sheet_1._ContextMenu(self.hostElement.querySelector('[wj-part="context-menu"]'), self);
                    self['_gpCells'] = new FlexSheetPanel(self, grid_1.CellType.Cell, self.rows, self.columns, self['_eCt']);
                    self['_gpCHdr'] = new FlexSheetPanel(self, grid_1.CellType.ColumnHeader, self['_hdrRows'], self.columns, self['_eCHdrCt']);
                    self['_gpRHdr'] = new FlexSheetPanel(self, grid_1.CellType.RowHeader, self.rows, self['_hdrCols'], self['_eRHdrCt']);
                    self['_gpTL'] = new FlexSheetPanel(self, grid_1.CellType.TopLeft, self['_hdrRows'], self['_hdrCols'], self['_eTLCt']);
                    self._sortManager = new sheet_1.SortManager(self);
                    self._filter = new sheet_1._FlexSheetFilter(self);
                    self._filter.filterApplied.addHandler(function () {
                        if (self._wholeColumnsSelected) {
                            self.selection = new grid_1.CellRange(self.selection.topRow, self.selection.col, self.rows.length - 1, self.selection.col2);
                        }
                    });
                    self._calcEngine = new sheet_1._CalcEngine(self);
                    self._calcEngine.unknownFunction.addHandler(function (sender, e) {
                        self.onUnknownFunction(e);
                    }, self);
                    self._initFuncsList();
                    self._undoStack = new sheet_1.UndoStack(self);
                    // Add header row for the bind sheet.
                    self.loadedRows.addHandler(function () {
                        if (self.itemsSource && !(self.rows[0] instanceof HeaderRow)) {
                            self.rows.insert(0, new HeaderRow());
                        }
                    });
                    // Setting the required property of the column to false for the bound sheet.
                    // TFS #126125
                    self.itemsSourceChanged.addHandler(function () {
                        var colIndex;
                        for (colIndex = 0; colIndex < self.columns.length; colIndex++) {
                            self.columns[colIndex].isRequired = false;
                        }
                    });
                    // Store the copied range for updating cell reference of the formula when pasting. (TFS 190785)
                    self.copied.addHandler(function (sender, args) {
                        self._copiedRange = args.range;
                    });
                    // If the rows\columns of FlexSheet were cleared, we should reset merged cells, styled cells and selection of current sheet to null. (TFS 140344)
                    self.rows.collectionChanged.addHandler(function (sender, e) {
                        self._clearForEmptySheet('rows');
                    }, self);
                    self.columns.collectionChanged.addHandler(function (sender, e) {
                        self._clearForEmptySheet('columns');
                    }, self);
                    self.addEventListener(self.hostElement, 'mousedown', function (e) {
                        document.addEventListener('mouseup', mouseUp);
                        // Only when the target is the child of the root container of the FlexSheet control, 
                        // it will deal with the mouse down event handler of the FlexSheet control. (TFS 152995)
                        if (self._isDescendant(self._divContainer, e.target)) {
                            self._mouseDown(e);
                        }
                    }, true);
                    self.addEventListener(self.hostElement, 'drop', function () {
                        self._columnHeaderClicked = false;
                    });
                    self.addEventListener(self.hostElement, 'contextmenu', function (e) {
                        var ht, selectedRow, selectedCol, colPos, rowPos, point, newSelection;
                        if (e.defaultPrevented) {
                            return;
                        }
                        if (!self._edtHdl.activeEditor) {
                            // Handle the hitTest for the keyboard context menu event in IE
                            // Since it can't get the correct position for the keyboard context menu event in IE (TFS 122943)
                            if (e.pageX === 0 && e.pageY === 0
                                && self.selection.row > -1 && self.selection.col > -1
                                && self.rows.length > 0 && self.columns.length > 0) {
                                selectedCol = self.columns[self.selection.col];
                                selectedRow = self.rows[self.selection.row];
                                colPos = selectedCol.pos + self.hostElement.offsetLeft + _this._ptScrl.x;
                                rowPos = selectedRow.pos + self.hostElement.offsetTop + _this._ptScrl.y;
                                point = new wijmo.Point(colPos + selectedCol.renderSize, rowPos + selectedRow.renderSize);
                                ht = self.hitTest(colPos, rowPos);
                            }
                            else {
                                ht = self.hitTest(e);
                            }
                            e.preventDefault();
                            if (ht && ht.cellType !== grid_1.CellType.None) {
                                // Disable add\remove rows\columns for bound sheet.
                                if (!_this.itemsSource) {
                                    self._contextMenu.show(e, point);
                                }
                                newSelection = new grid_1.CellRange(ht.row, ht.col);
                                if (ht.cellType === grid_1.CellType.Cell && !newSelection.intersects(self.selection)) {
                                    if (self.selectedSheet) {
                                        self.selectedSheet.selectionRanges.clear();
                                    }
                                    self.selection = newSelection;
                                    self.selectedSheet.selectionRanges.push(newSelection);
                                }
                            }
                        }
                    });
                    self.prepareCellForEdit.addHandler(self._prepareCellForEditHandler, self);
                    self.cellEditEnded.addHandler(function () {
                        setTimeout(function () {
                            self.hideFunctionList();
                        }, 200);
                    });
                    self.cellEditEnding.addHandler(function () {
                        self._clearCalcEngine();
                    });
                    self.pasted.addHandler(function () {
                        self._clearCalcEngine();
                    });
                    self.addEventListener(self.hostElement, 'keydown', function (e) {
                        var selectionCnt, args, text;
                        if (e.ctrlKey) {
                            if (e.keyCode === 89) {
                                self.finishEditing();
                                self.redo();
                                e.preventDefault();
                            }
                            if (e.keyCode === 90) {
                                self.finishEditing();
                                self.undo();
                                e.preventDefault();
                            }
                            if (!!self.selectedSheet && e.keyCode === 65) {
                                self.selectedSheet.selectionRanges.clear();
                                self.selectedSheet.selectionRanges.push(self.selection);
                            }
                            // Processing for 'Cut' operation. (TFS 191694)
                            if (e.keyCode === 88) {
                                self.finishEditing();
                                args = new grid_1.CellRangeEventArgs(self.cells, self.selection);
                                if (self.onCopying(args)) {
                                    text = self.getClipString();
                                    wijmo.Clipboard.copy(text);
                                    self.deferUpdate(function () {
                                        var row, col, bcol, contentDeleted = false, delAction = new sheet_1._EditAction(self);
                                        for (row = self.selection.topRow; row <= self.selection.bottomRow; row++) {
                                            for (col = self.selection.leftCol; col <= self.selection.rightCol; col++) {
                                                bcol = self._getBindingColumn(self.cells, row, self.columns[col]);
                                                if (bcol.isRequired == false || (bcol.isRequired == null && bcol.dataType == wijmo.DataType.String)) {
                                                    if (self.getCellData(row, col, true)) {
                                                        self.setCellData(row, col, '', true);
                                                        contentDeleted = true;
                                                    }
                                                }
                                            }
                                        }
                                        if (contentDeleted) {
                                            delAction.saveNewState();
                                            self._undoStack._addAction(delAction);
                                        }
                                    });
                                    self.onCopied(args);
                                }
                                e.stopPropagation();
                                return;
                            }
                        }
                        // When press 'Esc' key, we should hide the context menu (TFS 122527)
                        if (e.keyCode === wijmo.Key.Escape) {
                            self._contextMenu.hide();
                            e.preventDefault();
                        }
                        if (e.keyCode === wijmo.Key.Delete && !self._edtHdl.activeEditor) {
                            self._delSeletionContent();
                            e.preventDefault();
                        }
                        if (!!self.selectedSheet && (e.keyCode === wijmo.Key.Left || e.keyCode === wijmo.Key.Right
                            || e.keyCode === wijmo.Key.Up || e.keyCode === wijmo.Key.Down
                            || e.keyCode === wijmo.Key.PageUp || e.keyCode === wijmo.Key.PageDown
                            || e.keyCode === wijmo.Key.Home || e.keyCode === wijmo.Key.End
                            || e.keyCode === wijmo.Key.Tab || e.keyCode === wijmo.Key.Enter)) {
                            selectionCnt = self.selectedSheet.selectionRanges.length;
                            if (selectionCnt > 0) {
                                self.selectedSheet.selectionRanges[selectionCnt - 1] = self.selection;
                            }
                        }
                    });
                    document.body.addEventListener('click', self._clickHdl);
                    document.addEventListener('mousemove', self._mouseMoveHdl);
                    // Show/hide the customize context menu for iPad or iPhone 
                    if (userAgent.match(/iPad/i) || userAgent.match(/iPhone/i)) {
                        document.body.addEventListener('touchstart', self._touchStartHdl);
                        document.body.addEventListener('touchend', self._touchEndHdl);
                    }
                    // After dropping in flexsheet, the flexsheet._htDown should be reset to null. (TFS #142369)
                    self.addEventListener(self.hostElement, 'drop', function () {
                        self._htDown = null;
                    });
                };
                // initialize the function autocomplete list
                FlexSheet.prototype._initFuncsList = function () {
                    var self = this;
                    self._functionListHost = document.createElement('div');
                    wijmo.addClass(self._functionListHost, 'wj-flexsheet-formula-list');
                    document.querySelector('body').appendChild(self._functionListHost);
                    self._functionListHost.style.display = 'none';
                    self._functionListHost.style.position = 'absolute';
                    self._functionList = new wijmo.input.ListBox(self._functionListHost);
                    self._functionList.isContentHtml = true;
                    self._functionList.itemsSource = self._getFunctions();
                    self._functionList.displayMemberPath = 'displayValue';
                    self._functionList.selectedValuePath = 'actualvalue';
                    self.addEventListener(self._functionListHost, 'click', self.applyFunctionToCell.bind(self));
                    self.addEventListener(self._functionListHost, 'keydown', function (e) {
                        // When press 'Esc' key in the host element of the function list,
                        // the function list should be hidden and make the host element of the flexsheet get focus. (TFS 142370)
                        if (e.keyCode === wijmo.Key.Escape) {
                            self.hideFunctionList();
                            self.hostElement.focus();
                            e.preventDefault();
                        }
                        // When press 'Enter' key in the host element of the function list,
                        // the selected function of the function should be applied to the selected cell
                        // and make the host element of the flexsheet get focus.
                        if (e.keyCode === wijmo.Key.Enter) {
                            self.applyFunctionToCell();
                            self.hostElement.focus();
                            e.preventDefault();
                        }
                    });
                };
                // Organize the functions data the function list box
                FlexSheet.prototype._getFunctions = function () {
                    var functions = [], i = 0, func;
                    for (; i < FlexSheetFunctions.length; i++) {
                        func = FlexSheetFunctions[i];
                        functions.push({
                            displayValue: '<div class="wj-flexsheet-formula-name">' + func.name + '</div><div class="wj-flexsheet-formula-description">' + func.description + '</div>',
                            actualvalue: func.name
                        });
                    }
                    return functions;
                };
                // Add the description of the custom function in flexsheet.
                FlexSheet.prototype._addCustomFunctionDescription = function (name, description) {
                    var customFuncDesc = {
                        displayValue: '<div class="wj-flexsheet-formula-name">' + name + '</div>' + (description ? '<div class="wj-flexsheet-formula-description">' + description + '</div>' : ''),
                        actualvalue: name
                    }, funcList = this._functionList.itemsSource, funcIndex = -1, i = 0, funcDesc;
                    for (; i < funcList.length; i++) {
                        funcDesc = funcList[i];
                        if (funcDesc.actualvalue === name) {
                            funcIndex = i;
                            break;
                        }
                    }
                    if (funcIndex > -1) {
                        funcList.splice(funcIndex, 1, customFuncDesc);
                    }
                    else {
                        funcList.push(customFuncDesc);
                    }
                };
                // Get current processing formula index.
                FlexSheet.prototype._getCurrentFormulaIndex = function (searchText) {
                    var searchIndex = -1;
                    ['+', '-', '*', '/', '^', '(', '&'].forEach(function (val) {
                        var index = searchText.lastIndexOf(val);
                        if (index > searchIndex) {
                            searchIndex = index;
                        }
                    });
                    return searchIndex;
                };
                // Prepare cell for edit event handler.
                // This event handler will attach keydown, keyup and blur event handler for the edit cell.
                FlexSheet.prototype._prepareCellForEditHandler = function () {
                    var self = this, edt = self._edtHdl._edt;
                    if (!edt) {
                        return;
                    }
                    // bind keydown event handler for the edit cell.
                    self.addEventListener(edt, 'keydown', function (e) {
                        if (self.isFunctionListOpen) {
                            switch (e.keyCode) {
                                case wijmo.Key.Up:
                                    self.selectPreviousFunction();
                                    e.preventDefault();
                                    break;
                                case wijmo.Key.Down:
                                    self.selectNextFunction();
                                    e.preventDefault();
                                    break;
                                case wijmo.Key.Tab:
                                case wijmo.Key.Enter:
                                    self.applyFunctionToCell();
                                    e.preventDefault();
                                    break;
                                case wijmo.Key.Escape:
                                    self.hideFunctionList();
                                    e.preventDefault();
                                    break;
                            }
                        }
                    });
                    // bind the keyup event handler for the edit cell.
                    self.addEventListener(edt, 'keyup', function (e) {
                        if ((e.keyCode > 40 || e.keyCode < 32) && e.keyCode !== wijmo.Key.Tab && e.keyCode !== wijmo.Key.Escape) {
                            setTimeout(function () {
                                self.showFunctionList(edt);
                            }, 0);
                        }
                    });
                };
                // Add new sheet into the flexsheet.
                FlexSheet.prototype._addSheet = function (sheetName, rows, cols, pos, grid) {
                    var sheet = new sheet_1.Sheet(this, grid, sheetName, rows, cols);
                    if (!this.sheets.isValidSheetName(sheet)) {
                        sheet._setValidName(this.sheets.getValidSheetName(sheet));
                    }
                    if (typeof (pos) === 'number') {
                        if (pos < 0) {
                            pos = 0;
                        }
                        if (pos >= this.sheets.length) {
                            pos = this.sheets.length;
                        }
                    }
                    else {
                        pos = this.sheets.length;
                    }
                    this.sheets.insert(pos, sheet);
                    // If the new sheet is added before current selected sheet, we should adjust the index of current selected sheet. (TFS 143291)
                    if (pos <= this._selectedSheetIndex) {
                        this._selectedSheetIndex += 1;
                    }
                    this.selectedSheetIndex = pos;
                    return sheet;
                };
                // Show specific sheet in the FlexSheet.
                FlexSheet.prototype._showSheet = function (index) {
                    var oldSheet, newSheet;
                    if (!this.sheets || !this.sheets.length || index >= this.sheets.length
                        || index < 0 || index === this.selectedSheetIndex
                        || (this.sheets[index] && !this.sheets[index].visible)) {
                        return;
                    }
                    // finish any pending edits in the old sheet data.
                    this.finishEditing();
                    // save the old sheet data
                    if (this.selectedSheetIndex > -1 && this.selectedSheetIndex < this.sheets.length) {
                        this._copyTo(this.sheets[this.selectedSheetIndex]);
                        this._resetFilterDefinition();
                    }
                    // show the new sheet data
                    if (this.sheets[index]) {
                        this._selectedSheetIndex = index;
                        this._copyFrom(this.sheets[index]);
                    }
                    this._filter.closeEditor();
                };
                // Current sheet changed event handler.
                FlexSheet.prototype._selectedSheetChange = function (sender, e) {
                    this._showSheet(e.newValue);
                    this.invalidate(true);
                    this.onSelectedSheetChanged(e);
                };
                // SheetCollection changed event handler.
                FlexSheet.prototype._sourceChange = function (sender, e) {
                    var item;
                    if (e.action === wijmo.collections.NotifyCollectionChangedAction.Add || e.action === wijmo.collections.NotifyCollectionChangedAction.Change) {
                        item = e.item;
                        item._attachOwner(this);
                        if (e.action === wijmo.collections.NotifyCollectionChangedAction.Add) {
                            this._addingSheet = true;
                            if (e.index <= this.selectedSheetIndex) {
                                this._selectedSheetIndex += 1;
                            }
                        }
                        else {
                            if (e.index === this.selectedSheetIndex) {
                                this._copyFrom(e.item, true);
                            }
                        }
                        this.selectedSheetIndex = e.index;
                    }
                    else if (e.action === wijmo.collections.NotifyCollectionChangedAction.Reset) {
                        for (var i = 0; i < this.sheets.length; i++) {
                            item = this.sheets[i];
                            item._attachOwner(this);
                        }
                        if (this.sheets.length > 0) {
                            if (this.selectedSheetIndex === 0) {
                                this._copyFrom(this.selectedSheet, true);
                            }
                            this.selectedSheetIndex = 0;
                        }
                        else {
                            this.rows.clear();
                            this.columns.clear();
                            this._selectedSheetIndex = -1;
                        }
                    }
                    else {
                        if (this.sheets.length > 0) {
                            if (this.selectedSheetIndex >= this.sheets.length) {
                                this.selectedSheetIndex = 0;
                            }
                            else if (this.selectedSheetIndex > e.index) {
                                this._selectedSheetIndex -= 1;
                            }
                        }
                        else {
                            this.rows.clear();
                            this.columns.clear();
                            this._selectedSheetIndex = -1;
                        }
                    }
                    this.invalidate(true);
                };
                // Sheet visible changed event handler.
                FlexSheet.prototype._sheetVisibleChange = function (sender, e) {
                    if (!e.item.visible) {
                        if (e.index === this.selectedSheetIndex) {
                            if (this.selectedSheetIndex === this.sheets.length - 1) {
                                this.selectedSheetIndex = e.index - 1;
                            }
                            else {
                                this.selectedSheetIndex = e.index + 1;
                            }
                        }
                    }
                };
                // apply the styles for the selected cells.
                FlexSheet.prototype._applyStyleForCell = function (rowIndex, colIndex, cellStyle) {
                    var self = this, row = self.rows[rowIndex], currentCellStyle, mergeRange, cellIndex;
                    // Will ignore the cells in the HeaderRow. 
                    if (row instanceof HeaderRow || !row.isVisible) {
                        return;
                    }
                    cellIndex = rowIndex * self.columns.length + colIndex;
                    // Handle the merged range style.
                    mergeRange = self.selectedSheet._mergedRanges[cellIndex];
                    if (mergeRange) {
                        cellIndex = mergeRange.topRow * self.columns.length + mergeRange.leftCol;
                    }
                    currentCellStyle = self.selectedSheet._styledCells[cellIndex];
                    // Add new cell style for the cell.
                    if (!currentCellStyle) {
                        self.selectedSheet._styledCells[cellIndex] = {
                            className: cellStyle.className,
                            textAlign: cellStyle.textAlign,
                            verticalAlign: cellStyle.verticalAlign,
                            fontStyle: cellStyle.fontStyle,
                            fontWeight: cellStyle.fontWeight,
                            fontFamily: cellStyle.fontFamily,
                            fontSize: cellStyle.fontSize,
                            textDecoration: cellStyle.textDecoration,
                            backgroundColor: cellStyle.backgroundColor,
                            color: cellStyle.color,
                            format: cellStyle.format
                        };
                    }
                    else {
                        // Update the cell style.
                        currentCellStyle.className = cellStyle.className === 'normal' ? '' : cellStyle.className || currentCellStyle.className;
                        currentCellStyle.textAlign = cellStyle.textAlign || currentCellStyle.textAlign;
                        currentCellStyle.verticalAlign = cellStyle.verticalAlign || currentCellStyle.verticalAlign;
                        currentCellStyle.fontFamily = cellStyle.fontFamily || currentCellStyle.fontFamily;
                        currentCellStyle.fontSize = cellStyle.fontSize || currentCellStyle.fontSize;
                        currentCellStyle.backgroundColor = cellStyle.backgroundColor || currentCellStyle.backgroundColor;
                        currentCellStyle.color = cellStyle.color || currentCellStyle.color;
                        currentCellStyle.fontStyle = cellStyle.fontStyle === 'none' ? '' : cellStyle.fontStyle || currentCellStyle.fontStyle;
                        currentCellStyle.fontWeight = cellStyle.fontWeight === 'none' ? '' : cellStyle.fontWeight || currentCellStyle.fontWeight;
                        currentCellStyle.textDecoration = cellStyle.textDecoration === 'none' ? '' : cellStyle.textDecoration || currentCellStyle.textDecoration;
                        currentCellStyle.format = cellStyle.format || currentCellStyle.format;
                    }
                };
                // Check the format states for the cells of the selection.
                FlexSheet.prototype._checkCellFormat = function (rowIndex, colIndex, formatState) {
                    //return;
                    var cellIndex = rowIndex * this.columns.length + colIndex, mergeRange, cellStyle;
                    if (!this.selectedSheet) {
                        return;
                    }
                    mergeRange = this.selectedSheet._mergedRanges[cellIndex];
                    if (mergeRange) {
                        formatState.isMergedCell = true;
                        cellIndex = mergeRange.topRow * this.columns.length + mergeRange.leftCol;
                    }
                    cellStyle = this.selectedSheet._styledCells[cellIndex];
                    // get the format states for the cells of the selection.
                    if (cellStyle) {
                        formatState.isBold = formatState.isBold || cellStyle.fontWeight === 'bold';
                        formatState.isItalic = formatState.isItalic || cellStyle.fontStyle === 'italic';
                        formatState.isUnderline = formatState.isUnderline || cellStyle.textDecoration === 'underline';
                    }
                    // get text align state for the selected cells.
                    if (rowIndex === this.selection.row && colIndex === this.selection.col) {
                        if (cellStyle && cellStyle.textAlign) {
                            formatState.textAlign = cellStyle.textAlign;
                        }
                        else if (colIndex > -1) {
                            formatState.textAlign = this.columns[colIndex].getAlignment() || formatState.textAlign;
                        }
                    }
                };
                // Reset the merged range.
                FlexSheet.prototype._resetMergedRange = function (range) {
                    var rowIndex, colIndex, cellIndex, mergeRowIndex, mergeColIndex, mergeCellIndex, mergedCell, mergedCellExists = false;
                    for (rowIndex = range.topRow; rowIndex <= range.bottomRow; rowIndex++) {
                        for (colIndex = range.leftCol; colIndex <= range.rightCol; colIndex++) {
                            cellIndex = rowIndex * this.columns.length + colIndex;
                            mergedCell = this.selectedSheet._mergedRanges[cellIndex];
                            // Reset the merged state of each cell inside current merged range.
                            if (mergedCell) {
                                mergedCellExists = true;
                                for (mergeRowIndex = mergedCell.topRow; mergeRowIndex <= mergedCell.bottomRow; mergeRowIndex++) {
                                    for (mergeColIndex = mergedCell.leftCol; mergeColIndex <= mergedCell.rightCol; mergeColIndex++) {
                                        mergeCellIndex = mergeRowIndex * this.columns.length + mergeColIndex;
                                        {
                                            delete this.selectedSheet._mergedRanges[mergeCellIndex];
                                        }
                                    }
                                }
                            }
                        }
                    }
                    return mergedCellExists;
                };
                // update the styledCells hash and mergedRange hash for add\delete rows.
                FlexSheet.prototype._updateCellsForUpdatingRow = function (originalRowCount, index, count, isDelete) {
                    var _this = this;
                    //return;
                    var startIndex, cellIndex, newCellIndex, cellStyle, mergeRange, updatedMergeCell = {}, originalCellCount = originalRowCount * this.columns.length;
                    // update for deleting rows.
                    if (isDelete) {
                        startIndex = index * this.columns.length;
                        for (cellIndex = startIndex; cellIndex < originalCellCount; cellIndex++) {
                            newCellIndex = cellIndex - count * this.columns.length;
                            // Update the styledCells hash
                            cellStyle = this.selectedSheet._styledCells[cellIndex];
                            if (cellStyle) {
                                // if the cell is behind the delete cell range, we should update the cell index for the cell to store the style.
                                // if the cell is inside the delete cell range, it need be deleted directly.
                                if (cellIndex >= (index + count) * this.columns.length) {
                                    this.selectedSheet._styledCells[newCellIndex] = cellStyle;
                                }
                                delete this.selectedSheet._styledCells[cellIndex];
                            }
                            // Update the mergedRange hash
                            mergeRange = this.selectedSheet._mergedRanges[cellIndex];
                            if (mergeRange) {
                                if (index <= mergeRange.topRow && index + count > mergeRange.bottomRow) {
                                    // if the delete rows contain the merge cell range
                                    // we will delete the merge cell range directly.
                                    delete this.selectedSheet._mergedRanges[cellIndex];
                                }
                                else if (mergeRange.bottomRow < index || mergeRange.topRow >= index + count) {
                                    // Update the merge range when the deleted row is outside current merge cell range.
                                    if (mergeRange.topRow > index) {
                                        mergeRange.row -= count;
                                    }
                                    mergeRange.row2 -= count;
                                    this.selectedSheet._mergedRanges[newCellIndex] = mergeRange;
                                    delete this.selectedSheet._mergedRanges[cellIndex];
                                }
                                else {
                                    // Update the merge range when the deleted rows intersect with current merge cell range.
                                    this._updateCellMergeRangeForRow(mergeRange, index, count, updatedMergeCell, true);
                                }
                            }
                        }
                    }
                    else {
                        // Update for adding rows.
                        startIndex = index * this.columns.length - 1;
                        for (cellIndex = originalCellCount - 1; cellIndex > startIndex; cellIndex--) {
                            newCellIndex = cellIndex + this.columns.length * count;
                            // Update the styledCells hash
                            cellStyle = this.selectedSheet._styledCells[cellIndex];
                            if (cellStyle) {
                                this.selectedSheet._styledCells[newCellIndex] = cellStyle;
                                delete this.selectedSheet._styledCells[cellIndex];
                            }
                            // Update the mergedRange hash
                            mergeRange = this.selectedSheet._mergedRanges[cellIndex];
                            if (mergeRange) {
                                if (mergeRange.topRow < index && mergeRange.bottomRow >= index) {
                                    // Update the merge range when the added row is inside current merge cell range.
                                    this._updateCellMergeRangeForRow(mergeRange, index, count, updatedMergeCell);
                                }
                                else {
                                    // Update the merge range when the added row is outside current merge cell range.
                                    mergeRange.row += count;
                                    mergeRange.row2 += count;
                                    this.selectedSheet._mergedRanges[newCellIndex] = mergeRange;
                                    delete this.selectedSheet._mergedRanges[cellIndex];
                                }
                            }
                        }
                    }
                    Object.keys(updatedMergeCell).forEach(function (key) {
                        _this.selectedSheet._mergedRanges[key] = updatedMergeCell[key];
                    });
                };
                // Update the merge cell range when the add\delete rows intersect with current merge cell range.
                FlexSheet.prototype._updateCellMergeRangeForRow = function (currentRange, index, count, updatedMergeCell, isDelete) {
                    //return;
                    var rowIndex, columnIndex, cellIndex, newCellIndex, i, mergeRange, cloneRange;
                    if (isDelete) {
                        // Update the merge cell range for deleting rows.
                        for (rowIndex = currentRange.topRow; rowIndex <= currentRange.bottomRow; rowIndex++) {
                            for (columnIndex = currentRange.leftCol; columnIndex <= currentRange.rightCol; columnIndex++) {
                                cellIndex = rowIndex * this.columns.length + columnIndex;
                                newCellIndex = cellIndex - count * this.columns.length;
                                mergeRange = this.selectedSheet._mergedRanges[cellIndex];
                                if (mergeRange) {
                                    cloneRange = mergeRange.clone();
                                    // when the first delete row is above the merge cell range
                                    // we should adjust the topRow of the merge cell rang via the first delete row.
                                    if (cloneRange.row > index) {
                                        cloneRange.row -= cloneRange.row - index;
                                    }
                                    // when the last delete row is behind the merge cell range.
                                    // we should adjust the bottomRow of the merge cell rang via the first delete row.
                                    if (cloneRange.row2 < index + count - 1) {
                                        cloneRange.row2 -= cloneRange.row2 - index + 1;
                                    }
                                    else {
                                        cloneRange.row2 -= count;
                                    }
                                    if (rowIndex < index) {
                                        updatedMergeCell[cellIndex] = cloneRange;
                                    }
                                    else {
                                        if (rowIndex >= index + count) {
                                            updatedMergeCell[newCellIndex] = cloneRange;
                                        }
                                        delete this.selectedSheet._mergedRanges[cellIndex];
                                    }
                                }
                            }
                        }
                    }
                    else {
                        // Update the merge cell range for adding row.
                        for (rowIndex = currentRange.bottomRow; rowIndex >= currentRange.topRow; rowIndex--) {
                            for (columnIndex = currentRange.rightCol; columnIndex >= currentRange.leftCol; columnIndex--) {
                                cellIndex = rowIndex * this.columns.length + columnIndex;
                                mergeRange = this.selectedSheet._mergedRanges[cellIndex];
                                if (mergeRange) {
                                    cloneRange = mergeRange.clone();
                                    cloneRange.row2 += count;
                                    if (rowIndex < index) {
                                        updatedMergeCell[cellIndex] = cloneRange.clone();
                                    }
                                    for (i = 1; i <= count; i++) {
                                        newCellIndex = cellIndex + this.columns.length * i;
                                        updatedMergeCell[newCellIndex] = cloneRange;
                                    }
                                    delete this.selectedSheet._mergedRanges[cellIndex];
                                }
                            }
                        }
                    }
                };
                // update styledCells hash and mergedRange hash for add\delete columns.
                FlexSheet.prototype._updateCellsForUpdatingColumn = function (originalColumnCount, index, count, isDelete) {
                    var _this = this;
                    var cellIndex, newCellIndex, cellStyle, rowIndex, columnIndex, mergeRange, updatedMergeCell = {}, originalCellCount = this.rows.length * originalColumnCount;
                    // Update for deleting columns.
                    if (isDelete) {
                        for (cellIndex = index; cellIndex < originalCellCount; cellIndex++) {
                            rowIndex = Math.floor(cellIndex / originalColumnCount);
                            columnIndex = cellIndex % originalColumnCount;
                            newCellIndex = cellIndex - (count * (rowIndex + (columnIndex >= index ? 1 : 0)));
                            // Update the styledCells hash
                            cellStyle = this.selectedSheet._styledCells[cellIndex];
                            if (cellStyle) {
                                // if the cell is outside the delete cell range, we should update the cell index for the cell to store the style.
                                // otherwise it need be deleted directly.
                                if (columnIndex < index || columnIndex >= index + count) {
                                    this.selectedSheet._styledCells[newCellIndex] = cellStyle;
                                }
                                delete this.selectedSheet._styledCells[cellIndex];
                            }
                            // Update the mergedRange hash
                            mergeRange = this.selectedSheet._mergedRanges[cellIndex];
                            if (mergeRange) {
                                if (index <= mergeRange.leftCol && index + count > mergeRange.rightCol) {
                                    // if the delete columns contain the merge cell range
                                    // we will delete the merge cell range directly.
                                    delete this.selectedSheet._mergedRanges[cellIndex];
                                }
                                else if (mergeRange.rightCol < index || mergeRange.leftCol >= index + count) {
                                    // Update the merge range when the deleted column is outside current merge cell range.
                                    if (mergeRange.leftCol >= index) {
                                        mergeRange.col -= count;
                                        mergeRange.col2 -= count;
                                    }
                                    this.selectedSheet._mergedRanges[newCellIndex] = mergeRange;
                                    delete this.selectedSheet._mergedRanges[cellIndex];
                                }
                                else {
                                    // Update the merge range when the deleted columns intersect with current merge cell range.
                                    this._updateCellMergeRangeForColumn(mergeRange, index, count, originalColumnCount, updatedMergeCell, true);
                                }
                            }
                        }
                    }
                    else {
                        // Update for adding columns.
                        for (cellIndex = originalCellCount - 1; cellIndex >= index; cellIndex--) {
                            rowIndex = Math.floor(cellIndex / originalColumnCount);
                            columnIndex = cellIndex % originalColumnCount;
                            newCellIndex = cellIndex + rowIndex * count + (columnIndex >= index ? 1 : 0);
                            // Update the styledCells hash
                            cellStyle = this.selectedSheet._styledCells[cellIndex];
                            if (cellStyle) {
                                this.selectedSheet._styledCells[newCellIndex] = cellStyle;
                                delete this.selectedSheet._styledCells[cellIndex];
                            }
                            // Update the mergedRange hash
                            mergeRange = this.selectedSheet._mergedRanges[cellIndex];
                            if (mergeRange) {
                                if (mergeRange.leftCol < index && mergeRange.rightCol >= index) {
                                    // Update the merge range when the added column is inside current merge cell range.
                                    this._updateCellMergeRangeForColumn(mergeRange, index, count, originalColumnCount, updatedMergeCell);
                                }
                                else {
                                    // Update the merge range when the added column is outside current merge cell range.
                                    if (mergeRange.leftCol >= index) {
                                        mergeRange.col += count;
                                        mergeRange.col2 += count;
                                    }
                                    this.selectedSheet._mergedRanges[newCellIndex] = mergeRange;
                                    delete this.selectedSheet._mergedRanges[cellIndex];
                                }
                            }
                        }
                    }
                    Object.keys(updatedMergeCell).forEach(function (key) {
                        _this.selectedSheet._mergedRanges[key] = updatedMergeCell[key];
                    });
                };
                // Update the merge cell range when the add\delete columns intersect with current merge cell range.
                FlexSheet.prototype._updateCellMergeRangeForColumn = function (currentRange, index, count, originalColumnCount, updatedMergeCell, isDelete) {
                    var rowIndex, columnIndex, cellIndex, newCellIndex, i, mergeRange, cloneRange;
                    if (isDelete) {
                        // Update the merge cell range for deleting columns.
                        for (rowIndex = currentRange.topRow; rowIndex <= currentRange.bottomRow; rowIndex++) {
                            for (columnIndex = currentRange.leftCol; columnIndex <= currentRange.rightCol; columnIndex++) {
                                cellIndex = rowIndex * originalColumnCount + columnIndex;
                                newCellIndex = cellIndex - (count * (rowIndex + (columnIndex >= index ? 1 : 0)));
                                mergeRange = this.selectedSheet._mergedRanges[cellIndex];
                                if (mergeRange) {
                                    cloneRange = mergeRange.clone();
                                    // when the first delete column is before with merge cell range
                                    // we should adjust the leftCol of the merge cell rang via the first delete column.
                                    if (cloneRange.col > index) {
                                        cloneRange.col -= cloneRange.col - index;
                                    }
                                    // when the last delete row is behind the merge cell range.
                                    // we should adjust the bottomRow of the merge cell rang via the first delete row.
                                    if (cloneRange.col2 < index + count - 1) {
                                        cloneRange.col2 -= cloneRange.col2 - index + 1;
                                    }
                                    else {
                                        cloneRange.col2 -= count;
                                    }
                                    if (columnIndex < index || columnIndex >= index + count) {
                                        updatedMergeCell[newCellIndex] = cloneRange;
                                    }
                                    delete this.selectedSheet._mergedRanges[cellIndex];
                                }
                            }
                        }
                    }
                    else {
                        // Update the merge cell range for adding column.
                        for (rowIndex = currentRange.bottomRow; rowIndex >= currentRange.topRow; rowIndex--) {
                            for (columnIndex = currentRange.rightCol; columnIndex >= currentRange.leftCol; columnIndex--) {
                                cellIndex = rowIndex * originalColumnCount + columnIndex;
                                newCellIndex = cellIndex + rowIndex * count + (columnIndex >= index ? 1 : 0);
                                mergeRange = this.selectedSheet._mergedRanges[cellIndex];
                                if (mergeRange) {
                                    cloneRange = mergeRange.clone();
                                    cloneRange.col2 += count;
                                    if (columnIndex === index) {
                                        updatedMergeCell[newCellIndex - 1] = cloneRange.clone();
                                    }
                                    if (columnIndex >= index) {
                                        for (i = 0; i < count; i++) {
                                            updatedMergeCell[newCellIndex + i] = cloneRange;
                                        }
                                    }
                                    else {
                                        updatedMergeCell[newCellIndex] = cloneRange;
                                    }
                                    delete this.selectedSheet._mergedRanges[cellIndex];
                                }
                            }
                        }
                    }
                };
                // Clone the mergedRange of the Flexsheet
                FlexSheet.prototype._cloneMergedCells = function () {
                    var copy, mergedRanges;
                    if (!this.selectedSheet) {
                        return null;
                    }
                    mergedRanges = this.selectedSheet._mergedRanges;
                    // Handle the 3 simple types, and null or undefined
                    if (null == mergedRanges || "object" !== typeof mergedRanges)
                        return mergedRanges;
                    // Handle Object
                    if (mergedRanges instanceof Object) {
                        copy = {};
                        for (var attr in mergedRanges) {
                            if (mergedRanges.hasOwnProperty(attr)) {
                                if (mergedRanges[attr] && mergedRanges[attr].clone) {
                                    copy[attr] = mergedRanges[attr].clone();
                                }
                            }
                        }
                        return copy;
                    }
                    throw new Error("Unable to copy obj! Its type isn't supported.");
                };
                // Evaluate specified formula for flexsheet.
                FlexSheet.prototype._evaluate = function (formula, format, sheet, rowIndex, columnIndex) {
                    if (formula && formula.length > 1) {
                        formula = formula[0] === '=' ? formula : '=' + formula;
                        return this._calcEngine.evaluate(formula, format, sheet, rowIndex, columnIndex);
                    }
                    return formula;
                };
                // Copy the current flex sheet to the flexgrid of current sheet.
                FlexSheet.prototype._copyTo = function (sheet) {
                    var originAutoGenerateColumns = sheet.grid.autoGenerateColumns, colIndex, rowIndex, i;
                    sheet.grid.selection = new grid_1.CellRange();
                    sheet.grid.rows.clear();
                    sheet.grid.columns.clear();
                    sheet.grid.columnHeaders.columns.clear();
                    sheet.grid.rowHeaders.rows.clear();
                    if (this.itemsSource) {
                        sheet.grid.autoGenerateColumns = false;
                        sheet.itemsSource = this.itemsSource;
                        sheet.grid.collectionView.beginUpdate();
                        if (!(sheet.grid.itemsSource instanceof wijmo.collections.CollectionView)) {
                            sheet.grid.collectionView.sortDescriptions.clear();
                            for (i = 0; i < this.collectionView.sortDescriptions.length; i++) {
                                sheet.grid.collectionView.sortDescriptions.push(this.collectionView.sortDescriptions[i]);
                            }
                        }
                    }
                    else {
                        sheet.itemsSource = null;
                        for (rowIndex = 0; rowIndex < this.rows.length; rowIndex++) {
                            sheet.grid.rows.push(this.rows[rowIndex]);
                        }
                    }
                    sheet._filterDefinition = this._filter.filterDefinition;
                    for (colIndex = 0; colIndex < this.columns.length; colIndex++) {
                        sheet.grid.columns.push(this.columns[colIndex]);
                    }
                    if (sheet.grid.collectionView) {
                        this._resetMappedColumns(sheet.grid);
                        sheet.grid.collectionView.endUpdate();
                    }
                    sheet.grid.autoGenerateColumns = originAutoGenerateColumns;
                    sheet.grid.frozenRows = this.frozenRows;
                    sheet.grid.frozenColumns = this.frozenColumns;
                    sheet.grid.selection = this.selection;
                    sheet._scrollPosition = this.scrollPosition;
                    this.columns._dirty = true;
                    this.rows._dirty = true;
                };
                // Copy the flexgrid of current sheet to flexsheet.
                FlexSheet.prototype._copyFrom = function (sheet, needRefresh) {
                    if (needRefresh === void 0) { needRefresh = true; }
                    var self = this, originAutoGenerateColumns = self.autoGenerateColumns, colIndex, rowIndex, i, row;
                    self._isCopyingOrUndoing = true;
                    self._dragable = false;
                    self.rows.clear();
                    self.columns.clear();
                    self.columnHeaders.columns.clear();
                    self.rowHeaders.rows.clear();
                    self.selection = new grid_1.CellRange();
                    if (sheet.selectionRanges.length > 1 && self.selectionMode === grid_1.SelectionMode.CellRange) {
                        self._enableMulSel = true;
                    }
                    if (sheet.itemsSource) {
                        self.autoGenerateColumns = false;
                        self.itemsSource = sheet.itemsSource;
                        self.collectionView.beginUpdate();
                        if (!(self.itemsSource instanceof wijmo.collections.CollectionView)) {
                            self.collectionView.sortDescriptions.clear();
                            for (i = 0; i < sheet.grid.collectionView.sortDescriptions.length; i++) {
                                self.collectionView.sortDescriptions.push(sheet.grid.collectionView.sortDescriptions[i]);
                            }
                        }
                    }
                    else {
                        self.itemsSource = null;
                        for (rowIndex = 0; rowIndex < sheet.grid.rows.length; rowIndex++) {
                            self.rows.push(sheet.grid.rows[rowIndex]);
                        }
                    }
                    for (colIndex = 0; colIndex < sheet.grid.columns.length; colIndex++) {
                        self.columns.push(sheet.grid.columns[colIndex]);
                    }
                    if (self.collectionView) {
                        self._resetMappedColumns(self);
                        self.collectionView.endUpdate();
                        self.collectionView.collectionChanged.addHandler(function (sender, e) {
                            if (e.action === wijmo.collections.NotifyCollectionChangedAction.Reset) {
                                self.invalidate();
                            }
                        }, self);
                    }
                    if (self.rows.length && self.columns.length) {
                        self.selection = sheet.grid.selection;
                    }
                    if (sheet._filterDefinition) {
                        self._filter.filterDefinition = sheet._filterDefinition;
                    }
                    self.autoGenerateColumns = originAutoGenerateColumns;
                    // Hide the invisible row/column after freezing. (TFS 152188)
                    if (sheet._freezeHiddenRowCnt > 0) {
                        for (rowIndex = 0; rowIndex < sheet._freezeHiddenRowCnt; rowIndex++) {
                            row = self.rows[rowIndex];
                            if (!(row instanceof HeaderRow)) {
                                row.visible = false;
                            }
                        }
                    }
                    if (sheet._freezeHiddenColumnCnt > 0) {
                        for (colIndex = 0; colIndex < sheet._freezeHiddenColumnCnt; colIndex++) {
                            self.columns[colIndex].visible = false;
                        }
                    }
                    self.frozenRows = sheet.grid.frozenRows;
                    self.frozenColumns = sheet.grid.frozenColumns;
                    self._isCopyingOrUndoing = false;
                    if (self._addingSheet) {
                        if (self._toRefresh) {
                            clearTimeout(self._toRefresh);
                            self._toRefresh = null;
                        }
                        self._toRefresh = setTimeout(function () {
                            self.rows._dirty = true;
                            self.columns._dirty = true;
                            self.invalidate();
                        }, 10);
                        self._addingSheet = false;
                    }
                    else if (needRefresh) {
                        self.refresh();
                    }
                    self.scrollPosition = sheet._scrollPosition;
                };
                // Reset the _mappedColumns hash for the flexgrid. 
                FlexSheet.prototype._resetMappedColumns = function (flex) {
                    var col, sds, i = 0;
                    flex._mappedColumns = null;
                    if (flex.collectionView) {
                        sds = flex.collectionView.sortDescriptions;
                        for (; i < sds.length; i++) {
                            col = flex.columns.getColumn(sds[i].property);
                            if (col && col.dataMap) {
                                if (!flex._mappedColumns) {
                                    flex._mappedColumns = {};
                                }
                                flex._mappedColumns[col.binding] = col.dataMap;
                            }
                        }
                    }
                };
                // reset the filter definition for the flexsheet.
                FlexSheet.prototype._resetFilterDefinition = function () {
                    this._filter.filterDefinition = JSON.stringify({
                        defaultFilterType: wijmo.grid.filter.FilterType.Both,
                        filters: []
                    });
                };
                // Load the workbook instance to the flexsheet
                FlexSheet.prototype._loadFromWorkbook = function (workbook) {
                    var sheetCount, sheetIndex = 0, self = this;
                    if (workbook.sheets == null || workbook.sheets.length === 0) {
                        return;
                    }
                    self.clear();
                    self._reservedContent = workbook.reservedContent;
                    sheetCount = workbook.sheets.length;
                    for (; sheetIndex < sheetCount; sheetIndex++) {
                        if (sheetIndex > 0) {
                            self.addUnboundSheet();
                        }
                        wijmo.grid.xlsx.FlexGridXlsxConverter.load(self.selectedSheet.grid, workbook, { sheetIndex: sheetIndex, includeColumnHeaders: false });
                        if (self.selectedSheet.grid['wj_sheetInfo']) {
                            self.selectedSheet.name = self.selectedSheet.grid['wj_sheetInfo'].name;
                            self.selectedSheet.visible = self.selectedSheet.grid['wj_sheetInfo'].visible;
                            self.selectedSheet._styledCells = self.selectedSheet.grid['wj_sheetInfo'].styledCells;
                            self.selectedSheet._mergedRanges = self.selectedSheet.grid['wj_sheetInfo'].mergedRanges;
                        }
                        self._copyFrom(self.selectedSheet, false);
                    }
                    if (workbook.activeWorksheet != null && workbook.activeWorksheet > -1 && workbook.activeWorksheet < self.sheets.length) {
                        self.selectedSheetIndex = workbook.activeWorksheet;
                    }
                    else {
                        self.selectedSheetIndex = 0;
                    }
                    self.onLoaded();
                };
                // Save the flexsheet to the workbook instance.
                FlexSheet.prototype._saveToWorkbook = function () {
                    var mainBook, tmpBook, currentSheet, sheetIndex;
                    if (this.sheets.length === 0) {
                        throw 'The flexsheet is empty.';
                    }
                    currentSheet = this.sheets[0];
                    mainBook = wijmo.grid.xlsx.FlexGridXlsxConverter.save(currentSheet.grid, { sheetName: currentSheet.name, sheetVisible: currentSheet.visible, includeColumnHeaders: false });
                    mainBook.reservedContent = this._reservedContent;
                    for (sheetIndex = 1; sheetIndex < this.sheets.length; sheetIndex++) {
                        currentSheet = this.sheets[sheetIndex];
                        tmpBook = wijmo.grid.xlsx.FlexGridXlsxConverter.save(currentSheet.grid, { sheetName: currentSheet.name, sheetVisible: currentSheet.visible, includeColumnHeaders: false });
                        mainBook._addWorkSheet(tmpBook.sheets[0], sheetIndex);
                    }
                    mainBook.activeWorksheet = this.selectedSheetIndex;
                    return mainBook;
                };
                // mouseDown event handler.
                // This event handler for handling selecting columns
                FlexSheet.prototype._mouseDown = function (e) {
                    var userAgent = window.navigator.userAgent, ht = this.hitTest(e), cols = this.columns, currentRange, colIndex, selected, newSelection, edt;
                    this._wholeColumnsSelected = false;
                    if (this._dragable) {
                        this._isDragging = true;
                        this._draggingMarker = document.createElement('div');
                        wijmo.setCss(this._draggingMarker, {
                            position: 'absolute',
                            display: 'none',
                            borderStyle: 'dotted',
                            cursor: 'move'
                        });
                        document.body.appendChild(this._draggingMarker);
                        this._draggingTooltip = new wijmo.Tooltip();
                        this._draggingCells = this.selection;
                        if (this.selectedSheet) {
                            this.selectedSheet.selectionRanges.clear();
                        }
                        this.onDraggingRowColumn(new DraggingRowColumnEventArgs(this._draggingRow, e.shiftKey));
                        e.preventDefault();
                        return;
                    }
                    // Set the _htDown of the _EditHandler, when the slection of the FlexSheet contains the range of current hitDown (TFS #139847)
                    if (ht.cellType !== grid_1.CellType.None) {
                        edt = wijmo.tryCast(e.target, HTMLInputElement);
                        if (edt == null && this._checkHitWithinSelection(ht)) {
                            this._edtHdl._htDown = ht;
                        }
                        this._isClicking = true;
                    }
                    if (this.selectionMode === grid_1.SelectionMode.CellRange) {
                        if (e.ctrlKey) {
                            if (!this._enableMulSel) {
                                this._enableMulSel = true;
                            }
                        }
                        else {
                            if (ht.cellType !== grid_1.CellType.None) {
                                if (this.selectedSheet) {
                                    this.selectedSheet.selectionRanges.clear();
                                }
                                if (this._enableMulSel) {
                                    this.refresh(false);
                                }
                                this._enableMulSel = false;
                            }
                        }
                    }
                    else {
                        this._enableMulSel = false;
                        if (this.selectedSheet) {
                            this.selectedSheet.selectionRanges.clear();
                        }
                    }
                    this._htDown = ht;
                    // If there is no rows or columns in the flexsheet, we don't need deal with anything in the mouse down event(TFS 122628)
                    if (this.rows.length === 0 || this.columns.length === 0) {
                        return;
                    }
                    if (!userAgent.match(/iPad/i) && !userAgent.match(/iPhone/i)) {
                        this._contextMenu.hide();
                    }
                    if (this.selectionMode !== grid_1.SelectionMode.CellRange) {
                        return;
                    }
                    // When right click the row header, we should select current row. (TFS 121167)
                    if (ht.cellType === grid_1.CellType.RowHeader && e.which === 3) {
                        newSelection = new grid_1.CellRange(ht.row, 0, ht.row, this.columns.length - 1);
                        if (!this.selection.contains(newSelection)) {
                            this.selection = newSelection;
                        }
                        return;
                    }
                    if (ht.cellType !== grid_1.CellType.ColumnHeader && ht.cellType !== grid_1.CellType.None) {
                        return;
                    }
                    if (ht.col > -1 && this.columns[ht.col].isSelected) {
                        return;
                    }
                    if (!wijmo.hasClass(e.target, 'wj-cell') || ht.edgeRight) {
                        return;
                    }
                    this._columnHeaderClicked = true;
                    this._wholeColumnsSelected = true;
                    if (e.shiftKey) {
                        this._multiSelectColumns(ht);
                    }
                    else {
                        currentRange = new grid_1.CellRange(this.itemsSource ? 1 : 0, ht.col, this.rows.length - 1, ht.col);
                        if (e.which === 3 && this.selection.contains(currentRange)) {
                            return;
                        }
                        this.select(currentRange);
                    }
                };
                // mouseMove event handler
                // This event handler for handling multiple selecting columns.
                FlexSheet.prototype._mouseMove = function (e) {
                    var ht = this.hitTest(e), selection = this.selection, rowCnt = this.rows.length, colCnt = this.columns.length, cursor = this.hostElement.style.cursor, isTopRow;
                    if (this._isDragging) {
                        this.hostElement.style.cursor = 'move';
                        this._showDraggingMarker(e);
                        return;
                    }
                    if (this.itemsSource) {
                        isTopRow = selection.topRow === 0 || selection.topRow === 1;
                    }
                    else {
                        isTopRow = selection.topRow === 0;
                    }
                    if (selection && ht.cellType !== grid_1.CellType.None && !this.itemsSource) {
                        this._draggingColumn = isTopRow && selection.bottomRow === rowCnt - 1;
                        this._draggingRow = selection.leftCol === 0 && selection.rightCol === colCnt - 1;
                        if (ht.cellType === grid_1.CellType.Cell) {
                            if (this._draggingColumn && (((ht.col === selection.leftCol - 1 || ht.col === selection.rightCol) && ht.edgeRight)
                                || (ht.row === rowCnt - 1 && ht.edgeBottom))) {
                                cursor = 'move';
                            }
                            if (this._draggingRow && !this._containsGroupRows(selection) && ((ht.row === selection.topRow - 1 || ht.row === selection.bottomRow) && ht.edgeBottom
                                || (ht.col === colCnt - 1 && ht.edgeRight))) {
                                cursor = 'move';
                            }
                        }
                        else if (ht.cellType === grid_1.CellType.ColumnHeader) {
                            if (ht.edgeBottom) {
                                if (this._draggingColumn && (ht.col >= selection.leftCol && ht.col <= selection.rightCol)) {
                                    cursor = 'move';
                                }
                                else if (this._draggingRow && selection.topRow === 0) {
                                    cursor = 'move';
                                }
                            }
                        }
                        else if (ht.cellType === grid_1.CellType.RowHeader) {
                            if (ht.edgeRight) {
                                if (this._draggingColumn && selection.leftCol === 0) {
                                    cursor = 'move';
                                }
                                else if (this._draggingRow && (ht.row >= selection.topRow && ht.row <= selection.bottomRow) && !this._containsGroupRows(selection)) {
                                    cursor = 'move';
                                }
                            }
                        }
                        if (cursor === 'move') {
                            this._dragable = true;
                        }
                        else {
                            this._dragable = false;
                        }
                        this.hostElement.style.cursor = cursor;
                    }
                    if (!this._htDown || !this._htDown.panel) {
                        return;
                    }
                    ht = new grid_1.HitTestInfo(this._htDown.panel, e);
                    this._multiSelectColumns(ht);
                    this.scrollIntoView(ht.row, ht.col);
                };
                // mouseUp event handler.
                // This event handler for resetting the variable for handling multiple select columns
                FlexSheet.prototype._mouseUp = function (e) {
                    if (this._isDragging) {
                        if (!this._draggingCells.equals(this._dropRange)) {
                            this._handleDropping(e);
                            this.onDroppingRowColumn();
                        }
                        this._draggingCells = null;
                        this._dropRange = null;
                        document.body.removeChild(this._draggingMarker);
                        this._draggingMarker = null;
                        this._draggingTooltip.hide();
                        this._draggingTooltip = null;
                        this._isDragging = false;
                        this._draggingColumn = false;
                        this._draggingRow = false;
                    }
                    if (this._htDown && this._htDown.cellType !== grid_1.CellType.None && this.selection.isValid && this.selectedSheet) {
                        // Store current selection in the selection array for multiple selection.
                        if (this.selectionMode === grid_1.SelectionMode.ListBox || this.selectionMode === grid_1.SelectionMode.Row || this.selectionMode === grid_1.SelectionMode.RowRange) {
                            this.selectedSheet.selectionRanges.push(new grid_1.CellRange(this.selection.row, 0, this.selection.row2, this.columns.length - 1));
                        }
                        else if (this._htDown.cellType === grid_1.CellType.TopLeft) {
                            this.selectedSheet.selectionRanges.push(new grid_1.CellRange(this.selectedSheet.itemsSource ? 1 : 0, 0, this.rows.length - 1, this.columns.length - 1));
                        }
                        else {
                            this.selectedSheet.selectionRanges.push(this.selection);
                        }
                        this._enableMulSel = false;
                    }
                    this._isClicking = false;
                    this._columnHeaderClicked = false;
                    this._htDown = null;
                };
                // Click event handler.
                FlexSheet.prototype._click = function () {
                    var self = this, userAgent = window.navigator.userAgent;
                    // When click in the body, we also need hide the context menu.
                    if (!userAgent.match(/iPad/i) && !userAgent.match(/iPhone/i)) {
                        self._contextMenu.hide();
                    }
                    setTimeout(function () {
                        self.hideFunctionList();
                    }, 200);
                };
                // touch start event handler for iOS device
                FlexSheet.prototype._touchStart = function (e) {
                    var self = this;
                    if (!wijmo.hasClass(e.target, 'wj-context-menu-item')) {
                        self._contextMenu.hide();
                    }
                    self._longClickTimer = setTimeout(function () {
                        var ht;
                        ht = self.hitTest(e);
                        if (ht && ht.cellType !== grid_1.CellType.None && !self.itemsSource) {
                            self._contextMenu.show(undefined, new wijmo.Point(e.pageX + 10, e.pageY + 10));
                        }
                    }, 500);
                };
                // touch end event handler for iOS device
                FlexSheet.prototype._touchEnd = function () {
                    clearTimeout(this._longClickTimer);
                };
                // Show the dragging marker while the mouse moving.
                FlexSheet.prototype._showDraggingMarker = function (e) {
                    var hitInfo = new grid_1.HitTestInfo(this.cells, e), selection = this.selection, colCnt = this.columns.length, rowCnt = this.rows.length, scrollOffset = this._cumulativeScrollOffset(this.hostElement), rootBounds = this['_root'].getBoundingClientRect(), rootOffsetX = rootBounds.left + scrollOffset.x, rootOffsetY = rootBounds.top + scrollOffset.y, hitCellBounds, selectionCnt, hit, height, width, rootSize, i, content, css;
                    this.scrollIntoView(hitInfo.row, hitInfo.col);
                    if (this._draggingColumn) {
                        selectionCnt = selection.rightCol - selection.leftCol + 1;
                        hit = hitInfo.col;
                        width = 0;
                        if (hit < 0 || hit + selectionCnt > colCnt) {
                            hit = colCnt - selectionCnt;
                        }
                        hitCellBounds = this.cells.getCellBoundingRect(0, hit);
                        rootSize = this['_root'].offsetHeight - this['_eCHdr'].offsetHeight;
                        height = this.cells.height;
                        height = height > rootSize ? rootSize : height;
                        for (i = 0; i < selectionCnt; i++) {
                            width += this.columns[hit + i].renderSize;
                        }
                        content = FlexSheet.convertNumberToAlpha(hit) + ' : ' + FlexSheet.convertNumberToAlpha(hit + selectionCnt - 1);
                        if (this._dropRange) {
                            this._dropRange.col = hit;
                            this._dropRange.col2 = hit + selectionCnt - 1;
                        }
                        else {
                            this._dropRange = new grid_1.CellRange(0, hit, this.rows.length - 1, hit + selectionCnt - 1);
                        }
                    }
                    else if (this._draggingRow) {
                        selectionCnt = selection.bottomRow - selection.topRow + 1;
                        hit = hitInfo.row;
                        height = 0;
                        if (hit < 0 || hit + selectionCnt > rowCnt) {
                            hit = rowCnt - selectionCnt;
                        }
                        hitCellBounds = this.cells.getCellBoundingRect(hit, 0);
                        rootSize = this['_root'].offsetWidth - this['_eRHdr'].offsetWidth;
                        for (i = 0; i < selectionCnt; i++) {
                            height += this.rows[hit + i].renderSize;
                        }
                        width = this.cells.width;
                        width = width > rootSize ? rootSize : width;
                        content = (hit + 1) + ' : ' + (hit + selectionCnt);
                        if (this._dropRange) {
                            this._dropRange.row = hit;
                            this._dropRange.row2 = hit + selectionCnt - 1;
                        }
                        else {
                            this._dropRange = new grid_1.CellRange(hit, 0, hit + selectionCnt - 1, this.columns.length - 1);
                        }
                    }
                    if (!hitCellBounds) {
                        return;
                    }
                    css = {
                        display: 'inline',
                        zIndex: '9999',
                        opacity: 0.5,
                        top: hitCellBounds.top - (this._draggingColumn ? this._ptScrl.y : 0) + scrollOffset.y,
                        left: hitCellBounds.left - (this._draggingRow ? this._ptScrl.x : 0) + scrollOffset.x,
                        height: height,
                        width: width
                    };
                    hitCellBounds.top = hitCellBounds.top - (this._draggingColumn ? this._ptScrl.y : 0);
                    hitCellBounds.left = hitCellBounds.left - (this._draggingRow ? this._ptScrl.x : 0);
                    if (this._rtl && this._draggingRow) {
                        css.left = css.left - width + hitCellBounds.width + 2 * this._ptScrl.x;
                        hitCellBounds.left = hitCellBounds.left + 2 * this._ptScrl.x;
                    }
                    if (this._draggingRow) {
                        if (rootOffsetX + this['_eRHdr'].offsetWidth !== css.left || rootOffsetY + this['_root'].offsetHeight < css.top + css.height) {
                            return;
                        }
                    }
                    else {
                        if (rootOffsetY + this['_eCHdr'].offsetHeight !== css.top || rootOffsetX + this['_root'].offsetWidth < css.left + css.width) {
                            return;
                        }
                    }
                    wijmo.setCss(this._draggingMarker, css);
                    this._draggingTooltip.show(this.hostElement, content, hitCellBounds);
                };
                // Handle dropping rows or columns.
                FlexSheet.prototype._handleDropping = function (e) {
                    var self = this, srcRowIndex, srcColIndex, desRowIndex, desColIndex, moveCellsAction;
                    if (!self.selectedSheet || !self._draggingCells || !self._dropRange || self._containsMergedCells(self._draggingCells) || self._containsMergedCells(self._dropRange)) {
                        return;
                    }
                    self._clearCalcEngine();
                    if ((self._draggingColumn && self._draggingCells.leftCol > self._dropRange.leftCol)
                        || (self._draggingRow && self._draggingCells.topRow > self._dropRange.topRow)) {
                        // Handle changing the columns or rows position.
                        if (e.shiftKey) {
                            if (self._draggingColumn) {
                                desColIndex = self._dropRange.leftCol;
                                for (srcColIndex = self._draggingCells.leftCol; srcColIndex <= self._draggingCells.rightCol; srcColIndex++) {
                                    self.columns.moveElement(srcColIndex, desColIndex);
                                    desColIndex++;
                                }
                            }
                            else if (self._draggingRow) {
                                desRowIndex = self._dropRange.topRow;
                                for (srcRowIndex = self._draggingCells.topRow; srcRowIndex <= self._draggingCells.bottomRow; srcRowIndex++) {
                                    self.rows.moveElement(srcRowIndex, desRowIndex);
                                    desRowIndex++;
                                }
                            }
                            self._exchangeCellStyle(true);
                        }
                        else {
                            // Handle moving or copying the cell content.
                            moveCellsAction = new sheet_1._MoveCellsAction(self, self._draggingCells, self._dropRange, e.ctrlKey);
                            desRowIndex = self._dropRange.topRow;
                            for (srcRowIndex = self._draggingCells.topRow; srcRowIndex <= self._draggingCells.bottomRow; srcRowIndex++) {
                                desColIndex = self._dropRange.leftCol;
                                for (srcColIndex = self._draggingCells.leftCol; srcColIndex <= self._draggingCells.rightCol; srcColIndex++) {
                                    self._moveCellContent(srcRowIndex, srcColIndex, desRowIndex, desColIndex, e.ctrlKey);
                                    if (self._draggingColumn && desRowIndex === self._dropRange.topRow) {
                                        self.columns[desColIndex].dataType = self.columns[srcColIndex].dataType ? self.columns[srcColIndex].dataType : wijmo.DataType.Object;
                                        self.columns[desColIndex].align = self.columns[srcColIndex].align;
                                        self.columns[desColIndex].format = self.columns[srcColIndex].format;
                                        if (!e.ctrlKey) {
                                            self.columns[srcColIndex].dataType = wijmo.DataType.Object;
                                            self.columns[srcColIndex].align = null;
                                            self.columns[srcColIndex].format = null;
                                        }
                                    }
                                    desColIndex++;
                                }
                                desRowIndex++;
                            }
                            if (self._draggingColumn && !e.ctrlKey) {
                                desColIndex = self._dropRange.leftCol;
                                for (srcColIndex = self._draggingCells.leftCol; srcColIndex <= self._draggingCells.rightCol; srcColIndex++) {
                                    self._updateColumnFiler(srcColIndex, desColIndex);
                                    desColIndex++;
                                }
                            }
                            if (moveCellsAction.saveNewState()) {
                                self._undoStack._addAction(moveCellsAction);
                            }
                        }
                    }
                    else if ((self._draggingColumn && self._draggingCells.leftCol < self._dropRange.leftCol)
                        || (self._draggingRow && self._draggingCells.topRow < self._dropRange.topRow)) {
                        // Handle changing the columns or rows position.
                        if (e.shiftKey) {
                            if (self._draggingColumn) {
                                desColIndex = self._dropRange.rightCol;
                                for (srcColIndex = self._draggingCells.rightCol; srcColIndex >= self._draggingCells.leftCol; srcColIndex--) {
                                    self.columns.moveElement(srcColIndex, desColIndex);
                                    desColIndex--;
                                }
                            }
                            else if (self._draggingRow) {
                                desRowIndex = self._dropRange.bottomRow;
                                for (srcRowIndex = self._draggingCells.bottomRow; srcRowIndex >= self._draggingCells.topRow; srcRowIndex--) {
                                    self.rows.moveElement(srcRowIndex, desRowIndex);
                                    desRowIndex--;
                                }
                            }
                            self._exchangeCellStyle(false);
                        }
                        else {
                            // Handle moving or copying the cell content.
                            moveCellsAction = new sheet_1._MoveCellsAction(self, self._draggingCells, self._dropRange, e.ctrlKey);
                            desRowIndex = self._dropRange.bottomRow;
                            for (srcRowIndex = self._draggingCells.bottomRow; srcRowIndex >= self._draggingCells.topRow; srcRowIndex--) {
                                desColIndex = self._dropRange.rightCol;
                                for (srcColIndex = self._draggingCells.rightCol; srcColIndex >= self._draggingCells.leftCol; srcColIndex--) {
                                    self._moveCellContent(srcRowIndex, srcColIndex, desRowIndex, desColIndex, e.ctrlKey);
                                    if (self._draggingColumn && desRowIndex === self._dropRange.bottomRow) {
                                        self.columns[desColIndex].dataType = self.columns[srcColIndex].dataType ? self.columns[srcColIndex].dataType : wijmo.DataType.Object;
                                        self.columns[desColIndex].align = self.columns[srcColIndex].align;
                                        self.columns[desColIndex].format = self.columns[srcColIndex].format;
                                        if (!e.ctrlKey) {
                                            self.columns[srcColIndex].dataType = wijmo.DataType.Object;
                                            self.columns[srcColIndex].align = null;
                                            self.columns[srcColIndex].format = null;
                                        }
                                    }
                                    desColIndex--;
                                }
                                desRowIndex--;
                            }
                            if (self._draggingColumn && !e.ctrlKey) {
                                desColIndex = self._dropRange.rightCol;
                                for (srcColIndex = self._draggingCells.rightCol; srcColIndex >= self._draggingCells.leftCol; srcColIndex--) {
                                    self._updateColumnFiler(srcColIndex, desColIndex);
                                    desColIndex--;
                                }
                            }
                            if (moveCellsAction.saveNewState()) {
                                self._undoStack._addAction(moveCellsAction);
                            }
                        }
                    }
                    self.select(self._dropRange);
                    self.selectedSheet.selectionRanges.push(self.selection);
                    // Ensure that the host element of FlexSheet get focus after dropping. (TFS 142888)
                    self.hostElement.focus();
                };
                // Move the content and style of the source cell to the destination cell.
                FlexSheet.prototype._moveCellContent = function (srcRowIndex, srcColIndex, desRowIndex, desColIndex, isCopyContent) {
                    var val = this.getCellData(srcRowIndex, srcColIndex, false), srcCellIndex = srcRowIndex * this.columns.length + srcColIndex, desCellIndex = desRowIndex * this.columns.length + desColIndex, srcCellStyle = this.selectedSheet._styledCells[srcCellIndex];
                    this.setCellData(desRowIndex, desColIndex, val);
                    // Copy the cell style of the source cell to the destination cell.
                    if (srcCellStyle) {
                        this.selectedSheet._styledCells[desCellIndex] = JSON.parse(JSON.stringify(srcCellStyle));
                    }
                    else {
                        delete this.selectedSheet._styledCells[desCellIndex];
                    }
                    // If we just move the columns or the rows, we need remove the content and styles of the cells in the columns or the rows.
                    if (!isCopyContent) {
                        this.setCellData(srcRowIndex, srcColIndex, undefined);
                        delete this.selectedSheet._styledCells[srcCellIndex];
                    }
                };
                // Exchange the cell style for changing the rows or columns position.
                FlexSheet.prototype._exchangeCellStyle = function (isReverse) {
                    var rowIndex, colIndex, cellIndex, newCellIndex, draggingRange, index = 0, srcCellStyles = [];
                    // Store the style of the source cells and delete the style of the source cells.
                    // Since the stored style will be moved to the destination cells.
                    for (rowIndex = this._draggingCells.topRow; rowIndex <= this._draggingCells.bottomRow; rowIndex++) {
                        for (colIndex = this._draggingCells.leftCol; colIndex <= this._draggingCells.rightCol; colIndex++) {
                            cellIndex = rowIndex * this.columns.length + colIndex;
                            if (this.selectedSheet._styledCells[cellIndex]) {
                                srcCellStyles.push(JSON.parse(JSON.stringify(this.selectedSheet._styledCells[cellIndex])));
                                delete this.selectedSheet._styledCells[cellIndex];
                            }
                            else {
                                srcCellStyles.push(undefined);
                            }
                        }
                    }
                    // Adjust the style of the cells that is between the dragging cells and the drop range.
                    if (isReverse) {
                        if (this._draggingColumn) {
                            draggingRange = this._draggingCells.rightCol - this._draggingCells.leftCol + 1;
                            for (colIndex = this._draggingCells.leftCol - 1; colIndex >= this._dropRange.leftCol; colIndex--) {
                                for (rowIndex = 0; rowIndex < this.rows.length; rowIndex++) {
                                    cellIndex = rowIndex * this.columns.length + colIndex;
                                    newCellIndex = rowIndex * this.columns.length + colIndex + draggingRange;
                                    if (this.selectedSheet._styledCells[cellIndex]) {
                                        this.selectedSheet._styledCells[newCellIndex] = JSON.parse(JSON.stringify(this.selectedSheet._styledCells[cellIndex]));
                                        delete this.selectedSheet._styledCells[cellIndex];
                                    }
                                    else {
                                        delete this.selectedSheet._styledCells[newCellIndex];
                                    }
                                }
                            }
                        }
                        else if (this._draggingRow) {
                            draggingRange = this._draggingCells.bottomRow - this._draggingCells.topRow + 1;
                            for (rowIndex = this._draggingCells.topRow - 1; rowIndex >= this._dropRange.topRow; rowIndex--) {
                                for (colIndex = 0; colIndex < this.columns.length; colIndex++) {
                                    cellIndex = rowIndex * this.columns.length + colIndex;
                                    newCellIndex = (rowIndex + draggingRange) * this.columns.length + colIndex;
                                    if (this.selectedSheet._styledCells[cellIndex]) {
                                        this.selectedSheet._styledCells[newCellIndex] = JSON.parse(JSON.stringify(this.selectedSheet._styledCells[cellIndex]));
                                        delete this.selectedSheet._styledCells[cellIndex];
                                    }
                                    else {
                                        delete this.selectedSheet._styledCells[newCellIndex];
                                    }
                                }
                            }
                        }
                    }
                    else {
                        if (this._draggingColumn) {
                            draggingRange = this._draggingCells.rightCol - this._draggingCells.leftCol + 1;
                            for (colIndex = this._draggingCells.rightCol + 1; colIndex <= this._dropRange.rightCol; colIndex++) {
                                for (rowIndex = 0; rowIndex < this.rows.length; rowIndex++) {
                                    cellIndex = rowIndex * this.columns.length + colIndex;
                                    newCellIndex = rowIndex * this.columns.length + colIndex - draggingRange;
                                    if (this.selectedSheet._styledCells[cellIndex]) {
                                        this.selectedSheet._styledCells[newCellIndex] = JSON.parse(JSON.stringify(this.selectedSheet._styledCells[cellIndex]));
                                        delete this.selectedSheet._styledCells[cellIndex];
                                    }
                                    else {
                                        delete this.selectedSheet._styledCells[newCellIndex];
                                    }
                                }
                            }
                        }
                        else if (this._draggingRow) {
                            draggingRange = this._draggingCells.bottomRow - this._draggingCells.topRow + 1;
                            for (rowIndex = this._draggingCells.bottomRow + 1; rowIndex <= this._dropRange.bottomRow; rowIndex++) {
                                for (colIndex = 0; colIndex < this.columns.length; colIndex++) {
                                    cellIndex = rowIndex * this.columns.length + colIndex;
                                    newCellIndex = (rowIndex - draggingRange) * this.columns.length + colIndex;
                                    if (this.selectedSheet._styledCells[cellIndex]) {
                                        this.selectedSheet._styledCells[newCellIndex] = JSON.parse(JSON.stringify(this.selectedSheet._styledCells[cellIndex]));
                                        delete this.selectedSheet._styledCells[cellIndex];
                                    }
                                    else {
                                        delete this.selectedSheet._styledCells[newCellIndex];
                                    }
                                }
                            }
                        }
                    }
                    // Set the stored the style of the source cells to the destination cells.
                    for (rowIndex = this._dropRange.topRow; rowIndex <= this._dropRange.bottomRow; rowIndex++) {
                        for (colIndex = this._dropRange.leftCol; colIndex <= this._dropRange.rightCol; colIndex++) {
                            cellIndex = rowIndex * this.columns.length + colIndex;
                            if (srcCellStyles[index]) {
                                this.selectedSheet._styledCells[cellIndex] = srcCellStyles[index];
                            }
                            else {
                                delete this.selectedSheet._styledCells[cellIndex];
                            }
                            index++;
                        }
                    }
                };
                // Check whether the specific cell range contains merged cells.
                FlexSheet.prototype._containsMergedCells = function (rng) {
                    var rowIndex, colIndex, cellIndex, mergedRange;
                    if (!this.selectedSheet) {
                        return false;
                    }
                    for (rowIndex = rng.topRow; rowIndex <= rng.bottomRow; rowIndex++) {
                        for (colIndex = rng.leftCol; colIndex <= rng.rightCol; colIndex++) {
                            cellIndex = rowIndex * this.columns.length + colIndex;
                            mergedRange = this.selectedSheet._mergedRanges[cellIndex];
                            if (mergedRange && mergedRange.isValid && !mergedRange.isSingleCell) {
                                return true;
                            }
                        }
                    }
                    return false;
                };
                // Multiple select columns processing.
                FlexSheet.prototype._multiSelectColumns = function (ht) {
                    var range;
                    if (ht && this._columnHeaderClicked) {
                        range = new grid_1.CellRange(ht.row, ht.col);
                        range.row = 0;
                        range.row2 = this.rows.length - 1;
                        range.col2 = this.selection.col2;
                        this.select(range);
                    }
                };
                // Gets the absolute offset for the element.
                FlexSheet.prototype._cumulativeOffset = function (element) {
                    var top = 0, left = 0;
                    do {
                        top += element.offsetTop || 0;
                        left += element.offsetLeft || 0;
                        element = element.offsetParent;
                    } while (element);
                    return new wijmo.Point(left, top);
                };
                // Gets the absolute scroll offset for the element.
                FlexSheet.prototype._cumulativeScrollOffset = function (element) {
                    var scrollTop = 0, scrollLeft = 0;
                    do {
                        scrollTop += element.scrollTop || 0;
                        scrollLeft += element.scrollLeft || 0;
                        element = element.offsetParent;
                    } while (element && !(element instanceof HTMLBodyElement));
                    // Chrome and Safari always use document.body.scrollTop, 
                    // while IE and Firefox use document.body.scrollTop for quirks mode and document.documentElement.scrollTop for standard mode. 
                    // So we need check both the document.body.scrollTop and document.documentElement.scrollTop (TFS 142679)
                    scrollTop += document.body.scrollTop || document.documentElement.scrollTop;
                    scrollLeft += document.body.scrollLeft || document.documentElement.scrollLeft;
                    return new wijmo.Point(scrollLeft, scrollTop);
                };
                // Check whether current hit is within current selection.
                FlexSheet.prototype._checkHitWithinSelection = function (ht) {
                    var cellIndex, mergedRange;
                    if (ht != null && ht.cellType === grid_1.CellType.Cell) {
                        mergedRange = this.getMergedRange(this.cells, ht.row, ht.col);
                        if (mergedRange && mergedRange.intersects(this.selection)) {
                            return true;
                        }
                        if (this.selection.row === ht.row && this.selection.col === ht.col) {
                            return true;
                        }
                    }
                    return false;
                };
                // Clear the merged cells, styled cells and selection for the empty sheet.
                FlexSheet.prototype._clearForEmptySheet = function (rowsOrColumns) {
                    if (this.selectedSheet && this[rowsOrColumns].length === 0 && this._isCopyingOrUndoing !== true) {
                        this.selectedSheet._mergedRanges = null;
                        this.selectedSheet._styledCells = null;
                        this.select(new grid_1.CellRange());
                    }
                };
                // Check whether the specified cell range contains Group Row.
                FlexSheet.prototype._containsGroupRows = function (cellRange) {
                    var rowIndex, row;
                    for (rowIndex = cellRange.topRow; rowIndex <= cellRange.bottomRow; rowIndex++) {
                        row = this.rows[rowIndex];
                        if (row instanceof grid_1.GroupRow) {
                            return true;
                        }
                    }
                    return false;
                };
                // Delete the content of the selected cells.
                FlexSheet.prototype._delSeletionContent = function () {
                    var self = this, selections = self.selectedSheet.selectionRanges;
                    if (self.isReadOnly) {
                        return;
                    }
                    self.deferUpdate(function () {
                        var selection, index, colIndex, rowIndex, bcol, contentDeleted = false, delAction = new sheet_1._EditAction(self);
                        for (index = 0; index < selections.length; index++) {
                            selection = selections[index];
                            for (rowIndex = selection.topRow; rowIndex <= selection.bottomRow; rowIndex++) {
                                for (colIndex = selection.leftCol; colIndex <= selection.rightCol; colIndex++) {
                                    bcol = self._getBindingColumn(self.cells, rowIndex, self.columns[colIndex]);
                                    if (bcol.isRequired == false || (bcol.isRequired == null && bcol.dataType == wijmo.DataType.String)) {
                                        if (self.getCellData(rowIndex, colIndex, true)) {
                                            self.setCellData(rowIndex, colIndex, '', true);
                                            contentDeleted = true;
                                        }
                                    }
                                }
                            }
                        }
                        if (contentDeleted) {
                            delAction.saveNewState();
                            self._undoStack._addAction(delAction);
                        }
                    });
                };
                // Update the affected formulas for inserting/removing row/columns.
                FlexSheet.prototype._updateAffectedFormula = function (index, count, isAdding, isRow) {
                    var rowIndex, colIndex, newRowIndex, newColIndex, cellData, matches, cellRefIndex, isUpdated, cellRef, updatedCellRef, oldFormulas = [], newFormulas = [], cellAddress;
                    for (rowIndex = 0; rowIndex < this.rows.length; rowIndex++) {
                        for (colIndex = 0; colIndex < this.columns.length; colIndex++) {
                            var cellData = this.getCellData(rowIndex, colIndex, false);
                            if (!!cellData && typeof cellData === 'string' && cellData[0] === '=') {
                                matches = cellData.match(/(?=\b\D)\$?[A-Za-z]+\$?\d+/g);
                                if (!!matches && matches.length > 0) {
                                    isUpdated = false;
                                    for (cellRefIndex = 0; cellRefIndex < matches.length; cellRefIndex++) {
                                        cellRef = matches[cellRefIndex];
                                        if (cellRef.toLowerCase() !== 'atan2') {
                                            cellAddress = wijmo.xlsx.Workbook.tableAddress(cellRef);
                                            if (isRow) {
                                                if (cellAddress.row > index) {
                                                    if (isAdding) {
                                                        cellAddress.row += count;
                                                    }
                                                    else {
                                                        cellAddress.row -= count;
                                                    }
                                                    if (!isUpdated) {
                                                        isUpdated = true;
                                                        oldFormulas.push({
                                                            point: new wijmo.Point(rowIndex, colIndex),
                                                            formula: cellData
                                                        });
                                                    }
                                                }
                                            }
                                            else {
                                                if (cellAddress.col > index) {
                                                    if (isAdding) {
                                                        cellAddress.col += count;
                                                    }
                                                    else {
                                                        cellAddress.col -= count;
                                                    }
                                                    if (!isUpdated) {
                                                        isUpdated = true;
                                                        oldFormulas.push({
                                                            point: new wijmo.Point(rowIndex, colIndex),
                                                            formula: cellData
                                                        });
                                                    }
                                                }
                                            }
                                            updatedCellRef = wijmo.xlsx.Workbook.xlsxAddress(cellAddress.row, cellAddress.col, cellAddress.absRow, cellAddress.absCol);
                                            cellData = cellData.replace(cellRef, updatedCellRef);
                                        }
                                    }
                                    if (isUpdated) {
                                        this.setCellData(rowIndex, colIndex, cellData);
                                        newRowIndex = rowIndex;
                                        newColIndex = colIndex;
                                        if (isRow) {
                                            if (rowIndex > index) {
                                                if (isAdding) {
                                                    newRowIndex += count;
                                                }
                                                else {
                                                    newRowIndex -= count;
                                                }
                                            }
                                        }
                                        else {
                                            if (colIndex > index) {
                                                if (isAdding) {
                                                    newColIndex += count;
                                                }
                                                else {
                                                    newColIndex -= count;
                                                }
                                            }
                                        }
                                        newFormulas.push({
                                            point: new wijmo.Point(newRowIndex, newColIndex),
                                            formula: cellData
                                        });
                                    }
                                }
                            }
                        }
                    }
                    return {
                        oldFormulas: oldFormulas,
                        newFormulas: newFormulas
                    };
                };
                // Update the column filter for moving the column. 
                FlexSheet.prototype._updateColumnFiler = function (srcColIndex, descColIndex) {
                    var filterDef = JSON.parse(this._filter.filterDefinition);
                    for (var i = 0; i < filterDef.filters.length; i++) {
                        var filter = filterDef.filters[i];
                        if (filter.columnIndex === srcColIndex) {
                            filter.columnIndex = descColIndex;
                            break;
                        }
                    }
                    this._filter.filterDefinition = JSON.stringify(filterDef);
                };
                // Chech the specific element is the child of other element.
                FlexSheet.prototype._isDescendant = function (paranet, child) {
                    var node = child.parentNode;
                    while (node != null) {
                        if (node === paranet) {
                            return true;
                        }
                        node = node.parentNode;
                    }
                    return false;
                };
                // Clear the expression cache of the CalcEngine.
                FlexSheet.prototype._clearCalcEngine = function () {
                    this._calcEngine._clearExpressionCache();
                };
                /**
                 * Converts the number value to its corresponding alpha value.
                 * For instance: 0, 1, 2...to a, b, c...
                 * @param c The number value need to be converted.
                 */
                FlexSheet.convertNumberToAlpha = function (c) {
                    var content = '', dCount, pos;
                    if (c >= 0) {
                        do {
                            dCount = Math.floor(c / 26);
                            pos = c % 26;
                            content = String.fromCharCode(pos + 65) + content;
                            c = dCount - 1;
                        } while (dCount);
                    }
                    return content;
                };
                /**
                 * Overrides the template used to instantiate @see:FlexSheet control.
                 */
                FlexSheet.controlTemplate = '<div style="width:100%;height:100%">' +
                    '<div wj-part="container" style="width:100%">' +
                    grid_1.FlexGrid.controlTemplate +
                    '</div>' +
                    '<div wj-part="tab-holder" style="width:100%; min-width:100px">' +
                    '</div>' +
                    '<div wj-part="context-menu" style="display:none;z-index:100"></div>' +
                    '</div>';
                return FlexSheet;
            }(grid_1.FlexGrid));
            sheet_1.FlexSheet = FlexSheet;
            /**
             * Provides arguments for the @see:FlexSheet.draggingRowColumn event.
             */
            var DraggingRowColumnEventArgs = (function (_super) {
                __extends(DraggingRowColumnEventArgs, _super);
                /**
                 * Initializes a new instance of the @see:DraggingRowColumnEventArgs class.
                 *
                 * @param isDraggingRows Indicates whether the dragging event is triggered due to dragging rows or columns.
                 * @param isShiftKey Indicates whether the shift key is pressed when dragging.
                 */
                function DraggingRowColumnEventArgs(isDraggingRows, isShiftKey) {
                    _super.call(this);
                    this._isDraggingRows = isDraggingRows;
                    this._isShiftKey = isShiftKey;
                }
                Object.defineProperty(DraggingRowColumnEventArgs.prototype, "isDraggingRows", {
                    /**
                     * Gets a value indicating whether the event refers to dragging rows or columns.
                     */
                    get: function () {
                        return this._isDraggingRows;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(DraggingRowColumnEventArgs.prototype, "isShiftKey", {
                    /**
                     * Gets a value indicating whether the shift key is pressed.
                     */
                    get: function () {
                        return this._isShiftKey;
                    },
                    enumerable: true,
                    configurable: true
                });
                return DraggingRowColumnEventArgs;
            }(wijmo.EventArgs));
            sheet_1.DraggingRowColumnEventArgs = DraggingRowColumnEventArgs;
            /**
             * Provides arguments for unknown function events.
             */
            var UnknownFunctionEventArgs = (function (_super) {
                __extends(UnknownFunctionEventArgs, _super);
                /**
                 * Initializes a new instance of the @see:UnknownFunctionEventArgs class.
                 *
                 * @param funcName The name of the unknown function.
                 * @param params The parameters' value list of the nuknown function.
                 */
                function UnknownFunctionEventArgs(funcName, params) {
                    _super.call(this);
                    this._funcName = funcName;
                    this._params = params;
                }
                Object.defineProperty(UnknownFunctionEventArgs.prototype, "funcName", {
                    /**
                     * Gets the name of the unknown function.
                     */
                    get: function () {
                        return this._funcName;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(UnknownFunctionEventArgs.prototype, "params", {
                    /**
                     * Gets the parameters' value list of the nuknown function.
                     */
                    get: function () {
                        return this._params;
                    },
                    enumerable: true,
                    configurable: true
                });
                return UnknownFunctionEventArgs;
            }(wijmo.EventArgs));
            sheet_1.UnknownFunctionEventArgs = UnknownFunctionEventArgs;
            /**
             * Defines the extension of the @see:GridPanel class, which is used by <b>FlexSheet</b> where
             * the base @see:FlexGrid class uses @see:GridPanel. For example, the <b>cells</b> property returns an instance
             * of this class.
             */
            var FlexSheetPanel = (function (_super) {
                __extends(FlexSheetPanel, _super);
                /**
                 * Initializes a new instance of the @see:FlexSheetPanel class.
                 *
                 * @param grid The @see:FlexGrid object that owns the panel.
                 * @param cellType The type of cell in the panel.
                 * @param rows The rows displayed in the panel.
                 * @param cols The columns displayed in the panel.
                 * @param element The HTMLElement that hosts the cells in the control.
                 */
                function FlexSheetPanel(grid, cellType, rows, cols, element) {
                    _super.call(this, grid, cellType, rows, cols, element);
                }
                /**
                 * Gets a @see:SelectedState value that indicates the selected state of a cell.
                 *
                 * Overrides this method to support multiple selection showSelectedHeaders for @see:FlexSheet
                 *
                 * @param r Specifies Row index of the cell.
                 * @param c Specifies Column index of the cell.
                 * @param rng @see:CellRange that contains the cell that would be included.
                 */
                FlexSheetPanel.prototype.getSelectedState = function (r, c, rng) {
                    var selections, selectionCnt, index, selection, selectedState, mergedRange;
                    if (!this.grid) {
                        return undefined;
                    }
                    mergedRange = this.grid.getMergedRange(this, r, c);
                    selections = this.grid.selectedSheet ? this.grid.selectedSheet.selectionRanges : null;
                    selectedState = _super.prototype.getSelectedState.call(this, r, c, rng);
                    selectionCnt = selections ? selections.length : 0;
                    if (selectedState === grid_1.SelectedState.None && selectionCnt > 0 && this.grid._enableMulSel) {
                        for (index = 0; index < selections.length; index++) {
                            selection = selections[index];
                            if (selection && selection instanceof grid_1.CellRange) {
                                if (this.cellType === grid_1.CellType.Cell) {
                                    if (mergedRange) {
                                        if (mergedRange.contains(selection.row, selection.col)) {
                                            if (index === selectionCnt - 1 && !this.grid._isClicking) {
                                                return this.grid.showMarquee ? grid_1.SelectedState.None : grid_1.SelectedState.Cursor;
                                            }
                                            return grid_1.SelectedState.Selected;
                                        }
                                        if (mergedRange.intersects(selection)) {
                                            return grid_1.SelectedState.Selected;
                                        }
                                    }
                                    if (selection.row === r && selection.col === c) {
                                        if (index === selectionCnt - 1 && !this.grid._isClicking) {
                                            return this.grid.showMarquee ? grid_1.SelectedState.None : grid_1.SelectedState.Cursor;
                                        }
                                        return grid_1.SelectedState.Selected;
                                    }
                                    if (selection.contains(r, c)) {
                                        return grid_1.SelectedState.Selected;
                                    }
                                }
                                if (this.grid.showSelectedHeaders & grid_1.HeadersVisibility.Row
                                    && this.cellType === grid_1.CellType.RowHeader
                                    && selection.containsRow(r)) {
                                    return grid_1.SelectedState.Selected;
                                }
                                if (this.grid.showSelectedHeaders & grid_1.HeadersVisibility.Column
                                    && this.cellType === grid_1.CellType.ColumnHeader
                                    && selection.containsColumn(c)) {
                                    return grid_1.SelectedState.Selected;
                                }
                            }
                        }
                    }
                    return selectedState;
                };
                /**
                 * Sets the content of a cell in the panel.
                 *
                 * @param r The index of the row that contains the cell.
                 * @param c The index, name, or binding of the column that contains the cell.
                 * @param value The value to store in the cell.
                 * @param coerce A value indicating whether to change the value automatically to match the column's data type.
                 * @return Returns true if the value is stored successfully, otherwise false (failed cast).
                 */
                FlexSheetPanel.prototype.setCellData = function (r, c, value, coerce) {
                    if (coerce === void 0) { coerce = true; }
                    var parsedDateVal;
                    if (coerce && value && wijmo.isString(value)) {
                        if (!isNaN(+value)) {
                            value = +value;
                        }
                        else if (value[0] !== '=') {
                            parsedDateVal = wijmo.Globalize.parseDate(value, '');
                            if (parsedDateVal) {
                                value = parsedDateVal;
                            }
                        }
                    }
                    // When the cell data is formula, we shall not force to change the data type of the cell data.
                    if (value && wijmo.isString(value) && value[0] === '=') {
                        coerce = false;
                    }
                    return _super.prototype.setCellData.call(this, r, c, value, coerce);
                };
                // renders a cell
                // It overrides the _renderCell method of the parent class GridPanel.
                FlexSheetPanel.prototype._renderCell = function (r, c, vrng, state, ctr) {
                    var cell = this.hostElement.childNodes[ctr], cellStyle, cellIndex = r * this.grid.columns.length + c, mr = this.grid.getMergedRange(this, r, c);
                    ctr = _super.prototype._renderCell.call(this, r, c, vrng, state, ctr);
                    if (this.cellType !== wijmo.grid.CellType.Cell) {
                        return ctr;
                    }
                    // skip over cells that have been merged over
                    if (mr) {
                        if (cellIndex > mr.topRow * this.grid.columns.length + mr.leftCol) {
                            return ctr;
                        }
                    }
                    if (wijmo.hasClass(cell, 'wj-state-selected') || wijmo.hasClass(cell, 'wj-state-multi-selected')) {
                        // If the cell is selected state, we'll remove the custom background color and font color style.
                        cell.style.backgroundColor = '';
                        cell.style.color = '';
                    }
                    else if (this.grid.selectedSheet) {
                        // If the cell removes selected state, we'll resume the custom background color and font color style.
                        cellStyle = this.grid.selectedSheet._styledCells[cellIndex];
                        if (cell && cellStyle) {
                            cell.style.backgroundColor = cellStyle.backgroundColor;
                            cell.style.color = cellStyle.color;
                        }
                    }
                    return ctr;
                };
                return FlexSheetPanel;
            }(grid_1.GridPanel));
            sheet_1.FlexSheetPanel = FlexSheetPanel;
            /**
             * Represents a row used to display column header information for a bound sheet.
             */
            var HeaderRow = (function (_super) {
                __extends(HeaderRow, _super);
                /**
                * Initializes a new instance of the HeaderRow class.
                */
                function HeaderRow() {
                    _super.call(this);
                    this.isReadOnly = true;
                }
                return HeaderRow;
            }(grid_1.Row));
            sheet_1.HeaderRow = HeaderRow;
        })(sheet = grid_1.sheet || (grid_1.sheet = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=FlexSheet.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid_1) {
        var sheet;
        (function (sheet_1) {
            'use strict';
            /**
             * Represents a sheet within the @see:FlexSheet control.
             */
            var Sheet = (function () {
                /**
                 * Initializes a new instance of the @see:FlexSheet class.
                 *
                 * @param owner The owner @see: FlexSheet control.
                 * @param grid The associated @see:FlexGrid control used to store the sheet data. If not specified then the
                 * new <b>FlexGrid</b> control will be created.
                 * @param sheetName The name of the sheet within the @see:FlexSheet control.
                 * @param rows The row count for the sheet.
                 * @param cols The column count for the sheet.
                 */
                function Sheet(owner, grid, sheetName, rows, cols) {
                    this._visible = true;
                    this._unboundSortDesc = new wijmo.collections.ObservableArray();
                    this._currentStyledCells = {};
                    this._currentMergedRanges = {};
                    this._isEmptyGrid = false;
                    this._scrollPosition = new wijmo.Point();
                    this._freezeHiddenRowCnt = 0;
                    this._freezeHiddenColumnCnt = 0;
                    /**
                     * Occurs after the sheet name has changed.
                     */
                    this.nameChanged = new wijmo.Event();
                    /**
                     * Occurs after the visible of sheet has changed.
                     */
                    this.visibleChanged = new wijmo.Event();
                    var self = this, insertRows, insertCols, i;
                    self._owner = owner;
                    self._name = sheetName;
                    if (wijmo.isNumber(rows) && !isNaN(rows) && rows >= 0) {
                        self._rowCount = rows;
                    }
                    else {
                        self._rowCount = 200;
                    }
                    if (wijmo.isNumber(cols) && !isNaN(cols) && cols >= 0) {
                        self._columnCount = cols;
                    }
                    else {
                        self._columnCount = 20;
                    }
                    self._grid = grid || this._createGrid();
                    self._grid.itemsSourceChanged.addHandler(this._gridItemsSourceChanged, this);
                    self._unboundSortDesc.collectionChanged.addHandler(function () {
                        var arr = self._unboundSortDesc, i, sd;
                        for (i = 0; i < arr.length; i++) {
                            sd = wijmo.tryCast(arr[i], _UnboundSortDescription);
                            if (!sd) {
                                throw 'sortDescriptions array must contain SortDescription objects.';
                            }
                        }
                        if (self._owner) {
                            self._owner.rows.beginUpdate();
                            self._owner.rows.sort(self._compareRows());
                            self._owner.rows.endUpdate();
                            self._owner.rows._dirty = true;
                            self._owner.rows._update();
                            //Synch with current sheet.
                            if (self._owner.selectedSheet) {
                                self._owner._copyTo(self._owner.selectedSheet);
                                self._owner._copyFrom(self._owner.selectedSheet);
                            }
                        }
                    });
                }
                Object.defineProperty(Sheet.prototype, "grid", {
                    /**
                     * Gets the associated @see:FlexGrid control used to store the sheet data.
                     */
                    get: function () {
                        return this._grid;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(Sheet.prototype, "name", {
                    /**
                     * Gets or sets the name of the sheet.
                     */
                    get: function () {
                        return this._name;
                    },
                    set: function (value) {
                        if (!wijmo.isNullOrWhiteSpace(value) && ((this._name && this._name.toLowerCase() !== value.toLowerCase()) || !this._name)) {
                            this._name = value;
                            this._grid['wj_sheetInfo'].name = value;
                            this.onNameChanged(new wijmo.EventArgs());
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(Sheet.prototype, "visible", {
                    /**
                     * Gets or sets the sheet visibility.
                     */
                    get: function () {
                        return this._visible;
                    },
                    set: function (value) {
                        if (this._visible !== value) {
                            this._visible = value;
                            this._grid['wj_sheetInfo'].visible = value;
                            this.onVisibleChanged(new wijmo.EventArgs());
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(Sheet.prototype, "rowCount", {
                    /**
                     * Gets or sets the number of rows in the sheet.
                     */
                    get: function () {
                        if (this._grid != null) {
                            return this._grid.rows.length;
                        }
                        return 0;
                    },
                    set: function (value) {
                        var rowIndex;
                        if (wijmo.isNumber(value) && !isNaN(value) && value >= 0 && this._rowCount !== value) {
                            if (this._rowCount < value) {
                                for (rowIndex = 0; rowIndex < (value - this._rowCount); rowIndex++) {
                                    this._grid.rows.push(new grid_1.Row());
                                }
                            }
                            else {
                                this._grid.rows.splice(value, this._rowCount - value);
                            }
                            this._rowCount = value;
                            // If the sheet is current selected sheet of the flexsheet, we should synchronize the updating of the sheet to the flexsheet.
                            if (this._owner && this._owner.selectedSheet && this._name === this._owner.selectedSheet.name) {
                                this._owner._copyFrom(this, true);
                            }
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(Sheet.prototype, "columnCount", {
                    /**
                     * Gets or sets the number of columns in the sheet.
                     */
                    get: function () {
                        if (this._grid != null) {
                            return this._grid.columns.length;
                        }
                        return 0;
                    },
                    set: function (value) {
                        var colIndex;
                        if (wijmo.isNumber(value) && !isNaN(value) && value >= 0 && this._columnCount !== value) {
                            if (this._columnCount < value) {
                                for (colIndex = 0; colIndex < (value - this._columnCount); colIndex++) {
                                    this._grid.columns.push(new grid_1.Column());
                                }
                            }
                            else {
                                this._grid.columns.splice(value, this._columnCount - value);
                            }
                            this._columnCount = value;
                            // If the sheet is current seleced sheet of the flexsheet, we should synchronize the updating of the sheet to the flexsheet.
                            if (this._owner && this._owner.selectedSheet && this._name === this._owner.selectedSheet.name) {
                                this._owner._copyFrom(this, true);
                            }
                        }
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(Sheet.prototype, "selectionRanges", {
                    /**
                     * Gets the selection array.
                     */
                    get: function () {
                        var _this = this;
                        if (!this._selectionRanges) {
                            this._selectionRanges = new wijmo.collections.ObservableArray();
                            this._selectionRanges.collectionChanged.addHandler(function () {
                                var selectionCnt, lastSelection;
                                if (_this._owner && !_this._owner._isClicking) {
                                    selectionCnt = _this._selectionRanges.length;
                                    if (selectionCnt > 0) {
                                        lastSelection = _this._selectionRanges[selectionCnt - 1];
                                        if (lastSelection && lastSelection instanceof grid_1.CellRange) {
                                            _this._owner.selection = lastSelection;
                                        }
                                    }
                                    if (selectionCnt > 1) {
                                        _this._owner._enableMulSel = true;
                                    }
                                    _this._owner.refresh();
                                    _this._owner._enableMulSel = false;
                                }
                            }, this);
                        }
                        return this._selectionRanges;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(Sheet.prototype, "itemsSource", {
                    /**
                     * Gets or sets the array or @see:ICollectionView for the @see:FlexGrid instance of the sheet.
                     */
                    get: function () {
                        if (this._grid != null) {
                            return this._grid.itemsSource;
                        }
                        return null;
                    },
                    set: function (value) {
                        if (this._grid == null) {
                            this._createGrid();
                            this._grid.itemsSourceChanged.addHandler(this._gridItemsSourceChanged, this);
                        }
                        if (this._isEmptyGrid) {
                            this._clearGrid();
                        }
                        this._grid.itemsSource = value;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(Sheet.prototype, "_styledCells", {
                    /*
                     * Gets or sets the styled cells
                     * This property uses the cell index as the key and stores the @ICellStyle object as the value.
                     * { 1: { fontFamily: xxxx, fontSize: xxxx, .... }, 2: {...}, ... }
                     */
                    get: function () {
                        if (!this._currentStyledCells) {
                            this._currentStyledCells = {};
                        }
                        return this._currentStyledCells;
                    },
                    set: function (value) {
                        this._currentStyledCells = value;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(Sheet.prototype, "_mergedRanges", {
                    /*
                     * Gets or sets the merge ranges.
                     * This property uses the cell index as the key and stores the @CellRange object as the value.
                     * { 1: CellRange(row = 1, col = 1, row2 = 3, col2 = 4), 2: CellRange(), ...}
                     */
                    get: function () {
                        if (!this._currentMergedRanges) {
                            this._currentMergedRanges = {};
                        }
                        return this._currentMergedRanges;
                    },
                    set: function (value) {
                        this._currentMergedRanges = value;
                    },
                    enumerable: true,
                    configurable: true
                });
                /**
                 * Raises the @see:nameChanged event.
                 */
                Sheet.prototype.onNameChanged = function (e) {
                    this.nameChanged.raise(this, e);
                };
                /**
                 * Raises the @see:visibleChanged event.
                 */
                Sheet.prototype.onVisibleChanged = function (e) {
                    this.visibleChanged.raise(this, e);
                };
                /**
                 * Gets the style of specified cell.
                 *
                 * @param rowIndex the row index of the specified cell.
                 * @param columnIndex the column index of the specified cell.
                 */
                Sheet.prototype.getCellStyle = function (rowIndex, columnIndex) {
                    var cellIndex, rowCnt = this._grid.rows.length, colCnt = this._grid.columns.length;
                    if (rowIndex >= rowCnt || columnIndex >= colCnt) {
                        return null;
                    }
                    cellIndex = rowIndex * colCnt + columnIndex;
                    return this._styledCells[cellIndex];
                };
                // Attach the sheet to the @see: FlexSheet control as owner.
                Sheet.prototype._attachOwner = function (owner) {
                    if (this._owner !== owner) {
                        this._owner = owner;
                    }
                };
                // Update the sheet name with valid name.
                Sheet.prototype._setValidName = function (validName) {
                    this._name = validName;
                    this._grid['wj_sheetInfo'].name = validName;
                };
                // comparison function used in rows sort for unbound sheet.
                Sheet.prototype._compareRows = function () {
                    var self = this, sortDesc = this._unboundSortDesc;
                    return function (a, b) {
                        for (var i = 0; i < sortDesc.length; i++) {
                            // get values
                            var sd = sortDesc[i], v1 = a._ubv ? a._ubv[sd.column._hash] : '', v2 = b._ubv ? b._ubv[sd.column._hash] : '';
                            // if the cell value is formula, we should try to evaluate this formula.
                            if (wijmo.isString(v1) && v1[0] === '=') {
                                v1 = self._owner.evaluate(v1);
                                if (!wijmo.isPrimitive(v1)) {
                                    v1 = v1.value;
                                }
                            }
                            if (wijmo.isString(v2) && v2[0] === '=') {
                                v2 = self._owner.evaluate(v2);
                                if (!wijmo.isPrimitive(v2)) {
                                    v2 = v2.value;
                                }
                            }
                            // check for NaN (isNaN returns true for NaN but also for non-numbers)
                            if (v1 !== v1)
                                v1 = null;
                            if (v2 !== v2)
                                v2 = null;
                            // ignore case when sorting  (but add the original string to keep the 
                            // strings different and the sort consistent, 'aa' between 'AA' and 'bb')
                            if (wijmo.isString(v1))
                                v1 = v1.toLowerCase() + v1;
                            if (wijmo.isString(v2))
                                v2 = v2.toLowerCase() + v2;
                            // compare the values (at last!)
                            var cmp = (v1 < v2) ? -1 : (v1 > v2) ? +1 : 0;
                            if (cmp !== 0) {
                                return sd.ascending ? +cmp : -cmp;
                            }
                        }
                        return 0;
                    };
                };
                // Create a blank flexsheet.
                Sheet.prototype._createGrid = function () {
                    var hostElement = document.createElement('div'), grid, column, colIndex, rowIndex;
                    this._isEmptyGrid = true;
                    // We should append the host element of the data grid of current sheet to body before creating data grid,
                    // this will make the host element to inherit the style of body (TFS 121713)
                    hostElement.style.visibility = 'hidden';
                    document.body.appendChild(hostElement);
                    grid = new grid_1.FlexGrid(hostElement);
                    document.body.removeChild(hostElement);
                    for (rowIndex = 0; rowIndex < this._rowCount; rowIndex++) {
                        grid.rows.push(new grid_1.Row());
                    }
                    for (colIndex = 0; colIndex < this._columnCount; colIndex++) {
                        column = new grid_1.Column();
                        // Setting the required property of the column to false for the data grid of current sheet.
                        // TFS #126125
                        column.isRequired = false;
                        grid.columns.push(column);
                    }
                    // Add header row for the grid of the bind sheet.
                    grid.loadedRows.addHandler(function () {
                        if (grid.itemsSource && !(grid.rows[0] instanceof sheet_1.HeaderRow)) {
                            grid.rows.insert(0, new sheet_1.HeaderRow());
                        }
                    });
                    // Add sheet related info into the flexgrid.
                    // This property contains the name, style of cells and merge cells of current sheet.
                    grid['wj_sheetInfo'] = {
                        name: this.name,
                        visible: this.visible,
                        styledCells: this._styledCells,
                        mergedRanges: this._mergedRanges
                    };
                    return grid;
                };
                // Clear the grid of the sheet.
                Sheet.prototype._clearGrid = function () {
                    this._grid.rows.clear();
                    this._grid.columns.clear();
                    this._grid.columnHeaders.columns.clear();
                    this._grid.rowHeaders.rows.clear();
                };
                // Items source changed handler for the grid of the sheet.
                Sheet.prototype._gridItemsSourceChanged = function () {
                    // If the sheet is current seleced sheet of the flexsheet, we should synchronize the updating of the sheet to the flexsheet.
                    if (this._owner && this._owner.selectedSheet && this._name === this._owner.selectedSheet.name) {
                        this._owner._copyFrom(this, false);
                    }
                };
                return Sheet;
            }());
            sheet_1.Sheet = Sheet;
            /**
             * Defines the collection of the @see:Sheet objects.
             */
            var SheetCollection = (function (_super) {
                __extends(SheetCollection, _super);
                function SheetCollection() {
                    _super.apply(this, arguments);
                    this._current = -1;
                    /**
                     * Occurs when the @see:SheetCollection is cleared.
                     */
                    this.sheetCleared = new wijmo.Event();
                    /**
                     * Occurs when the <b>selectedIndex</b> property changes.
                     */
                    this.selectedSheetChanged = new wijmo.Event();
                    /**
                     * Occurs after the name of the sheet in the collection has changed.
                     */
                    this.sheetNameChanged = new wijmo.Event();
                    /**
                     * Occurs after the visible of the sheet in the collection has changed.
                     */
                    this.sheetVisibleChanged = new wijmo.Event();
                }
                /**
                 * Raises the sheetCleared event.
                 */
                SheetCollection.prototype.onSheetCleared = function () {
                    this.sheetCleared.raise(this, new wijmo.EventArgs());
                };
                Object.defineProperty(SheetCollection.prototype, "selectedIndex", {
                    /**
                     * Gets or sets the index of the currently selected sheet.
                     */
                    get: function () {
                        return this._current;
                    },
                    set: function (index) {
                        this._moveCurrentTo(index);
                    },
                    enumerable: true,
                    configurable: true
                });
                /**
                 * Raises the <b>currentChanged</b> event.
                 *
                 * @param e @see:PropertyChangedEventArgs that contains the event data.
                 */
                SheetCollection.prototype.onSelectedSheetChanged = function (e) {
                    this.selectedSheetChanged.raise(this, e);
                };
                /**
                 * Inserts an item at a specific position in the array.
                 * Overrides the insert method of its base class @see:ObservableArray.
                 *
                 * @param index Position where the item will be added.
                 * @param item Item to add to the array.
                 */
                SheetCollection.prototype.insert = function (index, item) {
                    var name;
                    name = item.name ? this.getValidSheetName(item) : this._getUniqueName();
                    if (name !== item.name) {
                        item.name = name;
                    }
                    _super.prototype.insert.call(this, index, item);
                    this._postprocessSheet(item);
                };
                /**
                 * Adds one or more items to the end of the array.
                 * Overrides the push method of its base class @see:ObservableArray.
                 *
                 * @param ...item One or more items to add to the array.
                 * @return The new length of the array.
                 */
                SheetCollection.prototype.push = function () {
                    var item = [];
                    for (var _i = 0; _i < arguments.length; _i++) {
                        item[_i - 0] = arguments[_i];
                    }
                    var currentLength = this.length, idx = 0, name;
                    for (; idx < item.length; idx++) {
                        name = item[idx].name ? this.getValidSheetName(item[idx]) : this._getUniqueName();
                        if (name !== item[idx].name) {
                            item[idx].name = name;
                        }
                        _super.prototype.push.call(this, item[idx]);
                        this._postprocessSheet(item[idx]);
                    }
                    return this.length;
                };
                /**
                 * Removes and/or adds items to the array.
                 * Overrides the splice method of its base class @see:ObservableArray.
                 *
                 * @param index Position where items will be added or removed.
                 * @param count Number of items to remove from the array.
                 * @param item Item to add to the array.
                 * @return An array containing the removed elements.
                 */
                SheetCollection.prototype.splice = function (index, count, item) {
                    var name;
                    if (item) {
                        name = item.name ? this.getValidSheetName(item) : this._getUniqueName();
                        if (name !== item.name) {
                            item.name = name;
                        }
                        this._postprocessSheet(item);
                        return _super.prototype.splice.call(this, index, count, item);
                    }
                    else {
                        return _super.prototype.splice.call(this, index, count, item);
                    }
                };
                /**
                 * Removes an item at a specific position in the array.
                 * Overrides the removeAt method of its base class @see:ObservableArray.
                 *
                 * @param index Position of the item to remove.
                 */
                SheetCollection.prototype.removeAt = function (index) {
                    var succeeded = this.hide(index);
                    if (succeeded) {
                        _super.prototype.removeAt.call(this, index);
                        if (index < this.selectedIndex) {
                            this._current -= 1;
                        }
                    }
                };
                /**
                 * Raises the <b>sheetNameChanged</b> event.
                 */
                SheetCollection.prototype.onSheetNameChanged = function (e) {
                    this.sheetNameChanged.raise(this, e);
                };
                /**
                 * Raises the <b>sheetVisibleChanged</b> event.
                 */
                SheetCollection.prototype.onSheetVisibleChanged = function (e) {
                    this.sheetVisibleChanged.raise(this, e);
                };
                /**
                 * Selects the first sheet in the @see:FlexSheet control.
                 */
                SheetCollection.prototype.selectFirst = function () {
                    return this._moveCurrentTo(0);
                };
                /**
                 * Selects the last sheet in the owner @see:FlexSheet control.
                 */
                SheetCollection.prototype.selectLast = function () {
                    return this._moveCurrentTo(this.length - 1);
                };
                /**
                 * Selects the previous sheet in the owner @see:FlexSheet control.
                 */
                SheetCollection.prototype.selectPrevious = function () {
                    return this._moveCurrentTo(this._current - 1);
                };
                /**
                 * Select the next sheet in the owner @see:FlexSheet control.
                 */
                SheetCollection.prototype.selectNext = function () {
                    return this._moveCurrentTo(this._current + 1);
                };
                /**
                 * Hides the sheet at the specified position.
                 *
                 * @param pos The position of the sheet to hide.
                 */
                SheetCollection.prototype.hide = function (pos) {
                    var succeeded = false;
                    if (pos < 0 && pos >= this.length) {
                        return false;
                    }
                    if (!this[pos].visible) {
                        return false;
                    }
                    this[pos].visible = false;
                    return true;
                };
                /**
                 * Unhide and selects the @see:Sheet at the specified position.
                 *
                 * @param pos The position of the sheet to show.
                 */
                SheetCollection.prototype.show = function (pos) {
                    var succeeded = false;
                    if (pos < 0 && pos >= this.length) {
                        return false;
                    }
                    this[pos].visible = true;
                    this._moveCurrentTo(pos);
                    return true;
                };
                /**
                 * Clear the SheetCollection.
                 */
                SheetCollection.prototype.clear = function () {
                    _super.prototype.clear.call(this);
                    this._current = -1;
                    this.onSheetCleared();
                };
                /**
                 * Checks whether the sheet name is valid.
                 *
                 * @param sheet The @see:Sheet for which the name needs to check.
                 */
                SheetCollection.prototype.isValidSheetName = function (sheet) {
                    var sheetIndex = this._getSheetIndexFrom(sheet.name), currentSheetIndex = this.indexOf(sheet);
                    return (sheetIndex === -1 || sheetIndex === currentSheetIndex);
                };
                /**
                 * Gets the valid name for the sheet.
                 *
                 * @param currentSheet The @see:Sheet need get the valid name.
                 */
                SheetCollection.prototype.getValidSheetName = function (currentSheet) {
                    var validName = currentSheet.name, index = 1, currentSheetIndex = this.indexOf(currentSheet), sheetIndex;
                    do {
                        sheetIndex = this._getSheetIndexFrom(validName);
                        if (sheetIndex === -1 || sheetIndex === currentSheetIndex) {
                            break;
                        }
                        else {
                            validName = currentSheet.name.concat((index + 1).toString());
                        }
                        index = index + 1;
                    } while (true);
                    return validName;
                };
                // Move the current index to indicated position.
                SheetCollection.prototype._moveCurrentTo = function (pos) {
                    var searchedPos = pos, e;
                    if (pos < 0 || pos >= this.length) {
                        return false;
                    }
                    if (this._current < searchedPos || searchedPos === 0) {
                        while (searchedPos < this.length && !this[searchedPos].visible) {
                            searchedPos++;
                        }
                    }
                    else if (this._current > searchedPos) {
                        while (searchedPos >= 0 && !this[searchedPos].visible) {
                            searchedPos--;
                        }
                    }
                    if (searchedPos === this.length) {
                        searchedPos = pos;
                        while (searchedPos >= 0 && !this[searchedPos].visible) {
                            searchedPos--;
                        }
                    }
                    if (searchedPos < 0) {
                        return false;
                    }
                    if (searchedPos !== this._current) {
                        e = new wijmo.PropertyChangedEventArgs('sheetIndex', this._current, searchedPos);
                        this._current = searchedPos;
                        this.onSelectedSheetChanged(e);
                    }
                    return true;
                };
                // Get the index for the sheet in the SheetCollection.
                SheetCollection.prototype._getSheetIndexFrom = function (sheetName) {
                    var result = -1, sheet, name;
                    if (!sheetName) {
                        return result;
                    }
                    sheetName = sheetName.toLowerCase();
                    for (var i = 0; i < this.length; i++) {
                        sheet = this[i];
                        name = sheet.name ? sheet.name.toLowerCase() : '';
                        if (name === sheetName) {
                            return i;
                        }
                    }
                    return result;
                };
                // Post process the newly added sheet. 
                SheetCollection.prototype._postprocessSheet = function (item) {
                    var self = this;
                    // Update the sheet name via the sheetNameChanged event handler.
                    item.nameChanged.addHandler(function () {
                        var e, index = self._getSheetIndexFrom(item.name);
                        if (!self.isValidSheetName(item)) {
                            item._setValidName(self.getValidSheetName(item));
                        }
                        e = new wijmo.collections.NotifyCollectionChangedEventArgs(wijmo.collections.NotifyCollectionChangedAction.Change, item, wijmo.isNumber(index) ? index : self.length - 1);
                        self.onSheetNameChanged(e);
                    });
                    item.visibleChanged.addHandler(function () {
                        var index = self._getSheetIndexFrom(item.name), e = new wijmo.collections.NotifyCollectionChangedEventArgs(wijmo.collections.NotifyCollectionChangedAction.Change, item, wijmo.isNumber(index) ? index : self.length - 1);
                        self.onSheetVisibleChanged(e);
                    });
                };
                // Get the unique name for the sheet in the SheetCollection.
                SheetCollection.prototype._getUniqueName = function () {
                    var validName = 'Sheet1', index = 0;
                    do {
                        if (this._getSheetIndexFrom(validName) === -1) {
                            break;
                        }
                        else {
                            validName = 'Sheet'.concat((index + 1).toString());
                        }
                        index = index + 1;
                    } while (true);
                    return validName;
                };
                return SheetCollection;
            }(wijmo.collections.ObservableArray));
            sheet_1.SheetCollection = SheetCollection;
            /*
             * Represents the control that shows tabs for switching between @see:FlexSheet sheets.
             */
            var _SheetTabs = (function (_super) {
                __extends(_SheetTabs, _super);
                /*
                 * Initializes a new instance of the @see:_SheetTabs class.
                 *
                 * @param element The DOM element that will host the control, or a selector for the host element (e.g. '#theCtrl').
                 * @param owner The @see: FlexSheet control what the SheetTabs control works with.
                 * @param options JavaScript object containing initialization data for the control.
                 */
                function _SheetTabs(element, owner, options) {
                    _super.call(this, element, options);
                    this._rtl = false;
                    this._sheetTabClicked = false;
                    var self = this;
                    self._owner = owner;
                    self._sheets = owner.sheets;
                    self._rtl = getComputedStyle(self._owner.hostElement).direction == 'rtl';
                    if (self.hostElement.attributes['tabindex']) {
                        self.hostElement.attributes.removeNamedItem('tabindex');
                    }
                    self._initControl();
                    self.deferUpdate(function () {
                        if (options) {
                            self.initialize(options);
                        }
                    });
                }
                /*
                 * Override to refresh the control.
                 *
                 * @param fullUpdate Whether to update the control layout as well as the content.
                 */
                _SheetTabs.prototype.refresh = function (fullUpdate) {
                    this._tabContainer.innerHTML = '';
                    this._tabContainer.innerHTML = this._getSheetTabs();
                    if (this._rtl) {
                        this._adjustSheetsPosition();
                    }
                    this._adjustSize();
                };
                // The items source changed event handler.
                _SheetTabs.prototype._sourceChanged = function (sender, e) {
                    if (e === void 0) { e = wijmo.collections.NotifyCollectionChangedEventArgs.reset; }
                    var eArgs = e, index;
                    switch (eArgs.action) {
                        case wijmo.collections.NotifyCollectionChangedAction.Add:
                            index = eArgs.index - 1;
                            if (index < 0) {
                                index = 0;
                            }
                            this._tabContainer.innerHTML = '';
                            this._tabContainer.innerHTML = this._getSheetTabs();
                            if (this._rtl) {
                                this._adjustSheetsPosition();
                            }
                            this._adjustSize();
                            break;
                        case wijmo.collections.NotifyCollectionChangedAction.Remove:
                            this._tabContainer.removeChild(this._tabContainer.children[eArgs.index]);
                            if (this._tabContainer.hasChildNodes()) {
                                this._updateTabActive(eArgs.index, true);
                            }
                            this._adjustSize();
                            break;
                        default:
                            this.invalidate();
                            break;
                    }
                };
                // The current changed of the item source event handler.
                _SheetTabs.prototype._selectedSheetChanged = function (sender, e) {
                    this._updateTabActive(e.oldValue, false);
                    this._updateTabActive(e.newValue, true);
                    if (this._sheetTabClicked) {
                        this._sheetTabClicked = false;
                    }
                    else {
                        this._scrollToActiveSheet(e.newValue, e.oldValue);
                    }
                    this._adjustSize();
                };
                // Initialize the SheetTabs control.
                _SheetTabs.prototype._initControl = function () {
                    var self = this;
                    //apply template
                    self.applyTemplate('', self.getTemplate(), {
                        _sheetContainer: 'sheet-container',
                        _tabContainer: 'container',
                        _sheetPage: 'sheet-page',
                        _newSheet: 'new-sheet'
                    });
                    //init opts
                    if (self._rtl) {
                        self._sheetPage.style.right = '0px';
                        self._tabContainer.parentElement.style.right = self._sheetPage.clientWidth + 'px';
                        self._tabContainer.style.right = '0px';
                        self._tabContainer.style.cssFloat = 'right';
                        self._newSheet.style.right = (self._sheetPage.clientWidth + self._tabContainer.parentElement.clientWidth) + 'px';
                    }
                    self.addEventListener(self._newSheet, 'click', function (evt) {
                        var oldIndex = self._owner.selectedSheetIndex;
                        self._owner.addUnboundSheet();
                        self._scrollToActiveSheet(self._owner.selectedSheetIndex, oldIndex);
                    });
                    self._sheets.collectionChanged.addHandler(self._sourceChanged, self);
                    self._sheets.selectedSheetChanged.addHandler(self._selectedSheetChanged, self);
                    self._sheets.sheetNameChanged.addHandler(self._updateSheetName, self);
                    self._sheets.sheetVisibleChanged.addHandler(self._updateTabShown, self);
                    self._initSheetPage();
                    self._initSheetTab();
                };
                // Initialize the sheet tab part.
                _SheetTabs.prototype._initSheetTab = function () {
                    var self = this;
                    self.addEventListener(self._tabContainer, 'mousedown', function (evt) {
                        var li = evt.target, idx;
                        if (li instanceof HTMLLIElement) {
                            self._sheetTabClicked = true;
                            idx = self._getItemIndex(self._tabContainer, li);
                            self._scrollSheetTabContainer(li);
                            if (idx > -1) {
                                self._sheets.selectedIndex = idx;
                            }
                        }
                    });
                    //todo
                    //contextmenu
                };
                // Initialize the sheet pager part.
                _SheetTabs.prototype._initSheetPage = function () {
                    var self = this;
                    self.hostElement.querySelector('div.wj-sheet-page').addEventListener('click', function (e) {
                        var btn = e.target.toString() === '[object HTMLButtonElement]' ? e.target : e.target.parentElement, index = self._getItemIndex(self._sheetPage, btn), currentSheetTab;
                        if (self._sheets.length === 0) {
                            return;
                        }
                        switch (index) {
                            case 0:
                                if (self._rtl) {
                                    self._sheets.selectLast();
                                }
                                else {
                                    self._sheets.selectFirst();
                                }
                                break;
                            case 1:
                                if (self._rtl) {
                                    self._sheets.selectNext();
                                }
                                else {
                                    self._sheets.selectPrevious();
                                }
                                break;
                            case 2:
                                if (self._rtl) {
                                    self._sheets.selectPrevious();
                                }
                                else {
                                    self._sheets.selectNext();
                                }
                                break;
                            case 3:
                                if (self._rtl) {
                                    self._sheets.selectFirst();
                                }
                                else {
                                    self._sheets.selectLast();
                                }
                                break;
                        }
                    });
                };
                // Get markup for the sheet tabs
                _SheetTabs.prototype._getSheetTabs = function () {
                    var html = '', i;
                    for (i = 0; i < this._sheets.length; i++) {
                        html += this._getSheetElement(this._sheets[i], this._sheets.selectedIndex === i);
                    }
                    return html;
                };
                // Get the markup for a sheet tab.
                _SheetTabs.prototype._getSheetElement = function (sheetItem, isActive) {
                    if (isActive === void 0) { isActive = false; }
                    var result = '<li';
                    if (!sheetItem.visible) {
                        result += ' class="hidden"';
                    }
                    else if (isActive) {
                        result += ' class="active"';
                    }
                    result += '>' + sheetItem.name + '</li>';
                    return result;
                };
                // Update the active state for the sheet tabs.
                _SheetTabs.prototype._updateTabActive = function (pos, active) {
                    if (pos < 0 || pos >= this._tabContainer.children.length) {
                        return;
                    }
                    if (active) {
                        wijmo.addClass(this._tabContainer.children[pos], 'active');
                    }
                    else {
                        wijmo.removeClass(this._tabContainer.children[pos], 'active');
                    }
                };
                // Update the show or hide state for the sheet tabs
                _SheetTabs.prototype._updateTabShown = function (sender, e) {
                    if (e.index < 0 || e.index >= this._tabContainer.children.length) {
                        return;
                    }
                    if (!e.item.visible) {
                        wijmo.addClass(this._tabContainer.children[e.index], 'hidden');
                    }
                    else {
                        wijmo.removeClass(this._tabContainer.children[e.index], 'hidden');
                    }
                    this._adjustSize();
                };
                // Adjust the size of the SheetTabs control.
                _SheetTabs.prototype._adjustSize = function () {
                    //adjust the size
                    var sheetCount = this._tabContainer.childElementCount, index, containerMaxWidth, width = 0, scrollLeft = 0;
                    if (this.hostElement.style.display === 'none') {
                        return;
                    }
                    // Get the scroll left of the tab container, before setting the size of the size of the tab container. (TFS 142788)
                    scrollLeft = this._tabContainer.parentElement.scrollLeft;
                    // Before adjusting the size of the sheet tab, we should reset the size to ''. (TFS #139846)
                    this._tabContainer.parentElement.style.width = '';
                    this._tabContainer.style.width = '';
                    this._sheetPage.parentElement.style.width = '';
                    for (index = 0; index < sheetCount; index++) {
                        width += this._tabContainer.children[index].offsetWidth + 1;
                    }
                    containerMaxWidth = this.hostElement.offsetWidth - this._sheetPage.offsetWidth - this._newSheet.offsetWidth - 2;
                    this._tabContainer.parentElement.style.width = (width > containerMaxWidth ? containerMaxWidth : width) + 'px';
                    this._tabContainer.style.width = width + 'px';
                    this._sheetPage.parentElement.style.width = this._sheetPage.offsetWidth + this._newSheet.offsetWidth + this._tabContainer.parentElement.offsetWidth + 3 + 'px';
                    // Reset the scroll left for the tab container. (TFS 142788)
                    this._tabContainer.parentElement.scrollLeft = scrollLeft;
                };
                // Get the index of the element in its parent container.
                _SheetTabs.prototype._getItemIndex = function (container, item) {
                    var idx = 0;
                    for (; idx < container.children.length; idx++) {
                        if (container.children[idx] === item) {
                            return idx;
                        }
                    }
                    return -1;
                };
                // Update the sheet tab name.
                _SheetTabs.prototype._updateSheetName = function (sender, e) {
                    this._tabContainer.querySelectorAll('li')[e.index].textContent = e.item.name;
                    this._adjustSize();
                };
                // Scroll the sheet tab container to display the invisible or partial visible sheet tab.
                _SheetTabs.prototype._scrollSheetTabContainer = function (currentSheetTab) {
                    var scrollLeft = this._tabContainer.parentElement.scrollLeft, sheetPageSize = this._sheetPage.offsetWidth, newSheetSize = this._newSheet.offsetWidth, containerSize = this._tabContainer.parentElement.offsetWidth, containerOffset;
                    if (this._rtl) {
                        switch (grid_1.FlexGrid['_getRtlMode']()) {
                            case 'rev':
                                containerOffset = -this._tabContainer.offsetLeft;
                                if (containerOffset + currentSheetTab.offsetLeft + currentSheetTab.offsetWidth > containerSize + scrollLeft) {
                                    this._tabContainer.parentElement.scrollLeft += currentSheetTab.offsetWidth;
                                }
                                else if (containerOffset + currentSheetTab.offsetLeft < scrollLeft) {
                                    this._tabContainer.parentElement.scrollLeft -= currentSheetTab.offsetWidth;
                                }
                                break;
                            case 'neg':
                                if (currentSheetTab.offsetLeft < scrollLeft) {
                                    this._tabContainer.parentElement.scrollLeft -= currentSheetTab.offsetWidth;
                                }
                                else if (currentSheetTab.offsetLeft + currentSheetTab.offsetWidth > containerSize + scrollLeft) {
                                    this._tabContainer.parentElement.scrollLeft += currentSheetTab.offsetWidth;
                                }
                                break;
                            default:
                                if (currentSheetTab.offsetLeft - newSheetSize + scrollLeft < 0) {
                                    this._tabContainer.parentElement.scrollLeft += currentSheetTab.offsetWidth;
                                }
                                else if (currentSheetTab.offsetLeft + currentSheetTab.offsetWidth - newSheetSize + scrollLeft > containerSize) {
                                    this._tabContainer.parentElement.scrollLeft -= currentSheetTab.offsetWidth;
                                }
                                break;
                        }
                    }
                    else {
                        if (currentSheetTab.offsetLeft + currentSheetTab.offsetWidth - sheetPageSize > containerSize + scrollLeft) {
                            this._tabContainer.parentElement.scrollLeft += currentSheetTab.offsetWidth;
                        }
                        else if (currentSheetTab.offsetLeft - sheetPageSize < scrollLeft) {
                            this._tabContainer.parentElement.scrollLeft -= currentSheetTab.offsetWidth;
                        }
                    }
                };
                // Adjust the position of each sheet tab for 'rtl' direction.
                _SheetTabs.prototype._adjustSheetsPosition = function () {
                    var sheets = this._tabContainer.querySelectorAll('li'), position = 0, sheet, index;
                    for (index = 0; index < sheets.length; index++) {
                        sheet = sheets[index];
                        sheet.style.cssFloat = 'right';
                        sheet.style.right = position + 'px';
                        position += sheets[index].clientWidth;
                    }
                };
                // Scroll to the active sheet tab.
                _SheetTabs.prototype._scrollToActiveSheet = function (newIndex, oldIndex) {
                    var sheets = this._tabContainer.querySelectorAll('li'), activeSheet, scrollLeft, i;
                    if (this._tabContainer.clientWidth > this._tabContainer.parentElement.clientWidth) {
                        scrollLeft = this._tabContainer.clientWidth - this._tabContainer.parentElement.clientWidth;
                    }
                    else {
                        scrollLeft = 0;
                    }
                    if (sheets.length > 0 && newIndex < sheets.length && oldIndex < sheets.length) {
                        if ((newIndex === 0 && !this._rtl) || (newIndex === sheets.length - 1 && this._rtl)) {
                            if (this._rtl) {
                                switch (grid_1.FlexGrid['_getRtlMode']()) {
                                    case 'rev':
                                        this._tabContainer.parentElement.scrollLeft = 0;
                                        break;
                                    case 'neg':
                                        this._tabContainer.parentElement.scrollLeft = -scrollLeft;
                                        break;
                                    default:
                                        this._tabContainer.parentElement.scrollLeft = scrollLeft;
                                        break;
                                }
                            }
                            else {
                                this._tabContainer.parentElement.scrollLeft = 0;
                            }
                            return;
                        }
                        if ((newIndex === 0 && this._rtl) || (newIndex === sheets.length - 1 && !this._rtl)) {
                            if (this._rtl) {
                                switch (grid_1.FlexGrid['_getRtlMode']()) {
                                    case 'rev':
                                        this._tabContainer.parentElement.scrollLeft = scrollLeft;
                                        break;
                                    case 'neg':
                                        this._tabContainer.parentElement.scrollLeft = 0;
                                        break;
                                    default:
                                        this._tabContainer.parentElement.scrollLeft = 0;
                                        break;
                                }
                            }
                            else {
                                this._tabContainer.parentElement.scrollLeft = scrollLeft;
                            }
                            return;
                        }
                        if (newIndex >= oldIndex) {
                            for (i = oldIndex + 1; i <= newIndex; i++) {
                                activeSheet = sheets[i];
                                this._scrollSheetTabContainer(activeSheet);
                            }
                        }
                        else {
                            for (i = oldIndex - 1; i >= newIndex; i--) {
                                activeSheet = sheets[i];
                                this._scrollSheetTabContainer(activeSheet);
                            }
                        }
                    }
                };
                _SheetTabs.controlTemplate = '<div wj-part="sheet-container" class="wj-sheet" style="height:100%;position:relative">' +
                    '<div wj-part="sheet-page" class="wj-btn-group wj-sheet-page">' +
                    '<button type="button" class="wj-btn wj-btn-default">' +
                    '<span class="wj-sheet-icon wj-glyph-step-backward"></span>' +
                    '</button>' +
                    '<button type="button" class="wj-btn wj-btn-default">' +
                    '<span class="wj-sheet-icon wj-glyph-left"></span>' +
                    '</button>' +
                    '<button type="button" class="wj-btn wj-btn-default">' +
                    '<span class="wj-sheet-icon wj-glyph-right"></span>' +
                    '</button>' +
                    '<button type="button" class="wj-btn wj-btn-default">' +
                    '<span class="wj-sheet-icon wj-glyph-step-forward"></span>' +
                    '</button>' +
                    '</div>' +
                    '<div class="wj-sheet-tab" style="height:100%;overflow:hidden">' +
                    '<ul wj-part="container"></ul>' +
                    '</div>' +
                    '<div wj-part="new-sheet" class="wj-new-sheet"><span class="wj-sheet-icon wj-glyph-file"></span></div>' +
                    '</div>';
                return _SheetTabs;
            }(wijmo.Control));
            sheet_1._SheetTabs = _SheetTabs;
            /*
             * Defines the class defining @see:FlexSheet column sorting criterion.
             */
            var _UnboundSortDescription = (function () {
                /*
                 * Initializes a new instance of the @see:UnboundSortDescription class.
                 *
                 * @param column The column to sort the rows by.
                 * @param ascending The sort order.
                 */
                function _UnboundSortDescription(column, ascending) {
                    this._column = column;
                    this._ascending = ascending;
                }
                Object.defineProperty(_UnboundSortDescription.prototype, "column", {
                    /*
                     * Gets the column to sort the rows by.
                     */
                    get: function () {
                        return this._column;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(_UnboundSortDescription.prototype, "ascending", {
                    /*
                     * Gets the sort order.
                     */
                    get: function () {
                        return this._ascending;
                    },
                    enumerable: true,
                    configurable: true
                });
                return _UnboundSortDescription;
            }());
            sheet_1._UnboundSortDescription = _UnboundSortDescription;
        })(sheet = grid_1.sheet || (grid_1.sheet = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=Sheet.js.map
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var sheet;
        (function (sheet) {
            'use strict';
            /**
             * Maintains sorting of the selected @see:Sheet of the @see:FlexSheet.
             */
            var SortManager = (function () {
                /**
                 * Initializes a new instance of the @see:SortManager class.
                 *
                 * @param owner The @see:FlexSheet control that owns this <b>SortManager</b>.
                 */
                function SortManager(owner) {
                    this._owner = owner;
                    this._sortDescriptions = new wijmo.collections.CollectionView();
                    this._committedList = [new ColumnSortDescription(-1, true)];
                    this._sortDescriptions.newItemCreator = function () {
                        return new ColumnSortDescription(-1, true);
                    };
                    this._refresh();
                }
                Object.defineProperty(SortManager.prototype, "sortDescriptions", {
                    /**
                     * Gets or sets the collection of the sort descriptions represented by the  @see:ColumnSortDescription objects.
                     */
                    get: function () {
                        return this._sortDescriptions;
                    },
                    set: function (value) {
                        this._sortDescriptions = value;
                        this.commitSort(true);
                        this._refresh();
                    },
                    enumerable: true,
                    configurable: true
                });
                /**
                 * Adds a blank sorting level to the sort descriptions.
                 *
                 * @param columnIndex The index of the column in the FlexSheet control.
                 * @param ascending The sort order for the sort level.
                 */
                SortManager.prototype.addSortLevel = function (columnIndex, ascending) {
                    if (ascending === void 0) { ascending = true; }
                    var item = this._sortDescriptions.addNew();
                    if (columnIndex != null && !isNaN(columnIndex) && wijmo.isInt(columnIndex)) {
                        item.columnIndex = columnIndex;
                    }
                    item.ascending = ascending;
                    this._sortDescriptions.commitNew();
                };
                /**
                 * Removes the current sorting level from the sort descriptions.
                 *
                 * @param columnIndex The index of the column in the FlexSheet control.
                 */
                SortManager.prototype.deleteSortLevel = function (columnIndex) {
                    var item;
                    if (columnIndex != null) {
                        item = this._getSortItem(columnIndex);
                    }
                    else {
                        item = this._sortDescriptions.currentItem;
                    }
                    if (item) {
                        this._sortDescriptions.remove(item);
                    }
                };
                /**
                 * Adds a copy of the current sorting level to the sort descriptions.
                 */
                SortManager.prototype.copySortLevel = function () {
                    var item = this._sortDescriptions.currentItem;
                    if (item) {
                        var newItem = this._sortDescriptions.addNew();
                        newItem.columnIndex = parseInt(item.columnIndex);
                        newItem.ascending = item.ascending;
                        this._sortDescriptions.commitNew();
                    }
                };
                /**
                 * Updates the current sort level.
                 *
                 * @param columnIndex The column index for the sort level.
                 * @param ascending The sort order for the sort level.
                 */
                SortManager.prototype.editSortLevel = function (columnIndex, ascending) {
                    if (columnIndex != null) {
                        this._sortDescriptions.currentItem.columnIndex = columnIndex;
                    }
                    if (ascending != null) {
                        this._sortDescriptions.currentItem.ascending = ascending;
                    }
                };
                /**
                 * Moves the current sorting level to a new position.
                 *
                 * @param offset The offset to move the current level by.
                 */
                SortManager.prototype.moveSortLevel = function (offset) {
                    var item = this._sortDescriptions.currentItem;
                    if (item) {
                        var arr = this._sortDescriptions.sourceCollection, index = arr.indexOf(item), newIndex = index + offset;
                        if (index > -1 && newIndex > -1) {
                            arr.splice(index, 1);
                            arr.splice(newIndex, 0, item);
                            this._sortDescriptions.refresh();
                            this._sortDescriptions.moveCurrentTo(item);
                        }
                    }
                };
                /**
                 * Check whether the sort item of specific column exists or not
                 *
                 * @param columnIndex The index of the column in the FlexSheet control.
                 */
                SortManager.prototype.checkSortItemExists = function (columnIndex) {
                    var i = 0, sortItemCnt = this._sortDescriptions.itemCount, sortItem;
                    for (; i < sortItemCnt; i++) {
                        sortItem = this._sortDescriptions.items[i];
                        if (+sortItem.columnIndex === columnIndex) {
                            return i;
                        }
                    }
                    return -1;
                };
                /**
                 * Commits the current sort descriptions to the FlexSheet control.
                 *
                 * @param undoable The boolean value indicating whether the commit sort action is undoable.
                 */
                SortManager.prototype.commitSort = function (undoable) {
                    var _this = this;
                    if (undoable === void 0) { undoable = true; }
                    var sd, newSortDesc, bindSortDesc, dataBindSortDesc, i, unSortDesc, sortAction, unboundRows, isCVItemsSource = this._owner.itemsSource && this._owner.itemsSource instanceof wijmo.collections.CollectionView;
                    if (!this._owner.selectedSheet) {
                        return;
                    }
                    unSortDesc = this._owner.selectedSheet._unboundSortDesc;
                    if (undoable) {
                        sortAction = new sheet._SortColumnAction(this._owner);
                    }
                    if (this._sortDescriptions.itemCount > 0) {
                        this._committedList = this._sortDescriptions.items.slice();
                    }
                    else {
                        this._committedList = [new ColumnSortDescription(-1, true)];
                    }
                    if (this._owner.collectionView) {
                        // Try to get the unbound row in the bound sheet.
                        unboundRows = this._scanUnboundRows();
                        // Update sorting for the bind booksheet
                        this._owner.collectionView.beginUpdate();
                        this._owner.selectedSheet.grid.collectionView.beginUpdate();
                        bindSortDesc = this._owner.collectionView.sortDescriptions;
                        bindSortDesc.clear();
                        // Synch the sorts for the grid of current sheet.
                        if (isCVItemsSource === false) {
                            dataBindSortDesc = this._owner.selectedSheet.grid.collectionView.sortDescriptions;
                            dataBindSortDesc.clear();
                        }
                        for (i = 0; i < this._sortDescriptions.itemCount; i++) {
                            sd = this._sortDescriptions.items[i];
                            if (sd.columnIndex > -1) {
                                newSortDesc = new wijmo.collections.SortDescription(this._owner.columns[sd.columnIndex].binding, sd.ascending);
                                bindSortDesc.push(newSortDesc);
                                // Synch the sorts for the grid of current sheet.
                                if (isCVItemsSource === false) {
                                    dataBindSortDesc.push(newSortDesc);
                                }
                            }
                        }
                        this._owner.collectionView.endUpdate();
                        this._owner.selectedSheet.grid.collectionView.endUpdate();
                        // Re-insert the unbound row into the sheet.
                        if (unboundRows) {
                            Object.keys(unboundRows).forEach(function (key) {
                                _this._owner.rows.splice(+key, 0, unboundRows[key]);
                            });
                        }
                    }
                    else {
                        // Update sorting for the unbound booksheet.
                        unSortDesc.clear();
                        for (i = 0; i < this._sortDescriptions.itemCount; i++) {
                            sd = this._sortDescriptions.items[i];
                            if (sd.columnIndex > -1) {
                                unSortDesc.push(new sheet._UnboundSortDescription(this._owner.columns[sd.columnIndex], sd.ascending));
                            }
                        }
                    }
                    this._owner._filter.apply();
                    if (undoable) {
                        sortAction.saveNewState();
                        this._owner.undoStack._addAction(sortAction);
                    }
                };
                /**
                 * Cancel the current sort descriptions to the FlexSheet control.
                 */
                SortManager.prototype.cancelSort = function () {
                    this._sortDescriptions.sourceCollection = this._committedList.slice();
                    this._refresh();
                };
                // Updates the <b>sorts</b> collection based on the current @see:Sheet sort conditions.
                SortManager.prototype._refresh = function () {
                    var sortList = [], i, sd;
                    if (!this._owner.selectedSheet) {
                        return;
                    }
                    if (this._owner.collectionView && this._owner.collectionView.sortDescriptions.length > 0) {
                        for (i = 0; i < this._owner.collectionView.sortDescriptions.length; i++) {
                            sd = this._owner.collectionView.sortDescriptions[i];
                            sortList.push(new ColumnSortDescription(this._getColumnIndex(sd.property), sd.ascending));
                        }
                    }
                    else if (this._owner.selectedSheet && this._owner.selectedSheet._unboundSortDesc.length > 0) {
                        for (i = 0; i < this._owner.selectedSheet._unboundSortDesc.length; i++) {
                            sd = this._owner.selectedSheet._unboundSortDesc[i];
                            sortList.push(new ColumnSortDescription(sd.column.index, sd.ascending));
                        }
                    }
                    else {
                        sortList.push(new ColumnSortDescription(-1, true));
                    }
                    this._sortDescriptions.sourceCollection = sortList;
                };
                // Get the index of the column by the binding property.
                SortManager.prototype._getColumnIndex = function (property) {
                    var i = 0, colCnt = this._owner.columns.length;
                    for (; i < colCnt; i++) {
                        if (this._owner.columns[i].binding === property) {
                            return i;
                        }
                    }
                    return -1;
                };
                // Get the sort item via the column index
                SortManager.prototype._getSortItem = function (columnIndex) {
                    var index = this.checkSortItemExists(columnIndex);
                    if (index > -1) {
                        return this._sortDescriptions.items[index];
                    }
                    return undefined;
                };
                // Scan the unbound row of the bound sheet.
                SortManager.prototype._scanUnboundRows = function () {
                    var rowIndex, processingRow, unboundRows;
                    for (rowIndex = 0; rowIndex < this._owner.rows.length; rowIndex++) {
                        processingRow = this._owner.rows[rowIndex];
                        if (!processingRow.dataItem) {
                            if (!(processingRow instanceof sheet.HeaderRow)) {
                                if (!unboundRows) {
                                    unboundRows = {};
                                }
                                unboundRows[rowIndex] = processingRow;
                            }
                        }
                    }
                    return unboundRows;
                };
                return SortManager;
            }());
            sheet.SortManager = SortManager;
            /**
             * Describes a @see:FlexSheet column sorting criterion.
             */
            var ColumnSortDescription = (function () {
                /**
                 * Initializes a new instance of the @see:ColumnSortDescription class.
                 *
                 * @param columnIndex Indicates which column to sort the rows by.
                 * @param ascending The sort order.
                 */
                function ColumnSortDescription(columnIndex, ascending) {
                    this._columnIndex = columnIndex;
                    this._ascending = ascending;
                }
                Object.defineProperty(ColumnSortDescription.prototype, "columnIndex", {
                    /**
                     * Gets or sets the column index.
                     */
                    get: function () {
                        return this._columnIndex;
                    },
                    set: function (value) {
                        this._columnIndex = value;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(ColumnSortDescription.prototype, "ascending", {
                    /**
                     * Gets or sets the ascending.
                     */
                    get: function () {
                        return this._ascending;
                    },
                    set: function (value) {
                        this._ascending = value;
                    },
                    enumerable: true,
                    configurable: true
                });
                return ColumnSortDescription;
            }());
            sheet.ColumnSortDescription = ColumnSortDescription;
        })(sheet = grid.sheet || (grid.sheet = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=SortManager.js.map
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var sheet;
        (function (sheet) {
            'use strict';
            /**
             * Controls undo and redo operations in the @see:FlexSheet.
             */
            var UndoStack = (function () {
                /**
                 * Initializes a new instance of the @see:UndoStack class.
                 *
                 * @param owner The @see:FlexSheet control that the @see:UndoStack works for.
                 */
                function UndoStack(owner) {
                    this.MAX_STACK_SIZE = 500;
                    this._stack = [];
                    this._pointer = -1;
                    this._resizingTriggered = false;
                    /**
                     * Occurs after the undo stack has changed.
                     */
                    this.undoStackChanged = new wijmo.Event();
                    var self = this;
                    self._owner = owner;
                    // Handles the cell edit action for editing cell
                    self._owner.prepareCellForEdit.addHandler(self._initCellEditAction, self);
                    self._owner.cellEditEnded.addHandler(function () {
                        // For edit cell content.
                        if (self._pendingAction instanceof sheet._EditAction && !self._pendingAction.isPaste) {
                            self._afterProcessCellEditAction(self);
                        }
                    }, self);
                    // Handles the cell edit action for copy\paste operation
                    self._owner.pasting.addHandler(self._initCellEditActionForPasting, self);
                    self._owner.pastingCell.addHandler(function (sender, e) {
                        if (self._pendingAction instanceof sheet._EditAction) {
                            self._pendingAction.updateForPasting(e.range);
                        }
                    }, self);
                    self._owner.pasted.addHandler(function () {
                        // For paste content to the cell.
                        if (self._pendingAction instanceof sheet._EditAction && self._pendingAction.isPaste) {
                            self._afterProcessCellEditAction(self);
                        }
                    }, self);
                    // Handles the resize column action
                    self._owner.resizingColumn.addHandler(function (sender, e) {
                        if (!self._resizingTriggered) {
                            self._pendingAction = new sheet._ColumnResizeAction(self._owner, e.panel, e.col);
                            self._resizingTriggered = true;
                        }
                    }, self);
                    self._owner.resizedColumn.addHandler(function (sender, e) {
                        if (self._pendingAction instanceof sheet._ColumnResizeAction && self._pendingAction.saveNewState()) {
                            self._addAction(self._pendingAction);
                        }
                        self._pendingAction = null;
                        self._resizingTriggered = false;
                    }, self);
                    // Handles the resize row action
                    self._owner.resizingRow.addHandler(function (sender, e) {
                        if (!self._resizingTriggered) {
                            self._pendingAction = new sheet._RowResizeAction(self._owner, e.panel, e.row);
                            self._resizingTriggered = true;
                        }
                    }, self);
                    self._owner.resizedRow.addHandler(function (sender, e) {
                        if (self._pendingAction instanceof sheet._RowResizeAction && self._pendingAction.saveNewState()) {
                            self._addAction(self._pendingAction);
                        }
                        self._pendingAction = null;
                        self._resizingTriggered = false;
                    }, self);
                    // Handle the changing rows\columns position action.
                    self._owner.draggingRowColumn.addHandler(function (sender, e) {
                        if (e.isShiftKey) {
                            if (e.isDraggingRows) {
                                self._pendingAction = new sheet._RowsChangedAction(self._owner);
                            }
                            else {
                                self._pendingAction = new sheet._ColumnsChangedAction(self._owner);
                            }
                        }
                    }, self);
                    self._owner.droppingRowColumn.addHandler(function () {
                        if (self._pendingAction && self._pendingAction.saveNewState()) {
                            self._addAction(self._pendingAction);
                        }
                        self._pendingAction = null;
                    }, self);
                }
                Object.defineProperty(UndoStack.prototype, "canUndo", {
                    /**
                     * Checks whether the undo action can be performed.
                     */
                    get: function () {
                        return this._pointer > -1 && this._pointer < this._stack.length;
                    },
                    enumerable: true,
                    configurable: true
                });
                Object.defineProperty(UndoStack.prototype, "canRedo", {
                    /**
                     * Checks whether the redo action can be performed.
                     */
                    get: function () {
                        return this._pointer + 1 > -1 && this._pointer + 1 < this._stack.length;
                    },
                    enumerable: true,
                    configurable: true
                });
                /**
                 * Raises the <b>undoStackChanged</b> event.
                 */
                UndoStack.prototype.onUndoStackChanged = function () {
                    this.undoStackChanged.raise(this);
                };
                /**
                 * Undo the latest action.
                 */
                UndoStack.prototype.undo = function () {
                    var action;
                    if (this.canUndo) {
                        action = this._stack[this._pointer];
                        this._beforeUndoRedo(action);
                        action.undo();
                        this._pointer--;
                        this.onUndoStackChanged();
                    }
                };
                /**
                 * Redo the latest undone action.
                 */
                UndoStack.prototype.redo = function () {
                    var action;
                    if (this.canRedo) {
                        this._pointer++;
                        action = this._stack[this._pointer];
                        this._beforeUndoRedo(action);
                        action.redo();
                        this.onUndoStackChanged();
                    }
                };
                /*
                 * Add the undo action into the undo stack.
                 *
                 * @param action The @see:_UndoAction undo/redo processing actions.
                 */
                UndoStack.prototype._addAction = function (action) {
                    // trim stack
                    if (this._stack.length > 0 && this._stack.length > this._pointer + 1) {
                        this._stack.splice(this._pointer + 1, this._stack.length - this._pointer - 1);
                    }
                    if (this._stack.length >= this.MAX_STACK_SIZE) {
                        this._stack.splice(0, this._stack.length - this.MAX_STACK_SIZE + 1);
                    }
                    // update pointer and add action to stack
                    this._pointer = this._stack.length;
                    this._stack.push(action);
                    this.onUndoStackChanged();
                };
                /**
                 * Clears the undo stack.
                 */
                UndoStack.prototype.clear = function () {
                    this._stack.length = 0;
                };
                // initialize the cell edit action.
                UndoStack.prototype._initCellEditAction = function (sender, args) {
                    this._pendingAction = new sheet._EditAction(this._owner, args.range);
                };
                // initialize the cell edit action for pasting action.
                UndoStack.prototype._initCellEditActionForPasting = function () {
                    this._pendingAction = new sheet._EditAction(this._owner);
                    this._pendingAction.markIsPaste();
                };
                // after processing the cell edit action.
                UndoStack.prototype._afterProcessCellEditAction = function (self) {
                    if (self._pendingAction instanceof sheet._EditAction && self._pendingAction.saveNewState()) {
                        self._addAction(this._pendingAction);
                    }
                    self._pendingAction = null;
                };
                // Called before an action is undone or redone.
                UndoStack.prototype._beforeUndoRedo = function (action) {
                    this._owner.selectedSheetIndex = action.sheetIndex;
                };
                return UndoStack;
            }());
            sheet.UndoStack = UndoStack;
        })(sheet = grid.sheet || (grid.sheet = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=UndoStack.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var sheet;
        (function (sheet) {
            'use strict';
            /*
             * Defines a value filter for a column on a @see:FlexSheet control.
             *
             * Value filters contain an explicit list of values that should be
             * displayed by the sheet.
             */
            var _FlexSheetValueFilter = (function (_super) {
                __extends(_FlexSheetValueFilter, _super);
                function _FlexSheetValueFilter() {
                    _super.apply(this, arguments);
                }
                /*
                 * Gets a value that indicates whether a value passes the filter.
                 *
                 * @param value The value to test.
                 */
                _FlexSheetValueFilter.prototype.apply = function (value) {
                    var flexSheet = this.column.grid;
                    if (!(flexSheet instanceof sheet.FlexSheet)) {
                        return false;
                    }
                    // values? accept everything
                    if (!this.showValues || !Object.keys(this.showValues).length) {
                        return true;
                    }
                    value = flexSheet.getCellValue(value, this.column.index, true);
                    // apply conditions
                    return this.showValues[value] != undefined;
                };
                return _FlexSheetValueFilter;
            }(wijmo.grid.filter.ValueFilter));
            sheet._FlexSheetValueFilter = _FlexSheetValueFilter;
        })(sheet = grid.sheet || (grid.sheet = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_FlexSheetValueFilter.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var sheet;
        (function (sheet) {
            'use strict';
            /*
             * The editor used to inspect and modify @see:FlexSheetValueFilter objects.
             *
             * This class is used by the @see:FlexSheetFilter class; you
             * rarely use it directly.
             */
            var _FlexSheetValueFilterEditor = (function (_super) {
                __extends(_FlexSheetValueFilterEditor, _super);
                function _FlexSheetValueFilterEditor() {
                    _super.apply(this, arguments);
                }
                /*
                 * Updates editor with current filter settings.
                 */
                _FlexSheetValueFilterEditor.prototype.updateEditor = function () {
                    var col = this.filter.column, flexSheet = col.grid, colIndex = col.index, values = [], keys = {}, row, mergedRange, value, sv, currentFilterResult, otherFilterResult, text;
                    // get list of unique values
                    if (this.filter.uniqueValues) {
                        _super.prototype.updateEditor.call(this);
                        return;
                    }
                    // format and add unique values to the 'values' array
                    for (var i = 0; i < flexSheet.rows.length; i++) {
                        // Get the result of current filter for current row.
                        currentFilterResult = this.filter.apply(i);
                        // Get the result of other filters for current row.
                        sv = this.filter.showValues;
                        this.filter.showValues = null;
                        otherFilterResult = flexSheet._filter['_filter'](i);
                        this.filter.showValues = sv;
                        mergedRange = flexSheet.getMergedRange(flexSheet.cells, i, colIndex);
                        if (mergedRange && (i !== mergedRange.topRow || colIndex !== mergedRange.leftCol)) {
                            continue;
                        }
                        row = flexSheet.rows[i];
                        if (row instanceof sheet.HeaderRow || (!row.isVisible && (currentFilterResult || !otherFilterResult))) {
                            continue;
                        }
                        value = flexSheet.getCellValue(i, colIndex);
                        text = flexSheet.getCellValue(i, colIndex, true);
                        if (!keys[text]) {
                            keys[text] = true;
                            values.push({ value: value, text: text });
                        }
                    }
                    // check the items that are currently selected
                    var showValues = this.filter.showValues;
                    if (!showValues || Object.keys(showValues).length == 0) {
                        for (var i = 0; i < values.length; i++) {
                            values[i].show = true;
                        }
                    }
                    else {
                        for (var key in showValues) {
                            for (var i = 0; i < values.length; i++) {
                                if (values[i].text == key) {
                                    values[i].show = true;
                                    break;
                                }
                            }
                        }
                    }
                    // honor isContentHtml property
                    this['_lbValues'].isContentHtml = col.isContentHtml;
                    // load filter and apply immediately
                    this['_cmbFilter'].text = this.filter.filterText;
                    this['_filterText'] = this['_cmbFilter'].text.toLowerCase();
                    // show the values
                    this['_view'].pageSize = this.filter.maxValues;
                    this['_view'].sourceCollection = values;
                    this['_view'].moveCurrentToPosition(-1);
                };
                return _FlexSheetValueFilterEditor;
            }(wijmo.grid.filter.ValueFilterEditor));
            sheet._FlexSheetValueFilterEditor = _FlexSheetValueFilterEditor;
        })(sheet = grid.sheet || (grid.sheet = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_FlexSheetValueFilterEditor.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var sheet;
        (function (sheet) {
            'use strict';
            /*
             * Defines a condition filter for a column on a @see:FlexSheet control.
             *
             * Condition filters contain two conditions that may be combined
             * using an 'and' or an 'or' operator.
             *
             * This class is used by the @see:FlexSheetFilter class; you will
             * rarely use it directly.
             */
            var _FlexSheetConditionFilter = (function (_super) {
                __extends(_FlexSheetConditionFilter, _super);
                function _FlexSheetConditionFilter() {
                    _super.apply(this, arguments);
                }
                /*
                 * Returns a value indicating whether a value passes this filter.
                 *
                 * @param value The value to test.
                 */
                _FlexSheetConditionFilter.prototype.apply = function (value) {
                    var col = this.column, flexSheet = col.grid, c1 = this.condition1, c2 = this.condition2, compareVal, compareVal1, compareVal2;
                    if (!(flexSheet instanceof sheet.FlexSheet)) {
                        return false;
                    }
                    // no binding or not active? accept everything
                    if (!this.isActive) {
                        return true;
                    }
                    // retrieve the value
                    compareVal = flexSheet.getCellValue(value, col.index);
                    compareVal1 = compareVal2 = compareVal;
                    if (col.dataMap) {
                        compareVal = col.dataMap.getDisplayValue(compareVal);
                        compareVal1 = compareVal2 = compareVal;
                    }
                    else if (wijmo.isDate(compareVal)) {
                        if (wijmo.isString(c1.value) || wijmo.isString(c2.value)) {
                            compareVal = flexSheet.getCellValue(value, col.index, true);
                            compareVal1 = compareVal2 = compareVal;
                        }
                    }
                    else if (wijmo.isNumber(compareVal)) {
                        compareVal = wijmo.Globalize.parseFloat(flexSheet.getCellValue(value, col.index, true));
                        compareVal1 = compareVal2 = compareVal;
                        if (compareVal === 0 && !col.dataType) {
                            if (c1.isActive && c1.value === '') {
                                compareVal1 = null;
                            }
                            if (c2.isActive && c2.value === '') {
                                compareVal2 = null;
                            }
                        }
                    }
                    // apply conditions
                    var rv1 = c1.apply(compareVal1), rv2 = c2.apply(compareVal2);
                    // combine results
                    if (c1.isActive && c2.isActive) {
                        return this.and ? rv1 && rv2 : rv1 || rv2;
                    }
                    else {
                        return c1.isActive ? rv1 : c2.isActive ? rv2 : true;
                    }
                };
                return _FlexSheetConditionFilter;
            }(wijmo.grid.filter.ConditionFilter));
            sheet._FlexSheetConditionFilter = _FlexSheetConditionFilter;
        })(sheet = grid.sheet || (grid.sheet = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_FlexSheetConditionFilter.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var sheet;
        (function (sheet) {
            'use strict';
            /*
             * Defines a filter for a column on a @see:FlexSheet control.
             *
             * The @see:FlexSheetColumnFilter contains a @see:FlexSheetConditionFilter and a
             * @see:FlexSheetValueFilter; only one of them may be active at a time.
             *
             * This class is used by the @see:FlexSheetFilter class; you
             * rarely use it directly.
             */
            var _FlexSheetColumnFilter = (function (_super) {
                __extends(_FlexSheetColumnFilter, _super);
                /*
                 * Initializes a new instance of the @see:FlexSheetColumnFilter class.
                 *
                 * @param owner The @see:FlexSheetFilter that owns this column filter.
                 * @param column The @see:Column to filter.
                 */
                function _FlexSheetColumnFilter(owner, column) {
                    _super.call(this, owner, column);
                    this['_valueFilter'] = new sheet._FlexSheetValueFilter(column);
                    this['_conditionFilter'] = new sheet._FlexSheetConditionFilter(column);
                }
                return _FlexSheetColumnFilter;
            }(wijmo.grid.filter.ColumnFilter));
            sheet._FlexSheetColumnFilter = _FlexSheetColumnFilter;
        })(sheet = grid.sheet || (grid.sheet = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_FlexSheetColumnFilter.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var sheet;
        (function (sheet) {
            'use strict';
            /*
             * The editor used to inspect and modify column filters.
             *
             * This class is used by the @see:FlexSheetFilter class; you
             * rarely use it directly.
             */
            var _FlexSheetColumnFilterEditor = (function (_super) {
                __extends(_FlexSheetColumnFilterEditor, _super);
                /*
                 * Initializes a new instance of the @see:FlexSheetColumnFilterEditor class.
                 *
                 * @param element The DOM element that hosts the control, or a selector
                 * for the host element (e.g. '#theCtrl').
                 * @param filter The @see:FlexSheetColumnFilter to edit.
                 * @param sortButtons Whether to show sort buttons in the editor.
                 */
                function _FlexSheetColumnFilterEditor(element, filter, sortButtons) {
                    if (sortButtons === void 0) { sortButtons = true; }
                    _super.call(this, element, filter, sortButtons);
                    var self = this, btnAsc, btnDsc;
                    if (sortButtons) {
                        this['_divSort'].style.display = '';
                    }
                    btnAsc = this.cloneElement(this['_btnAsc']);
                    btnDsc = this.cloneElement(this['_btnDsc']);
                    this['_btnAsc'].parentNode.replaceChild(btnAsc, this['_btnAsc']);
                    this['_btnDsc'].parentNode.replaceChild(btnDsc, this['_btnDsc']);
                    btnAsc.addEventListener('click', function (e) {
                        self._sortBtnClick(e, true);
                    });
                    btnDsc.addEventListener('click', function (e) {
                        self._sortBtnClick(e, false);
                    });
                }
                // shows the value or filter editor
                _FlexSheetColumnFilterEditor.prototype._showFilter = function (filterType) {
                    // create editor if we have to
                    if (filterType == wijmo.grid.filter.FilterType.Value && this['_edtVal'] == null) {
                        this['_edtVal'] = new sheet._FlexSheetValueFilterEditor(this['_divEdtVal'], this.filter.valueFilter);
                    }
                    _super.prototype._showFilter.call(this, filterType);
                };
                // sort button click event handler
                _FlexSheetColumnFilterEditor.prototype._sortBtnClick = function (e, asceding) {
                    var column = this.filter.column, sortManager = column.grid.sortManager, sortIndex, offset, sortItem;
                    e.preventDefault();
                    e.stopPropagation();
                    sortIndex = sortManager.checkSortItemExists(column.index);
                    if (sortIndex > -1) {
                        // If the sort item for current column doesn't exist, we add new sort item for current column
                        sortManager.sortDescriptions.moveCurrentToPosition(sortIndex);
                        sortItem = sortManager.sortDescriptions.currentItem;
                        sortItem.ascending = asceding;
                        offset = -sortIndex;
                    }
                    else {
                        sortManager.addSortLevel(column.index, asceding);
                        offset = -(sortManager.sortDescriptions.items.length - 1);
                    }
                    // Move sort item for current column to first level.
                    sortManager.moveSortLevel(offset);
                    sortManager.commitSort();
                    // show current filter state
                    this.updateEditor();
                    // raise event so caller can close the editor and apply the new filter
                    this.onButtonClicked();
                };
                // Clone dom element and its child node
                _FlexSheetColumnFilterEditor.prototype.cloneElement = function (element) {
                    var cloneEle = element.cloneNode();
                    while (element.firstChild) {
                        cloneEle.appendChild(element.lastChild);
                    }
                    return cloneEle;
                };
                return _FlexSheetColumnFilterEditor;
            }(wijmo.grid.filter.ColumnFilterEditor));
            sheet._FlexSheetColumnFilterEditor = _FlexSheetColumnFilterEditor;
        })(sheet = grid.sheet || (grid.sheet = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_FlexSheetColumnFilterEditor.js.map
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var wijmo;
(function (wijmo) {
    var grid;
    (function (grid) {
        var sheet;
        (function (sheet) {
            'use strict';
            /*
             * Implements an Excel-style filter for @see:FlexSheet controls.
             *
             * To enable filtering on a @see:FlexSheet control, create an instance
             * of the @see:FlexSheetFilter and pass the grid as a parameter to the
             * constructor.
             */
            var _FlexSheetFilter = (function (_super) {
                __extends(_FlexSheetFilter, _super);
                function _FlexSheetFilter() {
                    _super.apply(this, arguments);
                }
                Object.defineProperty(_FlexSheetFilter.prototype, "filterDefinition", {
                    /*
                     * Gets or sets the current filter definition as a JSON string.
                     */
                    get: function () {
                        var def = {
                            defaultFilterType: this.defaultFilterType,
                            filters: []
                        };
                        for (var i = 0; i < this['_filters'].length; i++) {
                            var cf = this['_filters'][i];
                            if (cf && cf.column) {
                                if (cf.conditionFilter.isActive) {
                                    var cfc = cf.conditionFilter;
                                    def.filters.push({
                                        columnIndex: cf.column.index,
                                        type: 'condition',
                                        condition1: { operator: cfc.condition1.operator, value: cfc.condition1.value },
                                        and: cfc.and,
                                        condition2: { operator: cfc.condition2.operator, value: cfc.condition2.value }
                                    });
                                }
                                else if (cf.valueFilter.isActive) {
                                    var cfv = cf.valueFilter;
                                    def.filters.push({
                                        columnIndex: cf.column.index,
                                        type: 'value',
                                        filterText: cfv.filterText,
                                        showValues: cfv.showValues
                                    });
                                }
                            }
                        }
                        return JSON.stringify(def);
                    },
                    set: function (value) {
                        var def = JSON.parse(wijmo.asString(value));
                        this.clear();
                        this.defaultFilterType = def.defaultFilterType;
                        for (var i = 0; i < def.filters.length; i++) {
                            var cfs = def.filters[i], col = this.grid.columns[cfs.columnIndex], cf = this.getColumnFilter(col, true);
                            if (cf) {
                                switch (cfs.type) {
                                    case 'condition':
                                        var cfc = cf.conditionFilter;
                                        cfc.condition1.value = col.dataType == wijmo.DataType.Date // handle times/times: TFS 125144, 143453
                                            ? wijmo.changeType(cfs.condition1.value, col.dataType, null)
                                            : cfs.condition1.value;
                                        cfc.condition1.operator = cfs.condition1.operator;
                                        cfc.and = cfs.and;
                                        cfc.condition2.value = col.dataType == wijmo.DataType.Date
                                            ? wijmo.changeType(cfs.condition2.value, col.dataType, null)
                                            : cfs.condition2.value;
                                        cfc.condition2.operator = cfs.condition2.operator;
                                        break;
                                    case 'value':
                                        var cfv = cf.valueFilter;
                                        cfv.filterText = cfs.filterText;
                                        cfv.showValues = cfs.showValues;
                                        break;
                                }
                            }
                        }
                        this.apply();
                    },
                    enumerable: true,
                    configurable: true
                });
                /*
                 * Applies the current column filters to the sheet.
                 */
                _FlexSheetFilter.prototype.apply = function () {
                    var self = this;
                    self.grid.deferUpdate(function () {
                        var row;
                        for (var i = 0; i < self.grid.rows.length; i++) {
                            row = self.grid.rows[i];
                            if (row instanceof sheet.HeaderRow) {
                                continue;
                            }
                            row.visible = self['_filter'](i);
                        }
                    });
                };
                /*
                 * Shows the filter editor for the given grid column.
                 *
                 * @param col The @see:Column that contains the filter to edit.
                 * @param ht A @see:HitTestInfo object containing the range of the cell that triggered the filter display.
                 */
                _FlexSheetFilter.prototype.editColumnFilter = function (col, ht) {
                    var _this = this;
                    // remove current editor
                    this.closeEditor();
                    // get column by name or by reference
                    col = wijmo.isString(col)
                        ? this.grid.columns.getColumn(col)
                        : wijmo.asType(col, grid.Column, false);
                    // raise filterChanging event
                    var e = new grid.CellRangeEventArgs(this.grid.cells, new grid.CellRange(-1, col.index));
                    this.onFilterChanging(e);
                    if (e.cancel) {
                        return;
                    }
                    e.cancel = true; // assume the changes will be canceled
                    // get the filter and the editor
                    var div = document.createElement('div'), flt = this.getColumnFilter(col), edt = new sheet._FlexSheetColumnFilterEditor(div, flt, this.showSortButtons);
                    wijmo.addClass(div, 'wj-dropdown-panel');
                    // handle RTL
                    if (this.grid._rtl) {
                        div.dir = 'rtl';
                    }
                    // apply filter when it changes
                    edt.filterChanged.addHandler(function () {
                        e.cancel = false; // the changes were not canceled
                        setTimeout(function () {
                            if (!e.cancel) {
                                _this.apply();
                            }
                        });
                    });
                    // close editor when editor button is clicked
                    edt.buttonClicked.addHandler(function () {
                        _this.closeEditor();
                        _this.onFilterChanged(e);
                    });
                    // close editor when it loses focus (changes are not applied)
                    edt.lostFocus.addHandler(function () {
                        setTimeout(function () {
                            var ctl = wijmo.Control.getControl(_this['_divEdt']);
                            if (ctl && !ctl.containsFocus()) {
                                _this.closeEditor();
                            }
                        }, 10); //200); // let others handle it first
                    });
                    // get the header cell to position editor
                    var ch = this.grid.columnHeaders, r = ht ? ht.row : ch.rows.length - 1, c = ht ? ht.col : col.index, rc = ch.getCellBoundingRect(r, c), hdrCell = document.elementFromPoint(rc.left + rc.width / 2, rc.top + rc.height / 2);
                    hdrCell = wijmo.closest(hdrCell, '.wj-cell');
                    // show editor and give it focus
                    if (hdrCell) {
                        wijmo.showPopup(div, hdrCell, false, false, false);
                    }
                    else {
                        wijmo.showPopup(div, rc);
                    }
                    edt.focus();
                    // save reference to editor
                    this['_divEdt'] = div;
                    this['_edtCol'] = col;
                };
                /*
                 * Gets the filter for the given column.
                 *
                 * @param col The @see:Column that the filter applies to (or column name or index).
                 * @param create Whether to create the filter if it does not exist.
                 */
                _FlexSheetFilter.prototype.getColumnFilter = function (col, create) {
                    if (create === void 0) { create = true; }
                    // get the column by name or index, check type
                    if (wijmo.isString(col)) {
                        col = this.grid.columns.getColumn(col);
                    }
                    else if (wijmo.isNumber(col)) {
                        col = this.grid.columns[col];
                    }
                    col = wijmo.asType(col, grid.Column);
                    // look for the filter
                    for (var i = 0; i < this['_filters'].length; i++) {
                        if (this['_filters'][i].column == col) {
                            return this['_filters'][i];
                        }
                    }
                    // not found, create one now
                    if (create) {
                        var cf = new sheet._FlexSheetColumnFilter(this, col);
                        this['_filters'].push(cf);
                        return cf;
                    }
                    // not found, not created
                    return null;
                };
                return _FlexSheetFilter;
            }(wijmo.grid.filter.FlexGridFilter));
            sheet._FlexSheetFilter = _FlexSheetFilter;
        })(sheet = grid.sheet || (grid.sheet = {}));
    })(grid = wijmo.grid || (wijmo.grid = {}));
})(wijmo || (wijmo = {}));
//# sourceMappingURL=_FlexSheetFilter.js.map
