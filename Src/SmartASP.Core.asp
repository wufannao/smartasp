<%@LANGUAGE="JAVASCRIPT" CODEPAGE="65001"%>
<%
/*
 *	SmartASP Library 0.1
 *	http://code.google.com/p/smartasp/
 *
 *	Copyright (c) 2009 heero
 *	licensed under MIT license
 *
 *	Date: 2009-11-24
 */


// 全局对象
var $;


// 在函数中执行减少全局变量
(function() {

// ----------------------------------------------------------------
// Request

/// 获取请求参数值
/// @param {String} 参数名
/// @param {Number} 搜索的参数集合，0为QueryString，1为Form，省略时为不限制
/// @return {String} 参数值
$ = function(name, method) {
	var value;
	if (0 == method) {
		value = Request.QueryString(name);
	} else if (1 == method) {
		value = Request.Form(name);
	} else {
		value = Request(name);
	}
	value = String(value);
	return "undefined" === value ? "" : value.trim();
};

/// 获取客户端IP
/// @return {String} IP
$.getIp = function() {
	var proxy = String(Request.ServerVariables("HTTP_X_FORWARDED_FOR")),
		ip = proxy && proxy.indexOf("unknown") != -1 ? proxy.split(/,;/g)[0] : String(Request.ServerVariables("REMOTE_ADDR"));
		
	ip = ip.trim().substring(0, 15);
	return "::1" === ip ? "127.0.0.1" : ip;
};

// ---------------------------------------------------------------- */


/// 标识版本
$.version = "0.1 Build 20091124";


/// 配置对象
$.config = {
	prefix : "SmartASP_",	// 应用程序前缀
	isDebug : false			// 是否调试模式
};

  
// ----------------------------------------------------------------
// String扩展

// 匹配头尾的空白
var reWhiteSpaces = /^\s+|\s+$/g;

/// 去掉当前字符串两端的某段字符串
/// @param {String} 要去掉的字符串，默认为空白
/// @return {String} 修整后的字符串
String.prototype.trim = function(str) {
	return this.replace(null == str ? reWhiteSpaces : new RegExp("^" + str + "+|" + str + "+$", "g"), "");
};

/// 从左边开始截取一定长度的子字符串
/// @param {Number} 长度
/// @return {String} 子字符串
String.prototype.left = function(n) {
	return this.substr(0, n);
};

/// 从右边开始截取一定长度的子字符串
/// @param {Number} 长度
/// @return {String} 子字符串
String.prototype.right = function(n) {
	return this.slice(-n);
};

/// 获取当前字符串的数字类型值
/// @return {Number} 当前字符串的数字类型值
String.prototype.toNumber = function() {
	return $.convert.toNumber(this);
};

/// 获取当前字符串的整型值
/// @return {Number} 当前字符串的整型值
String.prototype.toInt = function() {
	return $.convert.toInt(this);
};

/// 获取当前字符串的浮点值
/// @return {Number} 当前字符串的浮点值
String.prototype.toFloat = function() {
	return $.convert.toFloat(this);
};

/// 获取当前字符串的布尔值
/// @return {Boolean} 当前字符串的布尔值
String.prototype.toBool = function() {
	return $.convert.toBool(this);
};

/// 获取当前字符串的日期值
/// @return {Date} 当前字符串的日期值
String.prototype.toDateTime = function() {
	return $.convert.toDateTime(this);
};

// ---------------------------------------------------------------- */


// ----------------------------------------------------------------
// Date扩展

// 把数字转换成两位数的字符串
function toTwoDigit(num) {
	return num < 10 ? "0" + num : num;
}

// 临时记录正在转换的日期
var tempYear, tempMonth, tempDate, tempHour, tempMinute, tempSecond;

// 格式替换函数
function getDatePart(part) {
	switch (part) {
		case "yyyy": return tempYear;
		case "yy": return tempYear.toString().slice(-2);
		case "MM": return toTwoDigit(tempMonth);
		case "M": return tempMonth;
		case "dd": return toTwoDigit(tempDate);
		case "d": return tempDate;
		case "HH": return toTwoDigit(tempHour);
		case "H": return tempHour;
		case "hh": return toTwoDigit(tempHour > 12 ? tempHour - 12 : tempHour);
		case "h": return tempHour > 12 ? tempHour - 12 : tempHour;
		case "mm": return toTwoDigit(tempMinute);
		case "m": return tempMinute;
		case "ss": return toTwoDigit(tempSecond);
		case "s": return tempSecond;
		default: return part;
	}
}

/// 返回指定格式的日期字符串
/// @param {String} 格式字符串
/// @return {String} 指定格式的日期字符串
Date.prototype.format = function(formation) {
	tempYear = this.getFullYear();
	tempMonth = this.getMonth() + 1;
	tempDate = this.getDate();
	tempHour = this.getHours();
	tempMinute = this.getMinutes();
	tempSecond = this.getSeconds();

	return formation.replace(/y+|m+|d+|h+|s+|H+|M+/g, getDatePart);
};

// ---------------------------------------------------------------- */


// ----------------------------------------------------------------
// Error扩展

/// 根据调试标志修改异常信息
/// @param {String} 附加的异常信息
/// @return {Error} 当前异常对象
Error.prototype.addMsg = function(msg) {
	this.message = $.config.isDebug ? this.message + " <br /> \r\n" + msg : msg;
	return this;
};

// ---------------------------------------------------------------- */


// ----------------------------------------------------------------
// 类型转换

var reFalse = /^false$/i,		// 匹配字符串false(不分大小写)
	reDateTime = /^(\d{4})-(\d{1,2})(?:-(\d{1,2})(?: (\d{1,2}):(\d{1,2})(?::(\d{1,2}))?))?$/;	// 粗略匹配日期

/// 类型转换对象
$.convert = {
	
	/// 把值转换为数字类型
	/// @param {Mixed} 值
	/// @return {Number} 转换后的数字，如果该值无法转换为数字，则返回0
	toNumber : function(value) {
		value = Number(value);
		return isNaN(value) ? 0 : value;
	},
	
	/// 把值转换为整数类型
	/// @param {Mixed} 值
	/// @return {Number} 转换后的整数，如果该值无法转换为整数，则返回0
	toInt : function(value) {
		value = parseInt(value);
		return isNaN(value) ? 0 : value;
	},
	
	/// 把值转换为浮点数类型
	/// @param {Mixed} 值
	/// @return {Number} 转换后的浮点数，如果该值无法转换为浮点数，则返回0
	toFloat : function(value) {
		value = parseFloat(value);
		return isNaN(value) ? 0 : value;
	},
	
	/// 把字符串转换为布尔值
	/// @param {Mixed} 值
	/// @return {Number} 转换后的布尔值
	toBool : function(value) {
		return null == value || 0 == value || "" == value || reFalse.test(value) ? false : true;
	},
	
	/// 把字符串转换为日期类型
	/// @param {String} 字符串，格式为yyyy-MM-dd HH:mm:ss。
	/// @return 转换后的日期，如果该字符串无法转换为日期或日期不合法，则返回undefined
	toDateTime : function(value) {
		if (reDateTime.test(value)) {
			var year = parseInt(RegExp.$1), month = parseInt(RegExp.$2), day = RegExp.$3;
				hour = RegExp.$4, minute = RegExp.$5, second = RegExp.$6;
				
			if (year < 1970 || month < 1 || month > 12) {
				return;
			}
			if (day !== "") {
				day = parseInt(day);
				if (day < 1 || day > (new Date(year, month, 0)).getDate()) {
					return;
				}
			}
			
			if (hour !== "" && minute !== "") {
				hour = parseInt(hour); minute = parseInt(minute);
				if (hour < 0 || hour > 23 || minute < 0 || minute > 60) {
					return;
				}
				if (second != "") {
					second = parseInt(second);
					if (second < 0 || second > 59) {
						return;
					}
				}
			}
			
			return new Date(year, month - 1, day, hour, minute, second);
		}
	}
};

// ---------------------------------------------------------------- */


// ----------------------------------------------------------------
// JSON处理

var reEscapable = /[\\\"\x00-\x1f\x7f-\x9f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,
	meta = {
		'\b': '\\b',
		'\t': '\\t',
		'\n': '\\n',
		'\f': '\\f',
		'\r': '\\r',
		'"' : '\\"',
		'\\': '\\\\'
	};

// 编码特殊字符
function encodeChar($0) {
	var char = meta[$0];
	return typeof c === "string" ? c : "\\u" + ('0000' + $0.charCodeAt(0).toString(16)).slice(-4);
}
// 用引号括起字符串
function quoteStr(str) {
	reEscapable.lastIndex = 0;
	return '"' + str.replace(reEscapable, encodeChar) + '"';
}

/// JSON处理对象
$.json = {
	
	/// 把值序列化为JSON字符串
	/// @param {Mixed} 值
	/// @return {String} JSON字符串
	stringify : function(value) {
		var temp, i;
		switch (typeof value) {
			case "undefined":
				return "null";
			
			case "number":
			case "boolean":
			case "function":
				return String(value);
				
			case "string":
				return quoteStr(value);
				
			case "object":
				if (!value) {
					return "null";
				}
				if (value.toJSON) {
					return value.toJSON();
				}
				switch (value.constructor) {
					case Date:
						return "new Date(" + value.getMilliseconds() + ")";
						
					case Array:
						temp = [];
						for (i = 0; i < value.length; i++) {
							temp[i] = $.json.stringify(value[i]);
						}
						return "[" + temp.join(",") + "]";
						
					case RegExp:
						return String(value);
						
					default:
						temp = [];
						for (i in value) {
							if (Object.hasOwnProperty.call(value, i)) {
								temp.push(quoteStr(i) + ":" + $.json.stringify(value[i]));
							}
						}
						return "{" + temp.join(",") + "}";
				}
		}
		
	},
	
	/// 把JSON字符串反序列化为值
	/// @param {String} JSON字符串
	/// @return {Mixed} 值
	parse : function(str) {
		var json;
		try {
			eval("json = " + str);
		} catch(e) {
		}
		return json;
	}
};

// ---------------------------------------------------------------- */


// ----------------------------------------------------------------
// 缓存处理

/// 缓存操作类
$.cache = {
	
	/// 设置值类型缓存
	/// @param {String} 缓存名
	/// @param {Mixed} 缓存值
	addValue : function(name, value) {
		Application.Lock();
		Application($.config.prefix + name) = value;
		Application.UnLock();
	},
		
	/// 设置对象类型缓存
	/// @param {String} 缓存名
	/// @param {Mixed} 缓存对象
	addObject : function(name, obj) {
		$.cache.addValue(name, $.json.stringify(obj));
	},
	
	/// 获取值类型缓存
	/// @param {String} 缓存名
	/// @return {String} 缓存值
	getValue : function(name) {
		var value = String(Application($.config.prefix + name));
		return "undefined" === value ? "" : value;
	},
	
	/// 获取对象类型缓存
	/// @param {String} 缓存名
	/// @return {Mixed} 缓存对象
	getObject : function(name) {
		return $.json.parse($.cache.getValue(name));
	},
	
	/// 清理缓存
	/// @param {String} 缓存名，省略时为清理全部
	del : function(name) {
		Application.Lock();
		if (null == name || "" == name) {
			Application.Contents.RemoveAll();
		} else {
			Application.Contents.Remove($.config.prefix + name);
		}
		Application.UnLock();
	}
};

// ---------------------------------------------------------------- */


// ----------------------------------------------------------------
// 验证

var reNumbers = /^\d+(?:\s*\,\s*\d+)*$/;	// 匹配一个或多个数字

/// 验证对象
$.validate = {
	
	/// 验证指定值是否用逗号隔开的一段非负整数
	/// @param {Mixed} 值
	/// @return {Boolean} 是否符合条件
	isNumbers : function(value) {
		return reNumbers.test(value);
	}
	
};

// ---------------------------------------------------------------- */

})();
%>