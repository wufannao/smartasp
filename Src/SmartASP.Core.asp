<%@LANGUAGE="JAVASCRIPT" CODEPAGE="65001"%>
<%
/*
 *	Core Module v0.2.1
 *	of SmartASP Library v0.2
 *
 *	http://code.google.com/p/smartasp/
 *
 *	Copyright (c) 2009-2010 heero
 *	licensed under MIT license
 *
 *	Date: 2009/12/7
 */


// 全局对象
var $, nothing;


// 在函数中执行减少全局变量
(function() {

// ----------------------------------------------------------------
// Request

/// 获取请求参数值
/// @param {String} 参数名
/// @param {String} 搜索的参数集合，省略时为不限制
/// 	"get": QueryString
/// 	"post": Form
/// 	"serv": ServerVariables
/// @return {String} 参数值
$ = function(name, subset) {
	var value;
	if (!subset) {
		value = Request(name);		
	} else {
		switch (subset.toLowerCase()) {
			case "get":
				value = Request.QueryString(name);
			break;
			
			case "post":
				value = Request.Form(name);
			break;
			
			case "serv":
				value = Request.ServerVariables(name);
			break;
		}
	}
	
	return value === nothing || value.Item === nothing ? "" : value.Item.trim();
};

/// 获取客户端IP
/// @return {String} IP
$.getIp = function() {
	var proxy = $("HTTP_X_FORWARDED_FOR", "serv"),
		ip = proxy && proxy.indexOf("unknown") != -1 ? proxy.split(/,;/g)[0] : $("REMOTE_ADDR", "serv");
		
	ip = ip.trim().substring(0, 15);
	return "::1" === ip ? "127.0.0.1" : ip;
};

// ---------------------------------------------------------------- */


/// 标识版本
$.version = "0.2.1 Build 20091126";


/// 配置对象
$.config = {
	prefix : "SmartASP_",	// 应用程序前缀
	isDebug : true			// 是否调试模式
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
	return Number(this);
};

/// 获取当前字符串的整型值
/// @return {Number} 当前字符串的整型值
String.prototype.toInt = function() {
	return parseInt(this);
};

/// 获取当前字符串的浮点值
/// @return {Number} 当前字符串的浮点值
String.prototype.toFloat = function() {
	return parseFloat(this);
};

/// 获取当前字符串的布尔值
/// @return {Boolean} 当前字符串的布尔值
String.prototype.toBool = function() {
	return $.convert.toBool(this);
};

/// 获取当前字符串的日期值
/// @return {Date} 当前字符串的日期值
String.prototype.toDate = function() {
	return $.convert.toDate(this);
};

// ---------------------------------------------------------------- */


// ----------------------------------------------------------------
// Date扩展

// 把数字转换成两位数的字符串
function toTwoDigit(num) { return num < 10 ? "0" + num : num; }

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
	
	/// 把字符串转换为布尔值
	/// @param {Mixed} 值
	/// @return {Number} 转换后的布尔值
	toBool : function(value) {
		return null == value || 0 == value || "" == value || reFalse.test(value) ? false : true;
	},
	
	/// 把字符串转换为日期类型
	/// @param {String} 字符串，格式为yyyy-MM-dd HH:mm:ss。
	/// @return 转换后的日期，如果该字符串无法转换为日期或日期不合法，则返回undefined
	toDate : function(value) {
		if (reDateTime.test(value)) {
			var year = Number(RegExp.$1), month = Number(RegExp.$2), day = RegExp.$3;
				hour = RegExp.$4, minute = RegExp.$5, second = RegExp.$6;
				
			if (year < 1970 || month < 1 || month > 12) {
				return;
			}
			if (day !== "") {
				day = Number(day);
				if (day < 1 || day > (new Date(year, month, 0)).getDate()) {
					return;
				}
			}
			
			if (hour !== "" && minute !== "") {
				hour = Number(hour); minute = Number(minute);
				if (hour < 0 || hour > 23 || minute < 0 || minute > 60) {
					return;
				}
				if (second != "") {
					second = Number(second);
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

/// 缓存操作
$.cache = {
	
	/// 设置缓存
	/// @param {String} 缓存名
	/// @param {Mixed} 缓存值
	/// @param {Boolean} 是否存储为JSON字符串，默认为否；如果缓存的变量是引用类型，请设为true
	set : function(name, value, isToSerialize) {
		if (isToSerialize) {
			value = $.json.stringify(value);
		}
		Application.Lock();
		Application($.config.prefix + name) = value;
		Application.UnLock();
	},
	
	/// 获取缓存值
	/// @param {String} 缓存名
	/// @param {Boolean} 是否把值反序列化，默认为否；如果缓存的是JSON字符串，可以设为true以获取对应的值
	/// @return {String} 缓存值
	get : function(name, isToDeserialize) {
		var value = Application($.config.prefix + name);
		if (isToDeserialize && value !== nothing) {
			value = $.json.parse(value);
		}
		return value;
	},
	
	/// 清理缓存
	/// @param {String} 缓存名，省略时为清理全部
	del : function(name) {
		Application.Lock();
		if (name) {
			Application.Contents.Remove($.config.prefix + name);
		} else {
			Application.Contents.RemoveAll();
		}
		Application.UnLock();
	}
};

// ---------------------------------------------------------------- */


// ----------------------------------------------------------------
// Cookie

// Cookie过期时间格式
var EXPIRESWITHUNIT = /[smhdMy]$/,
	TIMEUNITS = {
		"s" : 1,
		"m" : 60,
		"h" : 60 * 60,
		"d" : 24 * 60 * 60,
		"M" : 30 * 24 * 60 * 60,
		"y" : 365 * 24 * 60 * 60
	};

/// Cookie操作
$.cookie = {
	
	/// 获取Cookie
	/// @param {String} Cookie名
	/// @return {String,Object} Cookie值或Cookie字典
	get : function(name) {
		var cookie = Request.Cookies($.config.prefix + name), value;	
		if (cookie.HasKeys) {
			value = {};
			for (var i = 1, len = cookie.Count; i <= len; i++) {
				value[cookie.Key(i)] = cookie.Item(i).trim();
			}
			return value;
		} else {
			return cookie.Item.trim();
		}
	},

	/// 设置Cookie
	/// @param {String} Cookie名
	/// @param {Mixed} Cookie值
	/// @param {Number,Date,String} 过期时间，可带单位：
	///		y: 年
	/// 	M: 月
	///		d: 日
	///		h: 小时
	///		m: 月
	///		s: 秒
	/// @param {String} 域
	/// @param {String} 路径
	/// @param {Boolean} 是否仅把Cookie发送给受保护的服务器(https)
	set : function(name, value, expires, domain, path, secure) {
		name = $.config.prefix + name;

		if ("string" === typeof value) {
			Response.Cookies(name) = value;
		} else {
			for (var i in value) {
				if (Object.hasOwnProperty.call(value, i)) {
					Response.Cookies(name)(i) = value[i];
				}
			}
		}
		
		if (expires) {
			var date, unit;
			if (expires instanceof Date) {
				date = expires;
			} else {
				if ("string" === typeof expires && EXPIRESWITHUNIT.test(expires)) {
					expires = expires.substring(0, expires.length - 1);
					unit = RegExp.lastMatch;
				}
				if (!isNaN(expires)) {
					date = new Date();
					date.setTime(date.getTime() + expires * TIMEUNITS[unit || "m"] * 1000);
				}
			}
			
			date && (Response.Cookies(name).Expires = date.format("M/d/yyyy H:m:s"));
		}

		domain && (Response.Cookies(name).Domain = domain);
		path && (Response.Cookies(name).Path = path);
		secure && (Response.Cookies(name).Secure = secure);
	},

	/// 删除Cookie
	/// @param {String} Cookie名，省略时为删除全部
	del : function(name) {
		if (name) {
			Response.Cookies($.config.prefix + name).Expires = "1/1/1980";
		} else {
			for (var i = 1, len = Request.Cookies.Count; i <= len; i++) {
				Response.Cookies.Item(i).Expires = "1/1/1980";
			}
		}
	}
};

// ---------------------------------------------------------------- */


// ----------------------------------------------------------------
// Session

/// Session操作
$.session = {
	
	/// 获取Session值
	/// @param {String} Session名
	/// @return {Mixed} Session值
	get : function(name) {
		return Session($.config.prefix + name);
	},
	
	/// 设置Session
	/// @param {String} Session名
	/// @param {Mixed} Session值
	set : function(name, value) {
		Session($.config.prefix + name) = value;
	},

	///	清理Session
	/// @param {String} Session名，省略时为清理全部
	del : function(name) {
		if (name) {
			Session.Contents.Remove($.config.prefix + name);
		} else {
			Session.Contents.RemoveAll();
		}
	}
};

// ---------------------------------------------------------------- */

})();
%>