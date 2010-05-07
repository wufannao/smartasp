<%
/*
 *	Validate Module v0.1
 *	of SmartASP Library v0.2
 *
 *	http://code.google.com/p/smartasp/
 *
 *	Copyright (c) 2009-2010 heero
 *	licensed under MIT license
 *
 *	Date: 2009/11/26
 */


(function() {

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

})();
%>