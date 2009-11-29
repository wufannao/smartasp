<%
/*
 *	Data Module v0.2
 *	of SmartASP Library 0.2
 *
 *	http://code.google.com/p/smartasp/
 *
 *	Copyright (c) 2009 heero
 *	licensed under MIT license
 *
 *	Date: 2009-11-29
 */


(function() {

/// 数据库操作命名空间
$.data = {};


/// 数据库操作辅助类
/// @overload
/// 	@param {String} 数据库服务器地址
/// 	@param {String} 用户名
/// 	@param {String} 密码
/// 	@param {String} 数据库名
/// 	@param {String} 数据库驱动，默认为SQL Server
/// @overload
/// 	@param {String} 文件数据库地址
/// 	@param {String} 数据库驱动，默认为Access
$.data.DbHelper = function() {
	var t = this;
	
	// Connection 对象
	t._conn = Server.CreateObject("ADODB.Connection");
	// Command 对象
	t._cmd = Server.CreateObject("ADODB.Command");
	// 数据库访问次数
	t._dbAccessTime = 0;
	
	var argCount = arguments.length;
	if (4 == argCount || 5 == argCount) {
		t._conn.ConnectionString = "driver={" + (arguments[4] || "SQL Server") + "};server=" + arguments[0] + ";uid=" + arguments[1] + ";pwd=" + arguments[2] + ";database=" + arguments[3];
	} else if (1 == argCount || 2 == argCount) {
		t._conn.ConnectionString = "Provider=" + (arguments[1] || "Microsoft.Jet.OLEDB.4.0") + ";Data Source=" + arguments[0];
	}
};

/// 数据库操作辅助类原型
$.data.DbHelper.prototype = {
	
	/// 连接数据库
	connect : function() {
		var conn = this._conn;
		if (0 == conn.State) {
			try {
				conn.Open();
				this._cmd.ActiveConnection = conn;
			} catch (e) {
				throw e.addMsg("数据库连接出错，请检查连接字符串：" + conn.ConnectionString);
			}
		}
	},
	
	/// 设置数据库操作超时时间
	/// @param {Number} 超时时间，单位秒
	setTimeout : function(timeout) { this._cmd.CommandTimeout = timeout; },
	
	/// 获取数据库访问次数
	/// @return {Number} 数据库访问次数
	getAccessTime : function() { return this._dbAccessTime; },
	
	/// 创建命令参数
	/// @param {String} 参数名
	/// @param {String,Number} 参数类型
	/// 	int : 3,
	/// 	bigint : 20,
	/// 	smallint : 2,
	/// 	tinyint : 16,
	/// 	decimal : 14,
	/// 	double : 5,		
	/// 	char : 129,
	/// 	nchar : 130,
	/// 	varchar : 200,
	/// 	text : 201,
	/// 	nvarchar : 202,
	/// 	ntext : 203,
	/// 	datetime : 135
	/// @param {Mixed} 参数值
	/// @param {Number} 参数长度限制
	/// @param {Number} 参数方向，默认为传入
	///		input : 1,
	///		output : 2,
	///		inputoutput : 3,
	///		returnvalue : 4
	/// @return {Object} 参数对象
	createParam : function(name, type, value, size, direction) {
		null == direction && (direction = 1);	// 默认为传入参数

		var iType = typeof type === "number" ?  type : dataTypes[type.toLowerCase()];
		if (!iType) {
			throw new Error("无此参数类型：" + type);
		}
		return this._cmd.CreateParameter("@" + name, iType, direction, size, value);
	},
	
	/// 附加命令参数
	/// @param {Array,Object} 参数对象
	attachParams : function() {
		var i = -1,
			params = arguments[0] instanceof Array ? arguments[0] : arguments,
			len = params.length,
			cmdParams = this._cmd.Parameters;

		while (++i < len) {
			cmdParams.Append(params[i]);
		}
	},
	
	/// 复制参数
	/// @param {Object,Array} 要复制的参数
	/// @return {Array} 原参数及其副本
	copyParams : function(params) {
		if (params instanceof Array === false) {
			params = [params];
		}
		var len = params.length, i = -1, copy = [], p;
		while (++i < len) {
			p = params[i];
			copy[i] = this._cmd.CreateParameter(p.Name, p.Type, p.Direction, p.Size, p.Value);
		}

		return params.concat(copy);
	},
	
	/// 清理所有命令参数
	clearParams : function() {
		var params = this._cmd.Parameters;
		while (params.Count > 0) {
			params.Delete(0);
		}
	},
	
	/// 执行命令
	/// @param {String} 命令文本
	/// @param {Object,Array} 命令参数
	/// @param {Number} 命令类型，1为sql语句，4为存储过程，默认为sql语句
	///	@param {Boolean} 是否请求数据
	/// @return {RecordSet} 数据集
	_execute : function(text, params, type, isQueryData) {
		var result, cmd = this._cmd;
		
		cmd.CommandType = type != null ? type : 1; 
		cmd.CommandText = text;

		// 添加参数
		params && this.attachParams(params);
		
		this._dbAccessTime++;
		
		try {
			if (isQueryData) {
				result = cmd.Execute();
			} else {
				cmd.Execute();
			}
		} catch (e) {
			var msg = ["执行命令 { " + text + " } 出错，参数:"], i = -1, len = cmd.Parameters.Count, p;
			while (++i < len) {
				p = cmd.Parameters(i);
				msg.push(p.Name + ", " + p.Type + ", " + p.Value);
			}
			throw e.addMsg(msg.join(" <br /> \r\n"));
		} finally {
			this.clearParams();
		}
		
		if (isQueryData) { return result; }
	},

	/// 执行命令并返回记录集
	/// @param {String} 命令文本
	/// @param {Object,Array} 命令参数
	/// @param {Number} 命令类型，1为sql语句，4为存储过程，默认为sql语句
	/// @return {RecordSet} 数据集
	executeReader : function(text, params, type) {
		return this._execute(text, params, type, true);
	},
	
	/// 执行命令，无返回结果
	/// @param
	///		@refer executeReader
	executeNonQuery : function(text, params, type) {
		return this._execute(text, params, type, false);
	},

	/// 执行命令并返回结果集合中第一行第一列的值
	/// @param
	///		@refer executeReader
	/// @return {Mixed} 结果集中第一行第一列的值
	executeScalar : function(text, params, type) {
		var rs = this.executeReader(text, params, type), value;
		!rs.EOF && (value = rs(0).value);
		rs.close();
		return value;
	},

	/// 执行命令并返回Json数据
	/// @param
	///		@refer executeReader
	/// @return {Array} JSON对象数组
	executeJson : function(text, params, type) {
		var rs = this.executeReader(text, type, isKeepParams), json;
		return this.rsToJson(rs);
	},
	
	/// 开始事务
	beginTrans : function() {
		this._conn.BeginTrans();
	},
	
	/// 回滚事务
	rollBackTrans : function() {
		this._conn.RollBackTrans();
	},

	/// 提交事务
	commitTrans : function() {
		this._conn.CommitTrans();
	},
	
	/// @param 把记录集转换成JSON对象数组
	/// @param {RecordSet} 记录集
	/// @return {Array} JSON对象数组
	rsToJson : function(rs) {
		var objs = [], i, len, value;
		
		while (!rs.EOF) {	// 循环转换每行数据
			var obj = {};		// 建立新对象
			i = -1; len = rs.Fields.Count;
			while (++i < len) {		// 循环设置对象属性
				value = rs(i).Value;	// 如果对rs(i).Value进行多次访问而且该字段类型为TEXT或NTEXT，rs(i).Value的值会丢失	
				obj[rs(i).Name] = "number" != typeof value ? String(value) : value;
			}
			// 放入数组
			objs.push(obj);
			// 移动到下一条数据
			rs.MoveNext();
		}
		
		rs.Close();		// 关闭记录集
		
		return objs;
	},
	
	/// 释放数据连接
	close : function() {
		this._conn.State != 0 && this._conn.Close();
		this._conn = null;
	}
};


// 枚举数据类型
var dataTypes = {
	int : 3,
	bigint : 20,
	smallint : 2,
	tinyint : 16,
	
	decimal : 14,
	double : 5,
		
	char : 129,
	nchar : 130,
	varchar : 200,
	text : 201,
	nvarchar : 202,
	ntext : 203,
	datetime : 135
};
		  
})();


	
	/*	
	// <summary>分页类</summary>
	// <param name="total">记录总数</param>
	// <param name="pageSize">每页记录数</param>
	// <param name="curPage">当前页码</param>
	function Pagination(total, pageSize, curPage) {
		this.pageCount = Math.ceil(total / pageSize);		// 计算页数
		this.curPage = Math.min(Math.max(1, curPage), this.pageCount);	// 修正页码
		this.pageSize = pageSize;
	};
*/

%>