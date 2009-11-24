<%
/*
 *	SmartASP Library 0.1
 *	http://code.google.com/p/smartasp/
 *
 *	Copyright (c) 2009 heero
 *	licensed under MIT license
 *
 *	Date: 2009-11-17
 */
 

/// 输入输出命名空间
$.io = {};


/// 上传类
$.io.Upload = function() {
	var t = this;
	
	// 读取文件的Stream
	t._readerStream = Server.CreateObject("ADODB.Stream");
	// 临时Stream
	t._tempStream = Server.CreateObject("ADODB.Stream");
	
	// 存放文件数据及信息的数组
	t._fileArray = [];
	// 指向fileArray对应项的字典
	t._fileHash = {};
	
	// 上传数据总大小
	t._totalBytes = Request.TotalBytes;
};

// 上传类原型
$.io.Upload.prototype = {
	
	/// 获取上传数据总量，单位为kb
	/// @return {Number} 上传数据总大小
	getTotalBytes : function() { return this._totalBytes; },
	
	/// 获取上传文件数
	/// @return {Number} 文件数
	getFileCount : function() { return this._fileArray.length; },
	
	/// 获取文件信息对象
	/// @param {String} 文件索引
	/// @return {Object} 文件信息对象
	getFile : function(index) {
		var type = typeof(index);
		if ("string" === type) {			// 通过字典获取
			return this._fileHash[index];
		} else if ("number" === type) {		// 通过数组获取
			return this._fileArray[index];
		}
	},
	
	/// 读取上传内容
	/// @param {Number} 每次读取的块大小，默认为50kb
	read : function(chunkSize, charset) {
		chunkSize = chunkSize ? chunkSize * 1024 : 50 * 1024;

		if (-1 === String(Request.ServerVariables("CONTENT_TYPE")).indexOf("multipart/form-data") ||
			"POST" !== String(Request.ServerVariables("REQUEST_METHOD")).toUpperCase()
		) {
			throw new Error("文件提交方式出错");
		}
			
		var temp = 0, t = this;			// 临时记录已读取数据的大小
		t._readerStream.type = 1;		// adTypeBinary
		t._readerStream.open();
		while (temp < t._totalBytes) {
			t._readerStream.write(Request.BinaryRead(chunkSize));		// 把数据逐块写入流中
			temp += chunkSize;
		}
		
		// @debug
		//t._readerStream.saveToFile(Server.MapPath("upload.log"), 2);
			
		// 把上传内容转换为字符串，以查找相关子字符串的位置
		t._readerStream.position = 0;
		t._readerStream.type = 2;			// adTypeText
		t._readerStream.charset = "ascii"; 
		
		var ret = t._readerStream.readText(-1),
			separator = ret.substring(0, ret.indexOf("\r\n")),	// 截取分隔符
			separatorLen = separator.length,					// 分隔符长度
			formHead = [];		// 临时存放分隔符位置及长度
			
		// @debug
		//Response.Write(separator);
		
		temp = ret.indexOf(separator, 0);	// 临时记录搜索位置
		// 找到所有报头的开始位置及长度，以便截取
		while (temp != -1) {
			// 分隔符后面紧跟报头，报头后面两个换行后紧跟文件内容
			formHead.push({
				start: temp + separatorLen,
				len: ret.indexOf("\r\n\r\n", temp) - separatorLen - temp
			});
			temp = ret.indexOf(separator, temp + 1);
		}

		var reFileInfo = /name="([^'"]+)";\s*filename="([^'"]+)"\s*Content-Type:\s*([^\s]+)/i,	// 匹配报头中文件信息的的正则表达式
			i = 0, len = formHead.length - 1,
			matchs, fullName, extNamePos;
			
		for (; i < len; i++) {
			temp = t._binToStr(formHead[i].start, formHead[i].len, charset);	// 获取报头字符串
			matchs = temp.match(reFileInfo);		// 通过正则表达式提取文件名、路径、mine类型
			
			if (null == matchs) { continue; }
			
			if (4 === matchs.length) {
				fullName = matchs[2].substr(matchs[2].lastIndexOf("\\") + 1);	// 截取完整路径
				if (-1 === fullName.indexOf(String.fromCharCode(0))) {			// 检查空字符，防上传漏洞
					extNamePos = fullName.lastIndexOf(".");						// 获取扩展名位置
					t._fileArray[i] = {
						fieldName : matchs[1],							// 上传表单字段名
						fileName : fullName.substr(0, extNamePos),		// 截取文件名(不包括扩展名)
						extName : -1 != extNamePos ? fullName.substr(extNamePos + 1) : "",	// 截取扩展名
						mineType : matchs[3],							// 获取MIME类型
						dataPos : formHead[i].start + formHead[i].len + 4,	// 获取数据开始位置
						dataSize : formHead[i + 1].start - separatorLen - formHead[i].start - formHead[i].len - 6		// 获取数据长度
					};
					t._fileHash[matchs[1]] = t._fileArray[i];	// 添加字典索引
				}
			}
		}
	},
	
	/// 把文件读取Stream中指定段的二进制数据读出为字符串
	/// @param {Number} 开始位置
	/// @param {Number} 读取的长度
	/// @return {String} 字符串 
	_binToStr : function(start, len) {
		var t = this, data;
		
		// 复制指定数据段到临时读取器
		t._tempStream.type = 1;
		try {
			t._tempStream.open();
			t._readerStream.position = start;
			t._readerStream.copyTo(t._tempStream, len);
		
			t._tempStream.position = 0;			// 重置指针位置
			t._tempStream.type = 2;				// adTypeText
			t._tempStream.charset = "utf-8";	// 指定编码

			data = t._tempStream.readText(-1);	// 读出为字符串
		} catch (e) {
			throw e;
		} finally {
			t._tempStream.close();
		}
		return data;
	},
	
	/// 生成自动文件名
	/// return {String} 文件名
	getAutoFileName : function() {
		return (new Date()).format("yyyyMMddHHmmss") + parseInt(Math.random() * 100000);
	},
	
	/// 保存数据为文件
	/// @param {String|Number} 要保存的文件的索引
	/// @param {String} 保存路径(绝对路径)
	/// @param {String} 文件名，默认为原文件名
	/// @param {String} 扩展名，默认为原扩展名
	/// @return {String} 完整路径
	saveToFile : function(index, path, fileName, extName) {
		var t = this, file;
		if (t._readerStream.state != 1) {
			throw new Error("无上传数据");
		}
		file = this.getFile(index);
		if (file) {
			extName = extName ? extName.replace(/^\.+/, "") : file.extName;
			if (path.lastIndexOf("\\") != path.length - 1) {
				path += "\\";
			}
			
			var fullPath = path + (fileName || file.fileName) + "." + extName;	// 保存路径

			// 把数据段保存到临时读取器
			t._tempStream.type = 1;
			t._tempStream.open();
			
			// @debug
			//Response.Write(fullPath);
			
			t._readerStream.position = file.dataPos;
			t._readerStream.copyTo(t._tempStream, file.dataSize);
			t._tempStream.saveToFile(fullPath, 2);	// 保存
			t._tempStream.close();
			
			return fullPath;
		} else {
			throw new Error("文件不存在");
		}
	},
	
	/// 关闭文件流
	close : function() {
		this._readerStream.state != 0 && this._readerStream.close();
		this._tempStream.state != 0 && this._tempStream.close();
	}
};
%>