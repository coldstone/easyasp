if (typeof(SWFUpload) === "function") {
	SWFUpload.extend("CheckFiles");
	SWFUpload.extendCallback("fileChecked","file_check_handler");
	SWFUpload.extend("SetFileStatus");
	SWFUpload.extend("SetPostParam");
	SWFUpload.extend("RemovePostParam");
}

function $_(selecter){
	if(typeof selecter!="string")return selecter;
	try{
		return document.getElementById(selecter);
	}catch(ex){
		return null;
	}
}
function formatBytes(bytes) {
	var s = ['B', 'KB', 'MB', 'GB', 'TB', 'PB'];
	var e = Math.floor(Math.log(bytes)/Math.log(1024));
	return (bytes/Math.pow(1024, Math.floor(e))).toFixed(2)+" "+s[e];
}
function timeString(used,cn){
	if(cn!==true)cn=false;
	var hours = Math.floor(used/3600);
	used = used - hours * 3600;
	var minutes = Math.floor(used/60);
	var seconds = (used % 60).toFixed(2);
	var ret = seconds+(cn?"秒":"");
	if(minutes>0){
		ret = minutes + (cn?"分":":") + ret;
	}
	if(hours>0){
		ret = hours + (cn?"小时":":") + ret;
	}
	return ret;
}

if (typeof(SWFUpload) === "function") {
	SWFUpload.handler = {};
	SWFUpload.handler.stoped=false;
	SWFUpload.handler.onlyOne=false;
	SWFUpload.prototype.initSettings = (function (oldInitSettings) {
		return function () {
			if (typeof(oldInitSettings) === "function") {
				oldInitSettings.call(this, []);
			}
			this.processer={};
			this.ensureDefault = function (settingName, defaultValue) {
				this.settings[settingName] = (this.settings[settingName] == undefined) ? defaultValue : this.settings[settingName];
			};
			this.ensureDefault("bind_id", null);
			this.ensureDefault("destroy_handler", null);
			this.ensureDefault("file_queue_start_handler", null);
			this.ensureDefault("auto", false);
			delete this.ensureDefault;
			this.processer.totalBytes=0;
			this.processer.uploadTotalBytes=0;
			this.processer.uploadFileBytes=0;
			this.processer.speed={
				lasttime:null,lastbytes:0,value:0,start:null,end:null,time_used:0
			};
			this.processer.bind=$_(this.settings["bind_id"]);
			this.processer.user_file_queued_handler = this.settings.file_queued_handler;
			this.processer.user_file_queue_error_handler = this.settings.file_queue_error_handler;
			this.processer.user_upload_start_handler = this.settings.upload_start_handler;
			this.processer.user_upload_error_handler = this.settings.upload_error_handler;
			this.processer.user_upload_progress_handler = this.settings.upload_progress_handler; 
			this.processer.file_queue_start_handler = this.settings.file_queue_start_handler;
			this.processer.user_upload_success_handler = this.settings.upload_success_handler;
			this.processer.user_upload_complete_handler = this.settings.upload_complete_handler;
			this.processer.debug_handler= this.settings.debug_handler;
			
			this.settings.file_queued_handler = SWFUpload.handler.fileQueuedHandler;
			this.settings.file_queue_error_handler = SWFUpload.handler.fileQueueErrorHandler;
			this.settings.upload_start_handler = SWFUpload.handler.uploadStartHandler;
			this.settings.upload_error_handler = SWFUpload.handler.uploadErrorHandler;
			this.settings.upload_progress_handler = SWFUpload.handler.uploadProgressHandler;
			this.settings.upload_success_handler = SWFUpload.handler.uploadSuccessHandler;
			this.settings.upload_complete_handler = SWFUpload.handler.uploadCompleteHandler;
			this.settings.file_dialog_complete_handler = SWFUpload.handler.fileDialogComplete;
			this.settings.debug_handler = SWFUpload.handler.debug;
			this.settings.file_queue_start_handler= SWFUpload.handler.fileQueueStart;
		};
	})(SWFUpload.prototype.initSettings);
	
	
	SWFUpload.handler.debug = function(msg){
		alert(msg);
	};
	
	SWFUpload.prototype.startUploadFiles=function(A,B){
		if(this.Status().busy){return}
		if(this.Status().queued>0){
			this.setButtonDisabled();
			SWFUpload.handler.stoped=false;
			this.processer.speed.start=new Date();
			this.startUpload(A);
			if(B===true)SWFUpload.handler.onlyOne=true;
			while(Files.length>0){Files.pop();}
			if(A!=undefined){
				if(this.processer.bind==null)return;
				$_("b_"+A).style.backgroundColor="#f6f6f6";
				$_(A).style.border="1px solid #ddd";
			}
		}else{
			$_("message").innerHTML = " *没有可上传的文件。请先选择文件。";	
		}
	},
	SWFUpload.handler.fileQueuedHandler = function (file) {
		this.processer.totalBytes+=file.size;
		if(this.processer.bind==null)return;
		if($_(file.id)==null){
			var o = this.processer.bind;
			var list = document.createElement("div");
			list.className = "filelist fl";
			list.id=file.id;
			o.appendChild(list);
			
			var processbar = document.createElement("div");
			processbar.id="b_" + file.id;
			processbar.className = "process_bar";
			list.appendChild(processbar);
			
			var infobar = document.createElement("div");
			infobar.className = "info_bar";
			infobar.id ="i_" + file.id;
			list.appendChild(infobar);
		}
		var filename = file.name;
		if(filename.length>30){filename = filename.replace(/^(.{13})(.+?)(.{13})$/igm,"$1&sdot;&sdot;&sdot;$3");}
		$_("b_" + file.id).style.width=0;
		$_("i_" + file.id).innerHTML="<ul><li class=\"w_name\"><span class=\"s_name\">" + filename + "</span><span class=\"gray\">(" + formatBytes(file.size) + ")</span></li><li class=\"w_process\" id=\"p_" + file.id + "\">等待中</li><li class=\"w_size\" id=\"sp_" + file.id + "\">0</li>"
		+"<li class=\"w_act\" id=\"a_" + file.id + "\">"
		+"<a href=\"javascript:void(0)\" onclick=\"SWFUpload.instances['" + this.movieName + "'].cancelUpload('" + file.id + "',true,true);\">取消</a>"
		+"</li></ul>";
		if (typeof this.processer.user_file_queued_handler === "function") return this.processer.user_file_queued_handler.call(this, file);
	};
	
	
	SWFUpload.handler.fileDialogComplete = function(){
		if(this.settings["auto"] && this.Status().queued>0)this.startUploadFiles();
	};
	
	SWFUpload.handler.fileQueueErrorHandler = function (file, errorCode, message) {
		var errorName='';
		switch (errorCode)
		{
			case SWFUpload.QUEUE_ERROR.QUEUE_LIMIT_EXCEEDED:
				errorName = "只能同时上传 "+this.settings.file_upload_limit+" 个文件，超过限制的文件被忽略";
				break;
			case SWFUpload.QUEUE_ERROR.FILE_EXCEEDS_SIZE_LIMIT:
				errorName = "选择的文件超过了当前大小限制："+this.settings.file_size_limit +"，文件被忽略";
				break;
			case SWFUpload.QUEUE_ERROR.ZERO_BYTE_FILE:
				errorName = "零大小文件，文件被忽略";
				break;
			case SWFUpload.QUEUE_ERROR.INVALID_FILETYPE:
				errorName = "文件扩展名必需为："+this.settings.file_types_description+" ("+this.settings.file_types+")，文件被忽略";
				break;
			default:
				errorName = "未知错误，文件被忽略";
				break;
		}
		var msg1 = "";
		if(file!=null)msg1=file.name + " ";
		$_("message").innerHTML = " *"+msg1+errorName;
		if (typeof this.processer.user_file_queue_error_handler === "function") return this.processer.user_file_queue_error_handler.call(this, file, errorCode, message);
	};
	
	SWFUpload.handler.uploadStartHandler = function (file) {
		this.processer.speed.lasttime = new Date();
		this.processer.speed.lastbytes = 0;
		this.processer.speed.start = new Date();
		if(typeof BeforeUploadCallBack=="function")BeforeUploadCallBack.apply(this,[file]);
		if (typeof this.processer.user_upload_start_handler === "function") return this.processer.user_upload_start_handler.call(this, file);
	};
	
	SWFUpload.handler.uploadProgressHandler = function (file, bytesComplete, bytesTotal) {
		this.processer.uploadTotalBytes=this.processer.uploadFileBytes+bytesComplete;
		if(this.processer.speed.lasttime==null){
			this.processer.speed.lasttime = new Date();
			this.processer.speed.lastbytes = bytesComplete;
		}else{
			var time = 	(new Date())-this.processer.speed.lasttime;
			var bytes = bytesComplete-this.processer.speed.lastbytes;
			if(time>0 && bytes>0){
				bytes = (bytes/time) * 1000;
				this.processer.speed.value = formatBytes(bytes) + "/S";
			}
		}
		this.processer.speed.end=new Date();
		var time_used=0;
		if(this.processer.speed.end!=null && this.processer.speed.start!=null)time_used = (this.processer.speed.end-this.processer.speed.start)
		this.processer.speed.time_used=time_used;
		this.processer.uploadTotalBytes=this.processer.uploadFileBytes+bytesComplete;
		
		if(this.processer.bind!=null){
			var txt = (bytesComplete/bytesTotal)*100;
			txt = txt.toFixed(2);
			$_("b_" + file.id).style.width=txt+"%";
			$_("p_" + file.id).innerHTML=txt+"%";
			$_("sp_" + file.id).innerHTML=this.processer.speed.value;
			if(txt=="100.00"){
				$_("a_"+file.id).innerHTML="无";
				$_("p_" + file.id).innerHTML=" <img src=\"images/loading.gif\" width=\"16\" height=\"16\" />";
			}
		}
		if (typeof this.processer.user_upload_progress_handler === "function") return this.processer.user_upload_progress_handler.call(this, file, bytesComplete, bytesTotal);
	};
	
	SWFUpload.handler.uploadSuccessHandler = function (file, serverData) {
		this.processer.speed.end=new Date();
		var time_used=0;
		if(this.processer.speed.end!=null && this.processer.speed.start!=null)time_used = (this.processer.speed.end-this.processer.speed.start)
		if(this.processer.bind!=null && $_("sp_" + file.id)!=null)$_("sp_" + file.id).innerHTML=timeString(parseInt(time_used)/1000,true);	
		
		var File = null;
		try{
		eval("File = (" + serverData + ");");
		}catch(ex){}
		if(!File)return"异常";
		if(File.err){
			this.SetFileStatus(file.id,SWFUpload.FILE_STATUS.ERROR);
			this.uploadError(file,500,File.msg);
			return "";
		}
		this.processer.uploadFileBytes+=file.size;
		file.newname = File.name;
		Files.push(file);
		if(this.processer.bind!=null){
			$_("a_"+file.id).innerHTML="无";
			$_("b_"+file.id).style.backgroundColor="#f6f6f6";
			$_(file.id).style.border="1px solid #ddd";
			$_("p_" + file.id).innerHTML=" <img src=\"images/right.png\" width=\"16\" height=\"16\" />";
		}
		if(typeof UploadSucceedCallBack)UploadSucceedCallBack.apply(this,[file,File]);
		if (typeof this.processer.user_upload_success_handler === "function") return this.processer.user_upload_success_handler.call(this, file, serverData);
	};
	
	SWFUpload.handler.uploadErrorHandler = function (file, errorCode, message,serverdata,args) {
		if(errorCode==SWFUpload.UPLOAD_ERROR.FILE_CANCELLED){
			if(args!=null && args==true){
				$_(file.id).parentNode.removeChild($_(file.id));
			}else{
				if(this.processer.bind!=null){
					$_("p_"+file.id).innerHTML="已取消";
					$_("b_"+file.id).style.width=0;
					$_("a_"+file.id).innerHTML="<a href=\"javascript:void(0)\" onclick=\"SWFUpload.instances['" + this.movieName + "'].requeueUpload('" + file.id + "');SWFUpload.instances['" + this.movieName + "'].startUploadFiles('" + file.id + "',true);\">上传</a>"
					+" <a href=\"javascript:void(0)\" onclick=\"$_('" + file.id + "').parentNode.removeChild($_('" + file.id + "'));\">移除</a>";
				}
			}
		}else if(errorCode==SWFUpload.UPLOAD_ERROR.UPLOAD_STOPPED){
			SWFUpload.handler.stoped=true;
		}else{
			if(this.processer.bind!=null){
				var m = encodeURIComponent(message);
				$_("a_"+file.id).innerHTML="<a href=\"javascript:void(0)\" onclick=\"alert(decodeURIComponent('" + m + "'));\">信息</a> <a href=\"javascript:void(0)\" onclick=\"swfu.requeueUpload('" + file.id + "');swfu.startUploadFiles('" + file.id + "')\">重传</a>";
				$_("b_"+file.id).style.backgroundColor="#fee";
				$_(file.id).style.border="1px solid #fcc";
				$_("p_" + file.id).innerHTML=" <img src=\"images/wrong.png\" width=\"16\" height=\"16\" />";
			}
		}
		if (typeof this.processer.user_upload_error_handler === "function") return this.processer.user_upload_error_handler.call(this, file, errorCode, message);
	};
	
	SWFUpload.handler.uploadCompleteHandler = function (file) {
		if(this.Status().queued>0 && !SWFUpload.handler.stoped && !SWFUpload.handler.onlyOne){
			this.startUpload();
		}else{
			this.setButtonDisabled(false);
			SWFUpload.handler.onlyOne = false;
			if(Files.length>0)$_("message").innerHTML = ("成功上传" + Files.length + "个文件。");
		}
		if (typeof this.processer.user_upload_complete_handler === "function") return this.processer.user_upload_complete_handler.call(this, file);
	};
	
	SWFUpload.handler.fileQueueStart = function (length) {
		if(this.processer.bind==null)return;
		var fl = this.processer.bind;
		while(fl.childNodes.length>1){fl.removeChild(fl.lastChild);}
		if (typeof this.processer.file_queue_start_handler === "function") return this.processer.file_queue_start_handler.call(this, length);
	};
}
function HandlerInit(Setting){
	var set_={
		flash_url : "scripts/SWFUpload.swf",
		post_params: {},
		file_queue_limit:0,
		custom_settings : {},
		debug: false
	};
	for(var i in Setting){
		set_[i] = 	Setting[i];
	}
	return new SWFUpload(set_);
}