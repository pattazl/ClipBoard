<%@LANGUAGE="JSCRIPT" CODEPAGE="65001"%>
<%
// 处理 Application 的数据，返回有效的 arr 数据
var fileFlag = 'F_Kp8~@k::'
var filePath = 'cb/temp/' // 临时文件路径
var fileReg = /(.*)<(\d+)>$/

function strItem(item){
	var str = String(item)
	if( str =='undefined' )
	{
		str = ''
	}
	return str
}
/// <summary>
/// 格式化文件大小的JS方法
/// </summary>
/// <param name="filesize">文件的大小,传入的是一个bytes为单位的参数</param>
/// <returns>格式化后的值</returns>
function renderSize(filesize){
  if(null==filesize||filesize==''){
      return "0 B";
  }
  var unitArr = new Array("B","KB","MB","GB","TB","PB","EB","ZB","YB");
  var index=0;
  var srcsize = parseFloat(filesize);
  index=Math.floor(Math.log(srcsize)/Math.log(1024));
  var size =srcsize/Math.pow(1024,index);
  size=size.toFixed(2);//保留的小数位数
  return size+unitArr[index];
}
function doApplication(){
   var maxCount = 50
   var rowSplit = '|-6VoPMA-|'
   var colSplit = '|-9YimUN-|'

   var nowTime = new Date().getTime()
   var arr = []

   // 读取原先数据，进行初步清理和赋值
   var app = Application("tempText")
   if(  app!= null && app != ""){
      var arr1 = app.split(rowSplit)
      for(var i=0; i < arr1.length; i++)
      {
         var arr2 = arr1[i].split(colSplit)
         if(arr2.length!=3){
            continue
         }
         var obj = {
            "txt": arr2[0],  // 具体内容
            "t": arr2[1],    // 创建时间
            "e": arr2[2]     // 失效时间
         }
         if(obj.e !=0 && obj.e < nowTime){
            // 数据无效，或 过失效时间，跳过
            continue
         }
         arr.unshift(obj)
      }
   }
   // 插入新数据
   var myText =''
   var clearAllFlag = ''
   var e = 0
   // 文件名
   var tick = strItem(Request.QueryString("tick"))
   var expireFile = strItem(Request.QueryString("expireFile"))
   if( /^\d+$/.test(tick) && strItem(Application(tick)) !='' ){
   // 有文件
	myText = fileFlag + Application(tick)
	e = expireFile
	Application(tick) = ''
   }
   if(Request.Form.Count>0){
      myText = strItem(Request.Form("myText"))
      e = strItem(Request.Form("expire"))
	  clearAllFlag  = strItem(Request.Form("clearAllFlag"))
	  if( clearAllFlag == '0oO1iIlLq9g' )
	  {
		// 全部清理
		Application("tempText") = ''
		arr = []
		cleanFile(arr)
		return arr
	  }
	}
	if(myText != ''){
	  if( e!=0 )
	  {
		 e = nowTime+e*24*3600*1000
	  }
	  obj = {
		 "txt": myText,       // 具体内容
		 "t": nowTime,    // 创建时间
		 "e": e           // 失效时间
	  }
	  arr.unshift(obj)
	}

   // 再次清理数据
   if (arr.length > maxCount) {
      arr.pop()
   }
   // 保存数据
   var tmpArr = []
   var fileArr = []
   for(i=0; i < arr.length; i++)
   {
      var o = arr[i]
      var str = o["txt"]+colSplit+o["t"]+colSplit+o["e"]
      tmpArr.unshift(str)
	  
	  if( o["txt"].indexOf(fileFlag) == 0 )
	  {
	  
		var file = getFileInfo(o["txt"])
		for( var j = 0;j<file.length;j++)
		{
			fileArr.push(file[j].name)
		}
	  }
   }
   Application("tempText") = tmpArr.join(rowSplit)
   cleanFile(fileArr)
   return arr
}

function getFileInfo(fileTxt){
	var fileInfo = fileTxt.replace(fileFlag,'').split('|')
	var arr = []
	for( var i = 0;i<fileInfo.length;i++)
	{
		var file = {}
		// filename<size> 
		var info = fileReg.exec(fileInfo[i])
		if(info!=null && info.length == 3){
			file.name = info[1]
			file.size = info[2]
			arr.push(file)
		}
	}
	return arr
}
function cleanFile(fileArr){
// loop file under filePath
	var fso, f, f1, fc, s;
	fso = new ActiveXObject("Scripting.FileSystemObject");
	f = fso.GetFolder(Server.MapPath(filePath));
	fc = new Enumerator(f.files);
	for (; !fc.atEnd(); fc.moveNext())
	{
		var file = fc.item()
		var canDel = true
		for(var i= 0;i<fileArr.length;i++)
		{
			if( fileArr[i] == file.Name)
			{
				canDel = false
				break
			}
		}
		if(canDel){
			file.Delete()
		}
	}
}
var arr = doApplication()
%>
<!DOCTYPE html>
<html lang="zh">

<head>
   <title> Temp Text </title>
   <meta name="viewport" content="width=device-width, initial-scale=1">
   <meta name="Author" content="Austin">
   <meta charset="utf-8">
</head>
<style>
   li div {
      font-size: small;
   }
   li {
      font-size: large
   }
   form,input{
	padding:2px
   }
</style>

<body>
<script>
	function getById(name){
		return document.getElementById(name)
	}
   function checkit() {
      var v = getById('myText').value
	  if(getById('clearAllFlag').value!=''){
		return true; // 清空数据
	  }
      return v.trim() != ''
   }
   function checkit2() {
      var v2 = document.querySelectorAll(".uploadFile")
	  var files = []
	  let count = 0;
	  for(let f of v2){
		count++;
		if( f.files.length == 0)
		{
			alert("第"+count+"个文件未选择!")
			f.focus()
			return false
		}
		files.push(encodeURI(f.files[0].name))
	  }
		var radioButtons = document.getElementsByName('expire');
		var selectedValue = 0
		// 遍历单选按钮
		for (var i = 0; i < radioButtons.length; i++) {
		  if (radioButtons[i].checked) {
			// 获取被选中的单选按钮的值
			selectedValue = radioButtons[i].value;
			break;
		  }
		}
	  getById('expireFile').value = selectedValue
	  getById('fileName').value = files.join('|')
      return true
   }
   let outFirstList = '', inFirstList = '',fileNum = 1;
   window.onload = function()
   {
	  outFirstList = getById('firstFile').outerHTML
	  inFirstList = getById('firstFile').innerHTML
      let items = document.querySelectorAll("li div")
      for(let k of items)
      {
         let d = k.attributes['data'].value.split(',')
         k.innerHTML="➕:"+ new Date(parseInt(d[0])).toLocaleString()+" ➖:"+new Date(parseInt(d[1])).toLocaleString()
      }
   }
   function addFile(obj,flag){
	let parent = obj.parentElement
	let divEl = document.createElement("div");
	divEl.innerHTML = inFirstList.replace('name="f_','name="f_'+fileNum) // change name
	fileNum++;
   // 添加
	if(flag>0){
		parent.parentElement.insertBefore(divEl,parent.nextSibling)
	}else
	{
		if(parent.id!="firstFile"){
			parent.parentElement.removeChild(parent);
		}
	}
   }
   function resetFile(){
	getById('fileList').innerHTML = outFirstList
	fileNum = 0
	return true;
   }
   function clearAll(){
	if(confirm('确定清除所有历史记录?')){
		getById('clearAllFlag').value ='0oO1iIlLq9g';
		return true;
	}
	return false;
   }
</script>
   <form name="form3" method="post" action="cb/uploadResult.asp" enctype="multipart/form-data" onSubmit="return checkit2();">
   <input type="hidden" id="expireFile" name="expireFile" value="0" />
   <input type="hidden" id="fileName" name="fileName" value="" />
    <div id="fileList">
      <div id="firstFile"><input class="uploadFile" name="f_" type="file" /><input type="button" onclick="addFile(this,1)" value="➕"> <input type="button" onclick="addFile(this,-1)" value="➖">
	  </div>
	</div>
	  <input type="submit" value="上传"  /> <input type="reset" onclick="return resetFile()" value="重置"  />
	</form>
   <form onsubmit="return checkit()" method="post">
      <textarea id="myText" name="myText" cols=50 rows=4 placeholder="复制相关数据后提交"></textarea>
      <label><input type="radio" name="expire" value="1" checked="checked">一天失效</input></label> 
      <label><input type="radio" name="expire" value="7" >一周失效</input></label> 
      <label><input type="radio" name="expire" value="0" >直到重启</input></label> 
      <br />
      <input type="submit" value="提交" /> <input type="submit" onclick="return clearAll()" value="清空全部记录" />
	  <input type="hidden" id="clearAllFlag" name="clearAllFlag" value="" />
   </form>
   <ol>
      <%
      for (var i = 0; i < arr.length; i++) {
         var obj = arr[i]
		 Response.write("<li><div data='" + obj.t +","+obj.e+"'></div></li>")
		 if( obj.txt.indexOf(fileFlag)== 0 ){
			var fileInfo = getFileInfo(obj.txt)
			for( var j = 0;j<fileInfo.length;j++)
			{
				Response.write("<a href='"+filePath+fileInfo[j].name+"' target='blank'>" +fileInfo[j].name+ "</a> "+ renderSize(fileInfo[j].size) +"<br/>")
			}
		 }else{
			Response.write("<textarea readonly='1' cols=50 rows=4>" + obj.txt + "</textarea>")
		 }
		 Response.write("</li>")
      }
      %>
   </ol>

</body>

</html>