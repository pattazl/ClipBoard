<%@LANGUAGE="JSCRIPT" CODEPAGE="65001"%>
<%
var fileInfo = ''
// 处理 Application 的数据，返回有效的 arr 数据
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
   if(Request.Form.Count>0){
      var myText = Request.Form("myText")
      var e = Request.Form("expire")
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
   // 文件名
   var tick = Request.QueryString("tick")
   if( tick　!='' && Application(tick) !=''){
   // 有文件
	fileInfo = Application(tick)
	Application(tick) = ''
   }

   // 再次清理数据
   if (arr.length > maxCount) {
      arr.pop()
   }
   // 保存数据
   var tmpArr = []
   for(i=0; i < arr.length; i++)
   {
      var o = arr[i]
      var str = o["txt"]+colSplit+o["t"]+colSplit+o["e"]
      tmpArr.unshift(str)
   }
   Application("tempText") = tmpArr.join(rowSplit)
   return arr
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
   function checkit() {
      var v = document.getElementById('myText').value
      return v.trim() != ''
   }
   function checkit2() {
      var v2 = document.querySelectorAll(".uploadFile")
	  let count = 0;
	  for(let f of v2){
		count++;
		if( f.value.trim() == '')
		{
			alert("第"+count+"个文件未选择!")
			f.focus()
			return false
		}
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
	  document.getElementById('expireFile').value = selectedValue
      return true
   }
   let outFirstList = '', inFirstList = '',fileNum = 1;
   window.onload = function()
   {
	  outFirstList = document.getElementById('firstFile').outerHTML
	  inFirstList = document.getElementById('firstFile').innerHTML
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
	document.getElementById('fileList').innerHTML = outFirstList
	fileNum = 0
	return true;
   }
</script>
   <form name="form3" method="post" action="cb/uploadResult.asp" enctype="multipart/form-data" onSubmit="return checkit2();">
   <input type="hidden" id="expireFile" name="expireFile" value="0" />
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
      <input type="submit" value="提交" />
   </form>
   <div>-
      <%
	  Response.write( fileInfo )
%>-
</div>
   <ol>
      <%
      for (var i = 0; i < arr.length; i++) {
         var obj = arr[i]
         Response.write("<li><div data='" + obj.t +","+obj.e+"'></div><textarea readonly='1' cols=50 rows=4>" + obj.txt + "</textarea></li>")
      }
      %>
   </ol>

</body>

</html>