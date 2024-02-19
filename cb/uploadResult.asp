<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!-- #include file="Upload.asp" -->
<html>
<head>
   <title> Temp Text </title>
   <meta name="viewport" content="width=device-width, initial-scale=1">
   <meta name="Author" content="Austin">
</head>
<body>
<%
' 获取路径
const rowSplit = "|"
dim ufp,path
path= "./temp/"
function myConvert(strIn)
	' 创建 ADODB.Stream 对象
	Set stream = Server.CreateObject("ADODB.Stream")

	' 设置流类型为文本类型
	stream.Type = 2 

	' 设置字符集为 UTF-8
	stream.Charset = "gb2312"
	' 打开流
	stream.Open
	' 将编码后的字符串写入流
	stream.WriteText strIn

	' 将流转换为 ANSI 编码
	stream.Position = 0

	stream.Charset = "utf-8"
	strOut = stream.ReadText
	' 关闭流
	stream.Close

	' 释放流对象
	Set stream = Nothing
	' 输出结果
	myConvert = strOut
end function



Server.ScriptTimeout = 900

set Upload = new DoteyUpload

Upload.Upload()

if Upload.ErrMsg <> "" then 
Response.Write(Upload.ErrMsg)
Response.End()
end if

if Upload.Files.Count > 0 then
Items = Upload.Files.Items
end if

fileArr = Array()
ReDim fileArr(Upload.Files.Count-1)

expireFile = Upload.Form("expireFile")
count = 0
tick = day(now)&hour(now)&minute(now)&second(now)
for each File in Upload.Files.Items
	upfilename = split(File.FileName,".")
	upfileext = Lcase(upfilename(ubound(upfilename)))
 ' 
	if InStr("|ini|md|mp3|bmp|jpg|jpeg|png|zip|7z|rar|pdf|doc|docx|xls|xlsx|ppt|pptx","|"&upfileext) > 0  then
		ufp= tick & "-" & myConvert(File.FileName )
		file.saveas Server.mappath(path&ufp)
		' 文件中一直 <size> 标记
		fileArr(count) =  ufp & "<"& file.fileSize& ">"
		
		count = count + 1
	else
		Response.Write("文件后缀" & upfileext & "不符合要求")
		Response.End
	end if
	' 保存记录到数据库
	' Application("tempText")
	' file.fileSize&",'"&path&ufp&"','"&session("username")&"','"&upfileext
next
Application(tick) = Join(fileArr,rowSplit)
Set upload=nothing
Response.Write("已经成功上传")

if ufp<>"" then %>
<script language="JavaScript">
window.location.href="../cb.asp?tick=<%=tick%>&expireFile=<%=expireFile%>"
</script>
<%
end if
%>

</body></html>