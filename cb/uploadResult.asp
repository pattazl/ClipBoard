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
' ��ȡ·��
const rowSplit = "|"
dim ufp,path
path= "./temp/"
function myConvert(strIn)
	' ���� ADODB.Stream ����
	Set stream = Server.CreateObject("ADODB.Stream")

	' ����������Ϊ�ı�����
	stream.Type = 2 

	' �����ַ���Ϊ UTF-8
	stream.Charset = "gb2312"
	' ����
	stream.Open
	' ���������ַ���д����
	stream.WriteText strIn

	' ����ת��Ϊ ANSI ����
	stream.Position = 0

	stream.Charset = "utf-8"
	strOut = stream.ReadText
	' �ر���
	stream.Close

	' �ͷ�������
	Set stream = Nothing
	' ������
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
		' �ļ���һֱ <size> ���
		fileArr(count) =  ufp & "<"& file.fileSize& ">"
		
		count = count + 1
	else
		Response.Write("�ļ���׺" & upfileext & "������Ҫ��")
		Response.End
	end if
	' �����¼�����ݿ�
	' Application("tempText")
	' file.fileSize&",'"&path&ufp&"','"&session("username")&"','"&upfileext
next
Application(tick) = Join(fileArr,rowSplit)
Set upload=nothing
Response.Write("�Ѿ��ɹ��ϴ�")

if ufp<>"" then %>
<script language="JavaScript">
window.location.href="../cb.asp?tick=<%=tick%>&expireFile=<%=expireFile%>"
</script>
<%
end if
%>

</body></html>