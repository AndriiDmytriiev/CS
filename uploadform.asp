<%@ Language=VBScript %>
<%Option explicit

Dim strMessage, strFolder
Dim httpref, lngFileSize
Dim strExcludes, strIncludes

	'-----------------------------------------------
	'This is the complete upload file program.
	'This is intended to upload graphics onto the web and
	'to delete them if required.
	'Set up the configurations below to define which
	'directory to use etc, then set the permissions on
	'the directory to 'Change' i.e. Read/Write
	'-----------------------------------------------

	%>
	<!-- #include file = "config.asp" -->	
	<%
	
	strMessage = Request.QueryString ("msg")
	
'--------------------------------------------
Sub main()

	%>
	<html>
	<head>
	<title>Загрузка файлов</title>
	<link rel="stylesheet" href="index.css">
	</head>
	<body>
	<table width=100% ID="Table2">
<tr>
<td align=left valign="top">
<a href="http://lovingme.ru/tiptop/common.asp?ID=<%=Session("ID")%>"><font  size=1>Общие данные(редактировать)</font></a>
		<hr>
		<a href="http://lovingme.ru/tiptop/vnutr.asp?ID=<%=Session("ID")%>"><font  size=1>ВНУТРЕННЯЯ УБОРКА ПОМЕЩЕНИЙ (редактировать)</font></a>
		<hr>



		<a href="http://lovingme.ru/tiptop/terr.asp?ID=<%=Session("ID")%>"><font  size=1>УБОРКА ТЕРРИТОРИИ (редактировать)</font></a>
		<hr>

		<a href="http://lovingme.ru/tiptop/spec.asp?ID=<%=Session("ID")%>"><font  size=1>ДОПОЛНИТЕЛЬНЫЕ УСЛУГИ (редактировать)</font></a>
		<hr>
		<a href="http://lovingme.ru/tiptop/default1_0_nonedit.asp?ID=<%=Session("ID")%>"><font  size=1>Заявка (1 PAGE)</font></a>
		
</td>	
	<%

	if Request.Form ("AskDelete") = "Delete" then	'ask if to delete
		call askDelete(Request.Form("fileId"))
	elseif Request.Form("delete") = "" then			'display at start up
		call displayform()
		call BuildFileList(strFolder)
	elseif Request.Form ("delete") = "Yes" then		'make deletion
		call delete(Request.form("fileId"))
		call displayForm()
		call BuildFileList(strFolder)
	elseif Request.Form ("delete") = "No" then		'do not make deletion
		call displayForm()
		call BuildFileList(strFolder)
	end if

	%>


	<%

end sub


'--------------------------------------------
'Displays the form to allow uploading
Sub displayForm()

Dim i, tempArray

	'Results box
	if strMessage <> "" then
	%>
	<td>
	<table border="1" align="center" cellspacing="0" cellpadding="2">
	
	<tr>
	<table ID="Table1">
  
 		<td class="text"><%=strMessage%></td>
	</tr>
	</table>
	<%
	end if


	%>
	
	<td>
	<table border="0" width="320" align="center" bgcolor="#0ff5f9" cellspacing="0" cellpadding="2">
	<tr>
	
	<td class="text">
		<%

		if lngFileSize > 0 then 
			Response.Write ("Максимальный размер файла = ") & cstr(cint(lngFileSize / 1024)+7)   & " Килобайт" & "<br>"
		end if
	
		if strExcludes <> "" then
			Response.Write("Тип файлов, который не может быть загружен = ") & "<br>"
			tempArray = Split(strExcludes,";")
			For i = 0 to UBOUND(tempArray)
				Response.Write (tempArray(i)) & " "
			Next
		end if

		if strIncludes <> "" then
			Response.Write("Формат файлов, которые могут быть загружены = ") & " "
			tempArray = Split(strIncludes,";")
			For i = 0 to UBOUND(tempArray)
				Response.Write (tempArray(i)) & " "
			Next
		end if
	
		%>	
		
		</td>
	</tr>
	</table>

	<form action="http://lovingme.ru/tiptop/uploadfile.asp" method="post" enctype="multipart/form-data">

		<table border="0" width="320" align="center" bgcolor="#0ff5f9" cellspacing="0" cellpadding="2">
		<tr>
			<td colspan="2" class="text">Выберите файл для загрузки</td>		
		</tr>
		<tr>
			<td class="text">
				<b>Файл: </b><input type="file" name="file1"><br>	
			</td>
		</tr>
		<tr>
			<td align="center">
				<input type="submit" value="Загрузить" name="submit">
			</td>
		</tr>
	</table>
		
	</form>

<%
end sub


'--------------------------------------------
'Builds a list of files on the directory
'INPUT : the folder to be used
Sub BuildFileList(strFolder)

    Dim oFS, oFolder, intNoOfFiles, FileName

    Set oFS = Server.CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFS.getFolder("/")
    %>
    <table border="0" width="320" align="center" bgcolor="#0ff5f9" cellspacing="0" cellpadding="2">
    <tr>
		<td class="text" colspan="2">Список файлов</tr>
    </td>
    <tr>
		<td class="text" colspan="2">&nbsp;</tr>
    </td>
    <tr>
		<td class="text"><b>Имя файла</b></td>
		
	</tr>
    </td>
    <%
	intNoOfFiles = 0

    For Each FileName in oFolder.Files	
		%>
		<tr>		
			<!--<form Name="frmDelete" method="post" action="requestsniffer.asp">-->
			<form Name="frmDelete" method="post" action="<%=Request.ServerVariables("PATH_INFO")%>">
				<td class="text">
					<a href="<%=httpref & "/" & mid(FileName.Path,instrrev(FileName.Path,"\")+1)%>" target="_blank"><%=mid(FileName.Path,instrrev(FileName.Path,"\")+1)%></a>
				</td>
				<td class="text">
					<input type="hidden" name="fileId" value="<%=mid(FileName.Path,instrrev(FileName.Path,"\")+1)%>">
					
				</td>
			</form>			
		</tr>
		<%
		intNoOfFiles = intNoOfFiles + 1
    Next
    
    Set oFolder = nothing
   
	if intNoOfFiles = 0 then
		%>
		<tr align="center">
			<td colspan="2" class="text">Нет доступных файлов</td>
		</tr>		
		<%
	end if
  
	%>
    </table>
    <%    
   
End Sub


'--------------------------------------------
'Ask if to delete this file
'INPUT : the file name to be deleted, less the path
Sub askDelete(strFileName)

	%>
	<html>
	<head>
	<title>Удалить файл да/нет?</title>
	<link rel="stylesheet" href="upload.css">
	</head>
	<body>
	
	<form name="frmConfirmDelete" method="post" action="<%=Request.ServerVariables("PATH_INFO")%>">
	<table border="0" align="center">
		<tr>
			<td class="text">
				Вы уверены, что хотите удалить <b><%=strFileName%></b> ?
			</td>
		</tr>
		<tr>
			<td class="text" align="center">
				<input type="hidden" name="fileId" value="<%=strFileName%>">
				<INPUT type="submit" value="Да" name="Delete">
				&nbsp;&nbsp;
				<INPUT type="submit" value="Нет" name="Delete">
			</td>
		</tr>
	</table>
	</form>

	</body>
	</html>
	<%

end sub

'--------------------------------------------
'Deletes the file given the full file name strFileName
'INPUT : the file name to be deleted, less the path
Sub delete(strFileName)

	'Response.write strFileName 
	'Response.End 

	Dim oFS, a

    Set oFS = Server.CreateObject("Scripting.FileSystemObject")
	a = oFS.DeleteFile(strFolder & "\" & strFileName)

	Set oFs = nothing
	Set a = nothing	
	
End sub


'--------------------------------------------
call main()

%>
</tr>
 </table>	
	</body>
	</html>
