<% ' Insert.asp 
%>
<!--#include file="Loader.asp"-->
<%
Dim connStr
      dim conn1 
      Dim rs
    ' Checking to make sure if file was uploaded
    
    
      ' Connection string
      
      'connStr = "DRIVER=Microsoft Access Driver (*.mdb);DBQ="
      'connStr = connStr & Server.MapPath("/andy26/database/missyou.mdb")
      set conn1 = server.CreateObject("Adodb.connection")
      
         connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
      Server.MapPath("missyou.mdb")
    conn1.Open connStr

'------------------
  Response.Buffer = True
  ' load object
  Dim load
    Set load = new Loader
    
    ' calling initialize method
    load.initialize
    
  ' File binary data
  Dim Photo1
    Photo1 = load.getFileData("Photo1")
    'Session("Photo1")=Photo1
  ' File name
  Dim fileName
    fileName = LCase(load.getFileName("Photo1"))
  ' File path
  Dim filePath
    filePath = load.getFilePath("Photo1")
  ' File path complete
  Dim filePathComplete
    filePathComplete = load.getFilePathComplete("Photo1")
  ' File size
  Dim fileSize
    fileSize = load.getFileSize("Photo1")
  ' File size translated
  Dim fileSizeTranslated
    fileSizeTranslated = load.getFileSizeTranslated("Photo1")
    'Response.Write fileSizeTranslated
			'if fileSizeTranslated<>"" then
			'if cint(left(fileSizeTranslated,instr(fileSizeTranslated,",")-1))>50 then 
			'Response.Redirect("reg.asp?wrong_picture=1")
			'end if
			'end if
  ' Content Type
  Dim contentType
    contentType = load.getContentType("Photo1")
  ' No. of Form elements
  Dim countElements
    countElements = load.Count
  ' Value of text input field "fname"
  Dim fnameInput
    fnameInput = load.getValue("fname")
    Session("fnameInput")=fnameInput
  ' Value of text input field "lname"
  Dim lnameInput
    lnameInput = load.getValue("lname")
    Session("lnameInput")=lnameInput
  ' Value of text input field "profession"
  Dim profession
    profession = load.getValue("profession")  
    Session("profession")=profession
  Dim City
  City=load.getValue("City")  
  Session("City")=City
  Dim PersonName
  PersonName=load.getValue("PersonName")
  Session("PersonName")=PersonName
  Dim Nick
  Nick=load.getValue("Nick")
  Session("Nick")=Nick
  Dim Password
  Password=load.getValue("pass1")
  Session("Password")=Password
  Dim Sex
  Sex=load.getValue("Sex")
  Session("Sex")=Sex
  Dim Age
  dim day
  dim month
  dim monthname
  dim year
  if cint(load.getValue("day"))<10 then  
  day="0" & load.getValue("day")
  else
  day=load.getValue("day")
  end if
  
  month=load.getValue("month")
  select case month
  case "1"
  monthname="январь"
  case "2"
  monthname="февраль"
  case "3"
  monthname="март"
  case "4"
  monthname="апрель"
  case "5"
  monthname="май"
  case "6"
  monthname="июнь"
  case "7"
  monthname="июль"
  case "8"
  monthname="август"
  case "9"
  monthname="сентябрь"
  case "10"
  monthname="октябрь"
  case "11"
  monthname="ноябрь"
  case "12"
  monthname="декабрь"
  end select
  year=load.getValue("year")
  
  Age= day & "." & month & "." & year
  
    Session("year")=right(Age,4)
	Session("month")=month
	Session("monthname")=monthname
	Session("day")=left(Age,2)
    
  Session("Age")=Age
  Dim Weight
  Weight=load.getValue("weight")
  Session("Weight")=Weight
  Dim Height
  Height=load.getValue("height")
  Session("Height")=Height
  Dim aboutme
  aboutme=load.getValue("aboutme")
  Session("aboutme")=aboutme
  Dim aboutyou
  aboutyou=load.getValue("aboutyou")
  Session("aboutyou")=aboutyou
  Dim phone
  phone=load.getValue("telephone")
  Session("phone")=phone
  Dim html_page
  html_page=load.getValue("html_page")
  Session("html_page")=html_page
  'Dim PersonID
  'html_page=load.getValue("PersonID")
  Dim Login
  Login=load.getValue("Login")
  Session("Login")=Login
  
  dim find_sex
  find_sex=load.getValue("find_sex")
  select case find_sex
  case "f"
  Session("find_sex")="девушку"
  case "m"
  Session("find_sex")="парня"

  end select
  
  dim pattern
  pattern=load.getValue("pattern")
  Session("pattern")=pattern
  dim dream_life
  dream_life=load.getValue("dream_life")
  Session("dream_life")=dream_life
  dim living
  living=load.getValue("living")
  Session("living")=living
  dim purpose_id
  purpose_id=load.getValue("purpose_id")
  select case purpose_id
  case "1"
  Session("purpose_id")="общения по интернету"
  case "2"
  Session("purpose_id")="серьезных отношений"
  case "3"
  Session("purpose_id")="дружбы"
  case "4"
  Session("purpose_id")="страстной любви"
  case "5"
  Session("purpose_id")="интимных отношений"
  case "6"
  Session("purpose_id")="путешествия"
  end select
  
  dim favorite_eat
  favorite_eat=load.getValue("favorite_eat")
  Session("favorite_eat")=favorite_eat
  dim favorite_drink
  favorite_drink=load.getValue("favorite_drink")
  Session("favorite_drink")=favorite_drink
  dim interest
  interest=load.getValue("interest")
  Session("interest")=interest
  
  
  
  if fileSizeTranslated<>"" then
    if cint(left(fileSizeTranslated,instr(fileSizeTranslated,",")-1))>50 then 
    Response.Redirect("reg.asp?wrong_picture=1")
    end if
    end if
    
  ' destroying load object
  Set load = Nothing
  
  
  
  
  
%>

<html>
<head>
  <title>Введите свою анкету</title>
 <link rel="stylesheet" type="text/css" href="index.css">
 <META content="text/html; charset=windows-1251" http-equiv=Content-Type> 
<META NAME="Keywords" CONTENT="любовь, отношения, брак, флирт, девушки,эротика, брак, анекдоты, развлечения">

</head>
<body background="bg.jpg" alink="#0000ff" vlink="#0000ff" link="#0000ff">
  <p align="center">
    <b></b><br>
    <a href="show.asp">Посмотреть данные</a>
  </p>
 <!--- 
  <center>
<table>
	<tr>
	 <td><FONT size="2"  color="#0000ff" style="FONT-WEIGHT: bold">&#1048;&#1084;&#1103;:</font></td><td><input name=PersonName 

type=text></td>
	</tr>
	<tr> 
	 <td><FONT size="2"  color="#0000ff" style="FONT-WEIGHT: bold">&#1053;&#1080;&#1082;:</font> </td><td><input name=Nick type=text></td>
	</tr> 
	<tr>
	 <td><FONT size="2" color="#0000ff" style="FONT-WEIGHT: bold">&#1043;&#1086;&#1088;&#1086;&#1076;:</font></td><td>
	 <select name="City"><option value="1" selected>&#1052;&#1086;&#1089;&#1082;&#1074;&#1072;</option>
						<option value="2" >&#1055;&#1080;&#1090;&#1077;&#1088;</option>
						<option value="3" >&#1053;&#1086;&#1074;&#1086;&#1089;&#1080;&#1073;&#1080;&#1088;&#1089;&#1082;</option>
						<option value="4" >&#1050;&#1088;&#1072;&#1089;&#1085;&#1086;&#1076;&#1072;&#1088;</option>
						<option value="5" >&#1042;&#1083;&#1072;&#1076;&#1080;&#1074;&#1086;&#1089;&#1090;&#1086;&#1082;</option>
	 </select></td>
     </tr> 
     <tr>
	 <td><FONT size="2" color="#0000ff" style="FONT-WEIGHT: bold">Пол:</font></td><td>
	 <select name="Sex"><option value="f" selected>Женcкий</option>
						<option value="m">Мужской</option>
						
	 </select></td>
     </tr> 
     <tr>
	 <td><FONT size="2" color="#0000ff" style="FONT-WEIGHT: bold">&#1042;&#1086;&#1079;&#1088;&#1072;&#1089;&#1090;:</font> </td><td><input name=age 

type=text></td>
	 </tr>
	 <tr>
	 <td><FONT size="2" color="#0000ff" style="FONT-WEIGHT: bold">&#1042;&#1077;&#1089;:</font></td><td><input name=weight type=text></td>
	 </tr>
	 <tr>
	 <td><FONT size="2" color="#0000ff" style="FONT-WEIGHT: bold">&#1056;&#1086;&#1089;&#1090;:</font></td><td><input name=height type=text></td>
	 </tr>
	 <tr>
	 <td><FONT size="2" color="#0000ff" style="FONT-WEIGHT: bold">&#1054; &#1089;&#1077;&#1073;&#1077;:</font></td><td><input name=aboutme 

type=text></td>
	 </tr>
	 <tr>
	 <td><FONT size="2" color="#0000ff" style="FONT-WEIGHT: bold">&#1054; &#1090;&#1086;&#1084;, &#1082;&#1086;&#1075;&#1086; &#1080;&#1097;&#1077;&#1090;&#1077;:</font></td><td><input name=aboutyou

type=text></td>
	 </tr>
	 <tr>
	 <td><FONT size="2" color="#0000ff" style="FONT-WEIGHT: bold">&#1058;&#1077;&#1083;&#1077;&#1092;&#1086;&#1085;:</font></td><td><input name=telephone 

type=text></td>
	 </tr>
	 <tr>
	 <td><FONT size="2" color="#0000ff" style="FONT-WEIGHT: bold">HTML-&#1089;&#1090;&#1088;&#1072;&#1085;&#1080;&#1094;&#1072;:</html></td><td><input name=html_page 

type=text></td>
	 </tr>
	 <td><FONT size="2" color="#0000ff" style="FONT-WEIGHT: bold">Фото :</td></td><td>
		<input type="file" name="Photo1" size="40"></td></tr>
	<td> </td><td>
	 <tr>
	 <td><input type="submit" value="&#1044;&#1086;&#1073;&#1072;&#1074;&#1080;&#1090;&#1100;" id=submit1 name=submit1></td><td><input type="reset" value="&#1054;&#1095;&#1080;&#1089;&#1090;&#1080;&#1090;&#1100;" id=reset1 name=reset1></td>
   
	 </tr>



	
</table>
</center>
  ---->
  <p style="padding-left:220;">
  <%= fileName %> получены данные ...<br>
  <% 
      If fileSize > 0 Then
      ' Recordset object
        
        Set rs = Server.CreateObject("ADODB.Recordset")
        
        rs.Open "Persons", connStr, 2, 2
        
        ' Adding data
        rs.AddNew
          rs("City") = City
          rs("PersonName") = PersonName
          rs("Photo1").AppendChunk Photo1
          rs("Nick") = Nick
          rs("Sex") = Sex
          rs("Age") = Age
          rs("Weight") = weight
          rs("Height") = height
          rs("aboutme") = aboutme
          rs("aboutyou") = aboutyou
          rs("phone") = phone
          rs("html_page") = html_page
          
          
          rs("pattern") = pattern
          rs("dream_life") = dream_life
          rs("living") = living
          rs("find_sex") = find_sex
          rs("purpose_id") = purpose_id
          
          rs("favorite_eat") = favorite_eat
          rs("favorite_drink") = favorite_drink
          rs("interest") = interest
          rs("purpose_id") = purpose_id
          if (Sex="m") and (find_sex="m") then rs("rubric_id") = 2
          if (Sex="m") and (find_sex="f") then rs("rubric_id") = 1
          if (Sex="f") and (find_sex="f") then rs("rubric_id") = 12
          if (Sex="f") and (find_sex="m") then rs("rubric_id") = 11
          
          
	
          
        rs.Update
        Set rs1 = Server.CreateObject("ADODB.Recordset")
        
        rs1.Open "Logins", connStr, 2, 2
        rs1.AddNew 
        rs1("PersonID")= rs("ID")
        rs1("Login")= rs("Nick")
        rs1("Password")= Password
       rs1.Update
       rs1.close
        
        rs.Close
       set rs = nothing 
      Response.Write "<font color=""green"">Ваша анкета была успешно загружена..."
      Response.Write "</font>"
      
      
        'set rs0 = server.CreateObject("Adodb.recordset")
        'set rs0 = conn1.execute("select ID from persons where Nick='"&Nick&"'")
        'rs0.MoveFirst
        'set rs1 = server.CreateObject("Adodb.recordset")
        'sql="insert into logins (PersonID,Login,password) values("&rs0(0)&","&Nick&","&pass1&")"
        'Response.Write sql
        'set rs1 = conn1.execute(sql)
        
        'rs1.Close
        'Set rs1 = Nothing 
  
      
      
      
    Else
    
      ' Connection string
      
      'connStr = "DRIVER=Microsoft Access Driver (*.mdb);DBQ="
      'connStr = connStr & Server.MapPath("/andy26/database/missyou.mdb")
      set conn1 = server.CreateObject("Adodb.connection")
      
        connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
         Server.MapPath("missyou.mdb")
    conn1.Open connStr
      ' Recordset object
      
        Set rs = Server.CreateObject("ADODB.Recordset")
        
        rs.Open "Persons", connStr, 2, 2
        
        ' Adding data
        rs.AddNew
          rs("City") = City
          rs("PersonName") = PersonName
          'rs("Photo1").AppendChunk Photo1
          rs("Nick") = Nick
          rs("Sex") = Sex
          rs("Age") = age
          rs("Weight") = weight
          rs("Height") = height
          rs("aboutme") = aboutme
          rs("aboutyou") = aboutyou
          rs("phone") = phone
          rs("html_page") = html_page
         
        rs.Update
        Set rs1 = Server.CreateObject("ADODB.Recordset")
        
        rs1.Open "Logins", connStr, 2, 2
        rs1.AddNew 
        rs1("PersonID")= rs("ID")
        rs1("Login")= rs("Nick")
        rs1("Password")= Password
       rs1.Update
        rs1.close
        rs.Close
        Set rs = Nothing
        
      Response.Write "<font color=""brown"">Вы не выбрали фото для загрузки, наверное можно потом как-нибудь:) "
      Response.Write "...</font>"
    End If
      
      
    If Err.number <> 0 Then
      Response.Write "<br><font color=""red"">Ой, какие-то ошибки при загрузке, попробуйте еще раз:(..."
      Response.Write "</font>"
    End If
    
  %>
  </p>
  
 <center><FONT size="1" color="#0000ff" style="FONT-WEIGHT: bold">&copy;2005, Solva SoftWare inc. All rights Reserved 

</font></center>

</body>
</html>
