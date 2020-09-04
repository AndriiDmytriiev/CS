<%
on error resume next
   ' -- show.asp --
   ' Generates a list of uploaded files
    Session("PersonID")=""
    'if Session("first")="" then
   'Session("first")=1
   'end if
   Response.Buffer = True
   
   ' Connection String
      Dim connStr
      'connStr = "DRIVER=Microsoft Access Driver (*.mdb);DBQ="
      'connStr = connStr & Server.MapPath("/andy26/database/missyou.mdb")

      connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
      Server.MapPath("missyou.mdb")


Sub SetPage(localRS, pag, lim, shift)
'		Response.Write "(pag-1)*lim - shift="&(pag-1)*lim - shift&"<br>"
'		Response.Write "localRS.RecordCount="&localRS.RecordCount&"<br>"
	If localRS.CursorType = adOpenKeyset Or localRS.CursorType = adOpenStatic Then
		If (pag-1)*lim+1 - shift > 0 And CInt(localRS.RecordCount)>=CInt((pag-1)*lim+1 - shift) Then
			localRS.AbsolutePosition = (pag-1)*lim+1 - shift
		ElseIf localRS.RecordCount<(pag-1)*lim+1 - shift Then
			If Not localRS.EOF Then localRS.MoveLast
			If Not localRS.EOF Then localRS.MoveNext 
		End If
	Else
		For i = 1 To (pag-1)*lim - shift
			If Not localRS.EOF Then localRS.MoveNext
		Next
	End If
	localRS.CacheSize = limit
End Sub
%>
<html>
<head>
        <link rel="stylesheet" type="text/css" href="index.css">
     <META content="text/html; charset=windows-1251" http-equiv=Content-Type> 
<META NAME="Keywords" CONTENT="любовь, отношения, брак, флирт, девушки,эротика, брак, анекдоты, развлечения">
<title>Окунитесь в роскошь человеческого общения</title>
   
</head>
<body background="bg.jpg" alink="#0000ff" vlink="#0000ff" link="#0000ff">
<p align=center>
   <center>
     <td align=center>
<img src="relationshipsromance.jpeg"></img>
</td><br>

<br>

      <a href="reg.asp">Зарегистрироваться можно здесь</a>
      <br>
   </center>
 </p>  
   <table width="80%"  align="center" color="#123123" border=0>
<%
   ' Recordset Object
   Dim rs
   dim sql
   dim flag
   flag=false
      Set rs = Server.CreateObject("ADODB.Recordset")
      
      ' opening connection
      sql="select [ID],[City],[PersonName],[Nick]," & _
	  	"[Sex],[Age],[Weight],[Height],[aboutme],[aboutyou],[phone],[email],[Photo1],[pattern],[dream_life],[living],[favorite_eat],[favorite_drink],[interest] from Persons order by [ID] desc" 
	  
	   rs.Open sql, connStr,3,4
  
      
        'Response.Write ("<font color=#0000ff")
	     Response.Write "<tr><td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Имя</font></b></td>"
	     'Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Ник</font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Пол</font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Возраст</font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Вес</font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Рост</font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>О себе</font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>О том, кого ищет</font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Характер</font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Мечта</font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Образ жизни</font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Любимая еда</font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Любимый напиток</font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Интересы</font></b></td>"

	     'Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Телефон</font></b></td>"
	     'Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>E-mail</font></b></td>"
         Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Фото</font></b></td>"
         Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Написать письмо</font></b></td>"
   howmuch=rs.RecordCount     
    'call SetPage (RS, 1, 10, 0)
			
			j=1
			Session("first")=1
			if request("page")=""  then 
			page=1
			else
			page=cint(request("page"))
			end if
			if request("page")="10000" or  (cint(request("page"))>10 and cint(request("page"))<=20) then
			    Session("first")=2   
				page=11
			end if
			if request("page")="20000" or  (cint(request("page"))>20 and cint(request("page"))<=30) then
			    Session("first")=3   
				page=21
			end if
			if request("page")="30000" or  (cint(request("page"))>30 and cint(request("page"))<=40) then
			    Session("first")=4   
				page=31
			end if
			
			if request("page")="40000" or  (cint(request("page"))>40 and cint(request("page"))<=50)then
			    Session("first")=5   
				page=41
			end if	         
			if request("page")="50000" or  (cint(request("page"))>50 and cint(request("page"))<=60)then
			    Session("first")=6   
				page=51
			end if	         
			if request("page")="60000" or  (cint(request("page"))>60 and cint(request("page"))<=70)then
			    Session("first")=7   
				page=61
			end if	         
			if request("page")="70000" or  (cint(request("page"))>70 and cint(request("page"))<=80)then
			    Session("first")=8   
				page=71
			end if	         
			if request("page")="80000" or  (cint(request("page"))>80 and cint(request("page"))<=90)then
			    Session("first")=9   
				page=81
			end if	
			
			
			
			
	        if page>1 then
			j=1
			While Not rs.EOF and j<(page-1)*10
                      rs.MoveNext
            j=j+1
            Wend
 			end if
 			
 			
         While Not rs.EOF and j<page*10
            if (j mod 2) = 1  then 
            Response.Write "<tr><td bgcolor=#eeeeee><font color=#685362 size=1 >"
            'Response.Write "<a href=""file.asp?ID=" & rs("ID") & """>"
            Response.Write rs("PersonName") & "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"
            'Response.Write rs("Nick")& "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"
            Response.Write rs("Sex")& "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            Response.Write cstr(datediff("yyyy",rs("Age"),Now()))& "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"
            Response.Write rs("Weight")& "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            Response.Write rs("Height")& "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"
            
            Response.Write rs("aboutme")& "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            Response.Write rs("aboutyou")& "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"

            Response.Write rs("pattern")& "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            Response.Write rs("dream_life")& "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"
            Response.Write rs("living")& "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            Response.Write rs("favorite_eat")& "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"
            Response.Write rs("favorite_drink")& "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            Response.Write rs("interest")& "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"

            'Response.Write rs("phone")& "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            'Response.Write rs("html_page")& "&nbsp;" & "</font></td><td bgcolor=#ffffff>"
            
            Response.Write "<a href=""javascript:window.open('file.asp?ID=" & rs("ID") & "','','toolbar=0,menubar=1,width=500,height=400,resizable=1,scrollbars=yes');location.href='showall.asp'" & """>"
            Response.Write "<font color=#685362 size=1>Фото</font></a>" & "</td><td bgcolor=#eeeeee>"
            if Session("Login")<>""  then
            Response.Write "<a href=""javascript:window.open('file1.asp?ID=" & rs("ID") &  "','','toolbar=0,menubar=1,width=500,height=400,resizable=1,scrollbars=yes');location.href='showall.asp?'+'" & request_str &"';"  &  """>"
            Response.Write "<font color=#685362 size=1>Отправьте сообщение</font></a>"
            else
			Response.Write "<font color=#685362 size=1>Для отправки сообщения введите пароль</font> <a href='default1.asp'>здесь</a>"
			end if
            else
            Response.Write "<tr><td bgcolor=#ffffff><font color=#685362 size=1 >"
            'Response.Write "<a href=""file.asp?ID=" & rs("ID") & """>"
            Response.Write rs("PersonName") & "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            'Response.Write rs("Nick")& "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            Response.Write rs("Sex")& "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"
            Response.Write cstr(datediff("yyyy",rs("Age"),Now()))& "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            Response.Write rs("Weight")& "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"
            Response.Write rs("Height")& "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            
            Response.Write rs("aboutme")& "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"
            Response.Write rs("aboutyou")& "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"

            Response.Write rs("pattern")& "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"
            Response.Write rs("dream_life")& "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            Response.Write rs("living")& "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"
            Response.Write rs("favorite_eat")& "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            Response.Write rs("favorite_drink")& "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"
            Response.Write rs("interest")& "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"

            'Response.Write rs("phone")& "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"
            'Response.Write rs("html_page")& "&nbsp;" & "</font></td><td bgcolor=#eeeeee>"
            'Response.Write "<a href=""file.asp?ID=" & rs("ID") & """>"
            Response.Write "<a href=""javascript:window.open('file.asp?ID=" & rs("ID") & "','','toolbar=0,menubar=1,width=500,height=400,resizable=1,scrollbars=yes');location.href='showall.asp'" & """>"
            Response.Write "<font color=#685362 size=1>Фото</font></a>" & "</td><td bgcolor=#ffffff>"
            if Session("Login")<>""  then
            Response.Write "<a href=""javascript:window.open('file1.asp?ID=" & rs("ID") &  "','','toolbar=0,menubar=1,width=500,height=400,resizable=1,scrollbars=yes');location.href='showall.asp?'+'" & request_str &"';"  &  """>"
            Response.Write "<font color=#685362 size=1>Отправьте сообщение</font></a>"
            else
			Response.Write "<font color=#685362 size=1>Для отправки сообщения введите пароль</font> <a href='default1.asp'>здесь</a>"
			end if
            end if 
            
            Response.Write "</td></tr>"
            rs.MoveNext
            j=j+1
            
         Wend
                  
			
            
   '   if rs.RecordCount=0 then
   '      Response.Write "Извините, ничего не найдено..."
         
   '   End If
      
     ' rs.Close
      Set rs = Nothing
%>
   </table>
   <br>
   <br>
   <br>
   <center>
   
   <%if Session("first")=1 then
  
   %>
			<%if howmuch>=1 then%>
			<a href='showall.asp?page=1'>1 </a>
			<%end if%>
			<%if howmuch>=10  then%>
            <a href='showall.asp?page=2'>| 2 </a>
            <%end if%>
            <%if howmuch>=20  then%>
            <a href='showall.asp?page=3'>| 3 </a>
            <%end if%>
            <%if howmuch>=30  then%>
            <a href='showall.asp?page=4'>| 4 </a>
            <%end if%>
            <%if howmuch>=40  then%>
            <a href='showall.asp?page=5'>| 5 </a>
            <%end if%>
            <%if howmuch>=50 then%>
            <a href='showall.asp?page=6'>| 6 </a>
            <%end if%>
            <%if howmuch>=60  then%>
            <a href='showall.asp?page=7'>| 7 </a>
            <%end if%>
            <%if howmuch>=70  then%>
            <a href='showall.asp?page=8'>| 8 </a>
            <%end if%>
            <%if howmuch>=80  then%>
            <a href='showall.asp?page=9'>| 9 </a>
            <%end if%>
            <%if howmuch>=100  then%>
            <a href='showall.asp?page=10'>| 10 </a>
            <%end if%>
            <%if howmuch>=110 then%>
            <a href='showall.asp?page=10000'>| Дальше</a>
            <%end if%>
   <%end if%>
   <%if Session("first")=2 then%>
            <%if howmuch>=110 then%>
			<a href='showall.asp?page=11'>11 </a>
			<%end if%>
			<%if howmuch>=120 then%>
            <a href='showall.asp?page=12'>| 12 </a>
            <%end if%>
            <%if howmuch>=130 then%>
            <a href='showall.asp?page=13'>| 13 </a>
            <%end if%>
            <%if howmuch>=140 then%>
            <a href='showall.asp?page=14'>| 14 </a>
            <%end if%>
            <%if howmuch>=150 then%>
            <a href='showall.asp?page=15'>| 15 </a>
            <%end if%>
            <%if howmuch>=160 then%>
            <a href='showall.asp?page=16'>| 16 </a>
            <%end if%>
            <%if howmuch>=170 then%>
            <a href='showall.asp?page=17'>| 17 </a>
            <%end if%>
            <%if howmuch>=180 then%>
            <a href='showall.asp?page=18'>| 18 </a>
            <%end if%>
            <%if howmuch>=190 then%>
            <a href='showall.asp?page=19'>| 19 </a>
            <%end if%>
            <%if howmuch>=200 then%>
            <a href='showall.asp?page=20'>| 20 </a>
            <%end if%>
            <%if howmuch>=210 then%>
            <a href='showall.asp?page=20000'>| Дальше</a>
            <%end if%>
   <%end if%>         
   <%if Session("first")=3 then%>
            <%if howmuch>=210 then%>
			<a href='showall.asp?page=21'>21 |</a>
			<%end if%>
			<%if howmuch>=220 then%>
            <a href='showall.asp?page=22'>22 |</a>
            <%end if%>
            <%if howmuch>=230 then%>
            <a href='showall.asp?page=23'>23 |</a>
            <%end if%>
            <%if howmuch>=240 then%>
            <a href='showall.asp?page=24'>24 |</a>
            <%end if%>
            <%if howmuch>=250 then%>
            <a href='showall.asp?page=25'>25 |</a>
            <%end if%>
            <%if howmuch>=260 then%>
            <a href='showall.asp?page=26'>26 |</a>
            <%end if%>
            <%if howmuch>=270 then%>
            <a href='showall.asp?page=27'>27 |</a>
            <%end if%>
            <%if howmuch>=280 then%>
            <a href='showall.asp?page=28'>28 |</a>
            <%end if%>
            <%if howmuch>=290 then%>
            <a href='showall.asp?page=29'>29 |</a>
            <%end if%>
            <%if howmuch>=300 then%>
            <a href='showall.asp?page=30'>30 |</a>
            <%end if%>
            <%if howmuch>=310 then%>
            <a href='showall.asp?page=30000'>Дальше</a>
            <%end if%>
   <%end if%>   
   
   <%if Session("first")=4 then%>
            <%if howmuch>=310 then%>
			<a href='show.asp?page=31'>31 |</a>
			<%end if%>
			<%if howmuch>=320 then%>
            <a href='show.asp?page=32'>32 |</a>
            <%end if%>
            <%if howmuch>=330 then%>
            <a href='show.asp?page=33'>33 |</a>
            <%end if%>
            <%if howmuch>=340 then%>
            <a href='show.asp?page=34'>34 |</a>
            <%end if%>
            <%if howmuch>=350 then%>
            <a href='show.asp?page=35'>35 |</a>
            <%end if%>
            <%if howmuch>=360 then%>
            <a href='show.asp?page=36'>36 |</a>
            <%end if%>
            <%if howmuch>=370 then%>
            <a href='show.asp?page=37'>37 |</a>
            <%end if%>
            <%if howmuch>=380 then%>
            <a href='show.asp?page=38'>38 |</a>
            <%end if%>
            <%if howmuch>=390 then%>
            <a href='show.asp?page=39'>39 |</a>
            <%end if%>
            <%if howmuch>=400 then%>
            <a href='show.asp?page=40'>40 |</a>
            <%end if%>
            <%if howmuch>=410 then%>
            <a href='show.asp?page=40000'>Дальше</a>
            <%end if%>
   <%end if%>        
   <%if Session("first")=5 then%>
            <%if howmuch>=410 then%>
			<a href='show.asp?page=41'>41 |</a>
			<%end if%>
			<%if howmuch>=420 then%>
            <a href='show.asp?page=42'>42 |</a>
            <%end if%>
            <%if howmuch>=430 then%>
            <a href='show.asp?page=43'>43 |</a>
            <%end if%>
            <%if howmuch>=440 then%>
            <a href='show.asp?page=44'>44 |</a>
            <%end if%>
            <%if howmuch>=450 then%>
            <a href='show.asp?page=45'>45 |</a>
            <%end if%>
            <%if howmuch>=460 then%>
            <a href='show.asp?page=46'>46 |</a>
            <%end if%>
            <%if howmuch>=470 then%>
            <a href='show.asp?page=47'>47 |</a>
            <%end if%>
            <%if howmuch>=480 then%>
            <a href='show.asp?page=48'>48 |</a>
            <%end if%>
            <%if howmuch>=490 then%>
            <a href='show.asp?page=49'>49 |</a>
            <%end if%>
            <%if howmuch>=500 then%>
            <a href='show.asp?page=50'>50 |</a>
            <%end if%>
            <%if howmuch>=510 then%>
            <a href='show.asp?page=50000'>Дальше</a>
            <%end if%>
   <%end if%>        
   <%if Session("first")=6 then%>
            <%if howmuch>=510 then%>
			<a href='show.asp?page=51'>51 |</a>
			<%end if%>
			<%if howmuch>=520 then%>
            <a href='show.asp?page=52'>52 |</a>
            <%end if%>
            <%if howmuch>=530 then%>
            <a href='show.asp?page=53'>53 |</a>
            <%end if%>
            <%if howmuch>=540 then%>
            <a href='show.asp?page=54'>54 |</a>
            <%end if%>
            <%if howmuch>=550 then%>
            <a href='show.asp?page=55'>55 |</a>
            <%end if%>
            <%if howmuch>=560 then%>
            <a href='show.asp?page=56'>56 |</a>
            <%end if%>
            <%if howmuch>=570 then%>
            <a href='show.asp?page=57'>57 |</a>
            <%end if%>
            <%if howmuch>=580 then%>
            <a href='show.asp?page=58'>58 |</a>
            <%end if%>
            <%if howmuch>=590 then%>
            <a href='show.asp?page=59'>59 |</a>
            <%end if%>
            <%if howmuch>=600 then%>
            <a href='show.asp?page=60'>60 |</a>
            <%end if%>
            <%if howmuch>=610 then%>
            <a href='show.asp?page=60000'>Дальше</a>
            <%end if%>
   <%end if%>        
   <%if Session("first")=7 then%>
            <%if howmuch>=610 then%>
			<a href='show.asp?page=61'>61 |</a>
			<%end if%>
			<%if howmuch>=620 then%>
            <a href='show.asp?page=62'>62 |</a>
            <%end if%>
            <%if howmuch>=630 then%>
            <a href='show.asp?page=63'>63 |</a>
            <%end if%>
            <%if howmuch>=640 then%>
            <a href='show.asp?page=64'>64 |</a>
            <%end if%>
            <%if howmuch>=650 then%>
            <a href='show.asp?page=65'>65 |</a>
            <%end if%>
            <%if howmuch>=660 then%>
            <a href='show.asp?page=66'>66 |</a>
            <%end if%>
            <%if howmuch>=670 then%>
            <a href='show.asp?page=67'>67 |</a>
            <%end if%>
            <%if howmuch>=680 then%>
            <a href='show.asp?page=68'>68 |</a>
            <%end if%>
            <%if howmuch>=690 then%>
            <a href='show.asp?page=69'>69 |</a>
            <%end if%>
            <%if howmuch>=700 then%>
            <a href='show.asp?page=70'>70 |</a>
            <%end if%>
            <%if howmuch>=710 then%>
            <a href='show.asp?page=70000'>Дальше</a>
            <%end if%>
   <%end if%>        
   <%if Session("first")=8 then%>
            <%if howmuch>=710 then%>
			<a href='show.asp?page=71'>71 |</a>
			<%end if%>
			<%if howmuch>=720 then%>
            <a href='show.asp?page=72'>72 |</a>
            <%end if%>
            <%if howmuch>=730 then%>
            <a href='show.asp?page=73'>73 |</a>
            <%end if%>
            <%if howmuch>=740 then%>
            <a href='show.asp?page=74'>74 |</a>
            <%end if%>
            <%if howmuch>=750 then%>
            <a href='show.asp?page=75'>75 |</a>
            <%end if%>
            <%if howmuch>=760 then%>
            <a href='show.asp?page=76'>76 |</a>
            <%end if%>
            <%if howmuch>=770 then%>
            <a href='show.asp?page=77'>77 |</a>
            <%end if%>
            <%if howmuch>=780 then%>
            <a href='show.asp?page=78'>78 |</a>
            <%end if%>
            <%if howmuch>=790 then%>
            <a href='show.asp?page=79'>79 |</a>
            <%end if%>
            <%if howmuch>=800 then%>
            <a href='show.asp?page=80'>80 |</a>
            <%end if%>
            <%if howmuch>=810 then%>
            <a href='show.asp?page=80000'>Дальше</a>
            <%end if%>
   <%end if%>        
   <%if Session("first")=9 and  Session("first")<>"" then%>
            <%if howmuch>=810 then%>
			<a href='show.asp?page=81'>81 |</a>
			<%end if%>
			<%if howmuch>=820 then%>
            <a href='show.asp?page=82'>82 |</a>
            <%end if%>
            <%if howmuch>=830 then%>
            <a href='show.asp?page=83'>83 |</a>
            <%end if%>
            <%if howmuch>=840 then%>
            <a href='show.asp?page=84'>84 |</a>
            <%end if%>
            <%if howmuch>=850 then%>
            <a href='show.asp?page=85'>85 |</a>
            <%end if%>
            <%if howmuch>=860 then%>
            <a href='show.asp?page=86'>86 |</a>
            <%end if%>
            <%if howmuch>=870 then%>
            <a href='show.asp?page=87'>87 |</a>
            <%end if%>
            <%if howmuch>=880 then%>
            <a href='show.asp?page=88'>88 |</a>
            <%end if%>
            <%if howmuch>=890 then%>
            <a href='show.asp?page=89'>89 |</a>
            <%end if%>
            <%if howmuch>=900 then%>
            <a href='show.asp?page=90'>90 |</a>
            <%end if%>
            <%if howmuch>=910 then%>
            <a href='show.asp?page=90000'>Дальше</a>
            <%end if%>
   <%end if%>        
                  
   </center>
   <br>
   <br>
   <table border=0 >
<td>
<tr>		
<td>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

</td>


</tr>
</table>
<center>
<table border=0>		
<td><a href="search.asp"><font size=3 color = #111abc>Найти друга или подругу</font></a></td><td>&nbsp;&nbsp;|&nbsp;&nbsp;</td>
<td><a href="history.asp"><font  size=3 color = #111abc>Веcелые истории, анекдоты, приколы</font></a></td><td>&nbsp;&nbsp;|&nbsp;&nbsp;</td>
<td><a href="reg.asp"><font  size=3 color = #111abc>Зарегистрироваться</font></a></td>
</table>
</center>
<table>
		


<tr>
<td>&nbsp;&nbsp;</td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
</tr>
<tr>
<td>&nbsp;&nbsp;</td>
</tr>

</td>
</table>

<center><FONT size="1" color="#0000ff" style="FONT-WEIGHT: bold">&copy;2005, Solva SoftWare inc. All rights Reserved 

</font>

</center>
</body>
</html>