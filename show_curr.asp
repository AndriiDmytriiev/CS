<%
   ' -- show.asp --
   ' Generates a list of uploaded files
  ' response.Write "-------------------2---"
  ' response.Write cstr(request.form("contact"))
  ' response.Write "--------------------3--"
   Session("PersonID")=""
   'on error resume next
   Session("first")=1
   if Request.QueryString<>"" and Session("qs")="" then
   Session("qs")=Request.QueryString
   end if
   'if len(Request.QueryString)<30 then
   '		 Session("qs")=Request.QueryString
   '  end if
on error resume next   
Response.Buffer = True
   
   ' Connection String
      Dim connStr
      dim conn 
      set conn = server.CreateObject("ADODB.Connection")
      
     'connStr = "DRIVER=Microsoft Access Driver (*.mdb);DBQ="
     'connStr = connStr & Server.MapPath("/andy26/database/missyou.mdb")
      
      connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
      Server.MapPath("missyou.mdb")
      conn.Open connStr
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="index.css">
<META content="text/html; charset=windows-1251" http-equiv=Content-Type> 
<META NAME="Keywords" CONTENT="">

<script>
//function ddd(){
//history.back();	
//}

</script>
   <title>Окунитесь в роскошь человеческого общения</title>
   
</head>
<body background="bg.jpg" alink="#0000ff" vlink="#0000ff" link="#0000ff">
   <p align=center>
   <center>
     <td align=center>
<img src="relationshipsromance.jpeg"></img>
</td><br>


<br>

      <a href="reg.asp"></a>
      <br>
   </center>
 </p>  

<table width="80%"  align="center" color="#123123" border=0>
<tr>
<td><font size=3>Инициатор:<%=Session("Login")%></font></td>
</tr>
<tr>
<td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td>
</tr>


		<tr>  
 		<td valign="top">
                <p>
                <p>
                <input type=button width=100 name=b1 onclick="javascript:location.href='show_curr.asp'" value="Текущие заявки         ">   
		<p>
                <input type=button width=100  name=b2 onclick="javascript:location.href='show_old.asp'" value="Прошлые заявки         ">   
        	<p>
                <input type=button width=100  name=b3 onclick="javascript:location.href='show.asp'"     value="Все заявки                   ">   
		<p>
		<input type=button width=100 name=b4 value="Создать новую заявку " onclick="javascript:window.open('common.asp','','toolbar=0,menubar=1,width=500,height=400,resizable=1,scrollbars=yes');">
		
		
                </td>



<%
   ' Recordset Object
   Dim rs,rs1,rs2, rs3
   dim flag
   dim cond(10)
'   for i=0 to 10 
'		cond(i)=0
 '  next
   '-------------Filling conditions of search
   
   'if Request("PersonName")<>"" then cond(0)=1
   'if Request("Nick")<>"" then cond(1)=1
   'if Request("City")<>"" then cond(2)=1 else cond(2)=0
   'if Request("Sex")<>"" then cond(3)=1
   'if Request("age1")<>"" then cond(4)=1
   'if Request("age2")<>"" then cond(5)=1
   'if Request("telephone")="on" then cond(6)=1
   'if Request("Photo")="on" then cond(7)=1
   'if Request("weight")<>"" then cond(8)=1
   'if Request("height")<>"" then cond(9)=1
   'Response.Write("<font color=white>")
   'for i=0 to 9
   'Response.Write cstr("cond(" & cstr(i)& ")=" & cond(i))

'   next
   
   '------------------------------------------
   flag=false
      Set rs = Server.CreateObject("ADODB.Recordset")
      Set rs1 = Server.CreateObject("ADODB.Recordset")

      set rs1=conn.Execute("select PersonName from Persons where Nick='" & Session("login") & "'")
      'set rs2=conn.Execute("select c.ClientName from Persons p, Client c  where c.ID = p.ID")
      
      ' opening connection
      'rs.Open "select [ID],[SaleID],[ClientID],[datebeg]," & _
	 ' 	"[dateend] from InternClean where SaleID='" & Request.Form("sale") & "'" & _
	  '	" and ClientID='" & Request.Form("client") & "'" & " and datebeg between '" & "01/01/2006" & "' and '" &  "01/01/2006" & "'" & _
	  '	" order by [ID] desc", connStr, 3, 4
      
      
      sql="select distinct i.NickName,c.ClientName,p.PersonName,c.Address,i.Address1, i.Address2, s.App1,s.App2,s.App3,s.App4,s.App5,s.App6,i.datebeg,i.dateend,i.ID from StatusHistory s,InternClean i,Persons p, Client c  where c.clientID = p.clientID and p.Nick='" & Session("login") & "' and p.ID=i.SaleID and i.ID=s.CleanID and i.old=0" 
      if request("client")<>"0" then
     ' response.Write sql
	       'sql= sql & " and c.ClientID='" & request("client") & "' order by i.ID desc"
      end if 	  
      if request.Form("contact")="" then
       'sql= sql & " and p.ID='" & request("contact") & "' order by i.ID desc"
      end if 	  
if request("client")="0" and request("contact")="0" then
	       sql= sql &  " order by i.ID desc"
end if
        'Response.Write ("<font color=#0000ff")
	     Response.Write "<td><table width='80%'  align=center color=#123123 border=0><tr><td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Имя заявки</font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Клиент</font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Инициатор</font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Адрес клиента </font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Место откуда</font></b></td>"
                   Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Место куда</font></b></td>"
	     Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Дата заявки</font></b></td>"
	     'Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Показать заявку</font></b></td></tr>"
	     
        Response.Write "<tr><td align=center bgcolor=#ffffff><b><font color=#685362 size=2>&nbsp;</font></b></td>"
	     Response.Write "<td align=center bgcolor=#ffffff><b><font color=#685362 size=2>&nbsp;</font></b></td>"
	     Response.Write "<td align=center bgcolor=#ffffff><b><font color=#685362 size=2>&nbsp;</font></b></td>"
	     Response.Write "<td align=center bgcolor=#ffffff><b><font color=#685362 size=2>&nbsp;</font></b></td>"	     
	      'Response.Write "<td align=center bgcolor=#68ffff><b><font color=#685362 size=2>Показать заявку</font></b></td></tr>"
  ' howmuch=rs.RecordCount     
    'call SetPage (RS, 1, 10, 0)

	  	  
	 '  sql1=" where 0=0 "	
	  	 
	 ' if cond(0)=1 then
	 ' sql1=sql1 & " and PersonName='" & Request("personname") & "' " 
	 ' end if
	 ' if cond(1)=1 then
	 ' sql1=sql1 &   " and Nick='" & Request("Nick") & "' " 
	 ' end if
	 ' if cond(2)=1 then
	 ' sql1=sql1 &   " and City='" & Request.QueryString("City") & "' " 
	 ' end if
	  
	  

	'  if cond(3)=1 then
	'  sql1=sql1 &   " and sex='" & Request("Sex") & "' " 
	'  end if
	'  on error resume next
	'  if request("age1")<>"" and request("age2")<>"" then
'		if cond(4)=1 and cond(5)=1 and cint(request("age1"))<=cint(request("age2")) then
'			sql1=sql1 &   " and (DateDiff('yyyy',Age,Date()) between '" & request("age1") & "' and '" &  request("age2") & "')"
'		end if
'	  end if
'	  if cond(6)=1 then
'	  sql1=sql1 &   " and  phone<>'' "
'	  end if
'	  if cond(7)=1 then
'	  sql1=sql1 &   " and (not isnull(Photo1) or  len(trim(pic1))>2) "
'	  end if
'	  if cond(8)=1 then
'	  sql1=sql1 &   " and weight<='" & Request("weight") & "' " 
'	  end if
'	  if cond(9)=1 then
'	  sql1=sql1 &   " and height<='" & Request("height") & "' " 
'	  end if
'	  sql1=sql1 & " order by [ID] desc"
'	  
 '     sql=sql & sql1
      
      
  '    request_str="personname="&request("personname")&"&nick="&request("nick")&"&City="&request("City")&"&Sex="&request("sex")&"&age1="&request("age1")&"&age2="&request("age2")&"&height="&request("height")&"&weight=" & request("weight") 
      'Response.Write sql
         'Response.Write page
'response.Write sql
         rs.open sql,connStr,3,4
         'Response.Write Session("first")
       j=1
			'Session("first")=1
			if request("page")="" then 
			page=1
			else
			page=cint(request("page"))
			end if
			
            if request("page")="10000" or  (cint(request("page"))>10 and cint(request("page"))<=20) then
			    Session("first")=2   
				page=11
			end if
			
			if request("page")="20000" or  (cint(request("page"))>20 and cint(request("page"))<=30)then
			    Session("first")=3   
				page=21
			end if
			if request("page")="30000" or  (cint(request("page"))>30 and cint(request("page"))<=40)then
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
			    rs.MoveFirst
				While Not rs.EOF and j<(page-1)*10
				          rs.MoveNext
				        
				j=j+1
				Wend
            
 			end if
 			
        
       howmuch=rs.RecordCount
	    'rs.Open "select [ID],[City],[PersonName],[Nick]," & _
		'"[Sex],[Age],[Weight],[Height],[aboutme],[aboutyou],[phone],[html_page],[Photo1] from Persons where PersonName='" & Request.Form("personname") & "'" & _
		'" and Sex='" & Request.Form("Sex") & "'" & " and Age between '" & request("age1") & "' and '" &  request("age2") & "'" & _
		'" and telephone <>'' and not isnull(Photo1)  " & " order by [ID] desc", connStr, 3, 4
		
	     'Response.Write ("<font color=#0000ff>")
	      
         'rs.MoveFirst
   on error resume next     
         'response.Write "<br>" & sql
         While Not rs.EOF 'and j<page*10
         '   if (j mod 2) = 1  then 
            Response.Write "<tr><td bgcolor=#ffffff><font color=#685362 size=1 >"
            'Response.Write "<a href=""file.asp?ID=" & rs("ID") & """>"
            'Response.Write "<a href=""javascript:window.open('default1_0.asp?ID=" & rs("ID") & "','','toolbar=0,menubar=1,width=500,height=400,resizable=1,scrollbars=yes');location.href='show.asp?'+'" & Session("qs") &"';history.go(0);"  &  """>"
            'Response.Write "<font color=#685362 size=1>"& rs(0) &"</font></a>" & "</td><td bgcolor=#ffffff>"
             Response.Write "<a href=""javascript:window.open('default1_0_nonedit.asp?ID=" & rs("ID") &  "','','toolbar=0,menubar=1,width=500,height=400,resizable=1,scrollbars=yes');location.href='show.asp?'+'" & Session("qs") &"';history.go(0);"  &  """>"
            Response.Write "<font color=#685362 size=1>" & rs(0) & "</font></a></td>"

            Response.Write "<td bgcolor=#eeeeee><font color=#685362 size=1 >"
            'Response.Write "<a href=""file.asp?ID=" & rs("ID") & """>"

            'Response.Write "</font></td><td bgcolor=#eeeeee valign='top'><font color=#685362 size=1>"
            Response.Write rs(1) & "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"
            Response.Write rs(2) & "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            Response.Write rs(3) & "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"
            Response.Write rs(4) & "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            Response.Write rs(5) & "&nbsp;" & "</font></td><td bgcolor=#ffffff><font color=#685362 size=1>"
            Response.Write rs(12) & "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            

            'Response.Write "&nbsp;" & rs("datebeg") & "&nbsp;" & "</font></td><td bgcolor=#eeeeee><font color=#685362 size=1>"
            'Response.Write "<a href=""javascript:window.open('default_1_0_nonedit.asp?ID=" & rs("ID") & "','','toolbar=0,menubar=1,width=500,height=400,resizable=1,scrollbars=yes');location.href='show.asp?'+'" & Session("qs") &"';history.go(0);"  &  """>"
            'Response.Write "<font color=#685362 size=1>Заявка</font></a>" & "</td>"
            'Response.Write "<tr><td bgcolor=#eeeeee><font color=#685362 size=1 >"
            'Response.Write "<a href=""file.asp?ID=" & rs("ID") & """>"
           
            
            Response.Write "</td></tr>"
            rs.MoveNext
            'j=j+1
         Wend
         
            response.write 
      
            
     
        
      if rs.RecordCount=0 then
         Response.Write "Извините, ничего не найдено..."
         
      End If
      
     ' rs.Close
      Set rs = Nothing
      	
	 
%>
   </table>
   <br>
   <br>
   <center>
   <%if Session("first")=1 then

   %>
			<%if howmuch>=1 then%>
			<a href='show.asp?page=1&<%=Session("qs")%>'> 1 </a>
			<%end if%>
			<%if howmuch>=10  then%>
			<a href='show.asp?page=2&<%=Session("qs")%>'>| 2 </a>
            <%end if%>
            <%if howmuch>=20  then%>
            <a href='show.asp?page=3&<%=Session("qs")%>'>| 3 </a>
            <%end if%>
            <%if howmuch>=30  then%>
            <a href='show.asp?page=4&<%=Session("qs")%>'>| 4 </a>
            <%end if%>
            <%if howmuch>=40  then%>
            <a href='show.asp?page=5&<%=Session("qs")%>'>| 5 </a>
            <%end if%>
            <%if howmuch>=50 then%>
            <a href='show.asp?page=6&<%=Session("qs")%>'>| 6 </a>
            <%end if%>
            <%if howmuch>=60  then%>
            <a href='show.asp?page=7&<%=Session("qs")%>'>| 7 </a>
            <%end if%>
            <%if howmuch>=70  then%>
            <a href='show.asp?page=8&<%=Session("qs")%>'>| 8 </a>
            <%end if%>
            <%if howmuch>=80  then%>
            <a href='show.asp?page=9&<%=Session("qs")%>'>| 9 </a>
            <%end if%>
            <%if howmuch>=100  then%>
            <a href='show.asp?page=10&<%=Session("qs")%>'>| 10 </a>
            <%end if%>
            <%if howmuch>=110 then%>
            <a href='show.asp?page=10000&<%=Session("qs")%>'>| Дальше</a>
            <%end if%>
   <%end if%>
   <%if Session("first")=2 then%>
            <%if howmuch>=110 then%>
			<a href='show.asp?page=11&<%=Session("qs")%>'>11 </a>
			<%end if%>
			<%if howmuch>=120 then%>
            <a href='show.asp?page=12&<%=Session("qs")%>'>| 12 </a>
            <%end if%>
            <%if howmuch>=130 then%>
            <a href='show.asp?page=13&<%=Session("qs")%>'>| 13 </a>
            <%end if%>
            <%if howmuch>=140 then%>
            <a href='show.asp?page=14&<%=Session("qs")%>'>| 14 </a>
            <%end if%>
            <%if howmuch>=150 then%>
            <a href='show.asp?page=15&<%=Session("qs")%>'>| 15 </a>
            <%end if%>
            <%if howmuch>=160 then%>
            <a href='show.asp?page=16&<%=Session("qs")%>'>| 16 </a>
            <%end if%>
            <%if howmuch>=170 then%>
            <a href='show.asp?page=17&<%=Session("qs")%>'>| 17 </a>
            <%end if%>
            <%if howmuch>=180 then%>
            <a href='show.asp?page=18&<%=Session("qs")%>'>| 18 </a>
            <%end if%>
            <%if howmuch>=190 then%>
            <a href='show.asp?page=19&<%=Session("qs")%>'>| 19 </a>
            <%end if%>
            <%if howmuch>=200 then%>
            <a href='show.asp?page=20&<%=Session("qs")%>'>| 20 </a>
            <%end if%>
            <%if howmuch>=210 then%>
            <a href='show.asp?page=20000&<%=Session("qs")%>'>| Дальше</a>
            <%end if%>
   <%end if%>         
   <%if Session("first")=3 then%>
            <%if howmuch>=210 then%>
			<a href='show.asp?page=21&<%=Session("qs")%>'>21 |</a>
			<%end if%>
			<%if howmuch>=220 then%>
            <a href='show.asp?page=22&<%=Session("qs")%>'>22 |</a>
            <%end if%>
            <%if howmuch>=230 then%>
            <a href='show.asp?page=23&<%=Session("qs")%>'>23 |</a>
            <%end if%>
            <%if howmuch>=240 then%>
            <a href='show.asp?page=24&<%=Session("qs")%>'>24 |</a>
            <%end if%>
            <%if howmuch>=250 then%>
            <a href='show.asp?page=25&<%=Session("qs")%>'>25 |</a>
            <%end if%>
            <%if howmuch>=260 then%>
            <a href='show.asp?page=26&<%=Session("qs")%>'>26 |</a>
            <%end if%>
            <%if howmuch>=270 then%>
            <a href='show.asp?page=27&<%=Session("qs")%>'>27 |</a>
            <%end if%>
            <%if howmuch>=280 then%>
            <a href='show.asp?page=28&<%=Session("qs")%>'>28 |</a>
            <%end if%>
            <%if howmuch>=290 then%>
            <a href='show.asp?page=29&<%=Session("qs")%>'>29 |</a>
            <%end if%>
            <%if howmuch>=300 then%>
            <a href='show.asp?page=30&<%=Session("qs")%>'>30 |</a>
            <%end if%>
            <%if howmuch>=310 then%>
            <a href='show.asp?page=30000&<%=Session("qs")%>'>Дальше</a>
            <%end if%>
   <%end if%>   
   <%if Session("first")=4 then%>
            <%if howmuch>=310 then%>
			<a href='show.asp?page=31&<%=Session("qs")%>'>31 |</a>
			<%end if%>
			<%if howmuch>=320 then%>
            <a href='show.asp?page=32&<%=Session("qs")%>'>32 |</a>
            <%end if%>
            <%if howmuch>=330 then%>
            <a href='show.asp?page=33&<%=Session("qs")%>'>33 |</a>
            <%end if%>
            <%if howmuch>=340 then%>
            <a href='show.asp?page=34&<%=Session("qs")%>'>34 |</a>
            <%end if%>
            <%if howmuch>=350 then%>
            <a href='show.asp?page=35&<%=Session("qs")%>'>35 |</a>
            <%end if%>
            <%if howmuch>=360 then%>
            <a href='show.asp?page=36&<%=Session("qs")%>'>36 |</a>
            <%end if%>
            <%if howmuch>=370 then%>
            <a href='show.asp?page=37&<%=Session("qs")%>'>37 |</a>
            <%end if%>
            <%if howmuch>=380 then%>
            <a href='show.asp?page=38&<%=Session("qs")%>'>38 |</a>
            <%end if%>
            <%if howmuch>=390 then%>
            <a href='show.asp?page=39&<%=Session("qs")%>'>39 |</a>
            <%end if%>
            <%if howmuch>=400 then%>
            <a href='show.asp?page=40&<%=Session("qs")%>'>40 |</a>
            <%end if%>
            <%if howmuch>=410 then%>
            <a href='show.asp?page=40000&<%=Session("qs")%>'>Дальше</a>
            <%end if%>
   <%end if%>        
   <%if Session("first")=5 then%>
            <%if howmuch>=410 then%>
			<a href='show.asp?page=41&<%=Session("qs")%>'>41 |</a>
			<%end if%>
			<%if howmuch>=420 then%>
            <a href='show.asp?page=42&<%=Session("qs")%>'>42 |</a>
            <%end if%>
            <%if howmuch>=430 then%>
            <a href='show.asp?page=43&<%=Session("qs")%>'>43 |</a>
            <%end if%>
            <%if howmuch>=440 then%>
            <a href='show.asp?page=44&<%=Session("qs")%>'>44 |</a>
            <%end if%>
            <%if howmuch>=450 then%>
            <a href='show.asp?page=45&<%=Session("qs")%>'>45 |</a>
            <%end if%>
            <%if howmuch>=460 then%>
            <a href='show.asp?page=46&<%=Session("qs")%>'>46 |</a>
            <%end if%>
            <%if howmuch>=470 then%>
            <a href='show.asp?page=47&<%=Session("qs")%>'>47 |</a>
            <%end if%>
            <%if howmuch>=480 then%>
            <a href='show.asp?page=48&<%=Session("qs")%>'>48 |</a>
            <%end if%>
            <%if howmuch>=490 then%>
            <a href='show.asp?page=49&<%=Session("qs")%>'>49 |</a>
            <%end if%>
            <%if howmuch>=500 then%>
            <a href='show.asp?page=50&<%=Session("qs")%>'>50 |</a>
            <%end if%>
            <%if howmuch>=510 then%>
            <a href='show.asp?page=50000&<%=Session("qs")%>'>Дальше</a>
            <%end if%>
   <%end if%>        
   <%if Session("first")=6 then%>
            <%if howmuch>=510 then%>
			<a href='show.asp?page=51&<%=Session("qs")%>'>51 |</a>
			<%end if%>
			<%if howmuch>=520 then%>
            <a href='show.asp?page=52&<%=Session("qs")%>'>52 |</a>
            <%end if%>
            <%if howmuch>=530 then%>
            <a href='show.asp?page=53&<%=Session("qs")%>'>53 |</a>
            <%end if%>
            <%if howmuch>=540 then%>
            <a href='show.asp?page=54&<%=Session("qs")%>'>54 |</a>
            <%end if%>
            <%if howmuch>=550 then%>
            <a href='show.asp?page=55&<%=Session("qs")%>'>55 |</a>
            <%end if%>
            <%if howmuch>=560 then%>
            <a href='show.asp?page=56&<%=Session("qs")%>'>56 |</a>
            <%end if%>
            <%if howmuch>=570 then%>
            <a href='show.asp?page=57&<%=Session("qs")%>'>57 |</a>
            <%end if%>
            <%if howmuch>=580 then%>
            <a href='show.asp?page=58&<%=Session("qs")%>'>58 |</a>
            <%end if%>
            <%if howmuch>=590 then%>
            <a href='show.asp?page=59&<%=Session("qs")%>'>59 |</a>
            <%end if%>
            <%if howmuch>=600 then%>
            <a href='show.asp?page=60&<%=Session("qs")%>'>60 |</a>
            <%end if%>
            <%if howmuch>=610 then%>
            <a href='show.asp?page=60000&<%=Session("qs")%>'>Дальше</a>
            <%end if%>
   <%end if%>        
   <%if Session("first")=7 then%>
            <%if howmuch>=610 then%>
			<a href='show.asp?page=61&<%=Session("qs")%>'>61 |</a>
			<%end if%>
			<%if howmuch>=620 then%>
            <a href='show.asp?page=62&<%=Session("qs")%>'>62 |</a>
            <%end if%>
            <%if howmuch>=630 then%>
            <a href='show.asp?page=63&<%=Session("qs")%>'>63 |</a>
            <%end if%>
            <%if howmuch>=640 then%>
            <a href='show.asp?page=64&<%=Session("qs")%>'>64 |</a>
            <%end if%>
            <%if howmuch>=650 then%>
            <a href='show.asp?page=65&<%=Session("qs")%>'>65 |</a>
            <%end if%>
            <%if howmuch>=660 then%>
            <a href='show.asp?page=66&<%=Session("qs")%>'>66 |</a>
            <%end if%>
            <%if howmuch>=670 then%>
            <a href='show.asp?page=67&<%=Session("qs")%>'>67 |</a>
            <%end if%>
            <%if howmuch>=680 then%>
            <a href='show.asp?page=68&<%=Session("qs")%>'>68 |</a>
            <%end if%>
            <%if howmuch>=690 then%>
            <a href='show.asp?page=69&<%=Session("qs")%>'>69 |</a>
            <%end if%>
            <%if howmuch>=700 then%>
            <a href='show.asp?page=70&<%=Session("qs")%>'>70 |</a>
            <%end if%>
            <%if howmuch>=710 then%>
            <a href='show.asp?page=70000&<%=Session("qs")%>'>Дальше</a>
            <%end if%>
   <%end if%>        
   <%if Session("first")=8 then%>
            <%if howmuch>=710 then%>
			<a href='show.asp?page=71&<%=Session("qs")%>'>71 |</a>
			<%end if%>
			<%if howmuch>=720 then%>
            <a href='show.asp?page=72&<%=Session("qs")%>'>72 |</a>
            <%end if%>
            <%if howmuch>=730 then%>
            <a href='show.asp?page=73&<%=Session("qs")%>'>73 |</a>
            <%end if%>
            <%if howmuch>=740 then%>
            <a href='show.asp?page=74&<%=Session("qs")%>'>74 |</a>
            <%end if%>
            <%if howmuch>=750 then%>
            <a href='show.asp?page=75&<%=Session("qs")%>'>75 |</a>
            <%end if%>
            <%if howmuch>=760 then%>
            <a href='show.asp?page=76&<%=Session("qs")%>'>76 |</a>
            <%end if%>
            <%if howmuch>=770 then%>
            <a href='show.asp?page=77&<%=Session("qs")%>'>77 |</a>
            <%end if%>
            <%if howmuch>=780 then%>
            <a href='show.asp?page=78&<%=Session("qs")%>'>78 |</a>
            <%end if%>
            <%if howmuch>=790 then%>
            <a href='show.asp?page=79&<%=Session("qs")%>'>79 |</a>
            <%end if%>
            <%if howmuch>=800 then%>
            <a href='show.asp?page=80&<%=Session("qs")%>'>80 |</a>
            <%end if%>
            <%if howmuch>=810 then%>
            <a href='show.asp?page=80000&<%=Session("qs")%>'>Дальше</a>
            <%end if%>
   <%end if%>        
   <%if Session("first")=9 then%>
            <%if howmuch>=810 then%>
			<a href='show.asp?page=81&<%=Session("qs")%>'>81 |</a>
			<%end if%>
			<%if howmuch>=820 then%>
            <a href='show.asp?page=82&<%=Session("qs")%>'>82 |</a>
            <%end if%>
            <%if howmuch>=830 then%>
            <a href='show.asp?page=83&<%=Session("qs")%>'>83 |</a>
            <%end if%>
            <%if howmuch>=840 then%>
            <a href='show.asp?page=84&<%=Session("qs")%>'>84 |</a>
            <%end if%>
            <%if howmuch>=850 then%>
            <a href='show.asp?page=85&<%=Session("qs")%>'>85 |</a>
            <%end if%>
            <%if howmuch>=860 then%>
            <a href='show.asp?page=86&<%=Session("qs")%>'>86 |</a>
            <%end if%>
            <%if howmuch>=870 then%>
            <a href='show.asp?page=87&<%=Session("qs")%>'>87 |</a>
            <%end if%>
            <%if howmuch>=880 then%>
            <a href='show.asp?page=88&<%=Session("qs")%>'>88 |</a>
            <%end if%>
            <%if howmuch>=890 then%>
            <a href='show.asp?page=89&<%=Session("qs")%>'>89 |</a>
            <%end if%>
            <%if howmuch>=900 then%>
            <a href='show.asp?page=90&<%=Session("qs")%>'>90 |</a>
            <%end if%>
            <%if howmuch>=910 then%>
            <a href='show.asp?page=90000&<%=Session("qs")%>'>Дальше</a>
            <%end if%>
   <%end if%>        
   </center>
   <br>   <br>   <br>   <br>
   <table border=0 >
    
   </table>
          
   
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
<p></p>
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
</td>
</tr>
</table>
<center><FONT size="1" color="#0000ff" style="FONT-WEIGHT: bold">&copy;2005, Solva SoftWare inc. All rights Reserved 
<%'Response.Write sql & "<br>"		%>
</font></center>
</body>
</html>