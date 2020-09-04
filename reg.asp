<%
on error resume next
 Dim connStr
      dim conn 
      set conn = server.CreateObject("ADODB.Connection")
      
     'connStr = "DRIVER=Microsoft Access Driver (*.mdb);DBQ="
     'connStr = connStr & Server.MapPath("/andy26/database/missyou.mdb")
      
      connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
      Server.MapPath("missyou.mdb")
      conn.Open connStr
      Dim rs1
      Set rs = Server.CreateObject("ADODB.Recordset")



      Set rs1 = Server.CreateObject("ADODB.Recordset")
      set rs1=conn.Execute("select distinct [City] from [Persons]")
    

%>
<HTML>
  <HEAD>
  <META content="text/html; charset=windows-1251" http-equiv=Content-Type> 
<META NAME="Keywords" CONTENT="любовь, отношения, брак, флирт, девушки,эротика, брак, анекдоты, развлечения">

      <script language="javascript">
  function submit_form(){
  //document.form.pass2.value=document.form.pass1.value;
 //if (document.f1.Nick.value=='') ||  (document.f1.pass1.value=='')
 //{
 //alert("Вы не ввели данные! Пожалуйста, заполните.")
 // location.href = 'reg.asp';
 // }
 var i;
 i=0
 if (document.f1.Nick.value=='') i++; 
 if  (document.f1.pass1.value=='') i++;
 if  (document.f1.PersonName.value=='') i++;
 //if  (document.f1.year.value=='') i++;
 //if  (document.f1.weight.value=='') i++;
 //if  (document.f1.height.value=='') i++;
 if (i!=0)
 alert('Вы не ввели данные! Пожалуйста, заполните.')
 else
 document.f1.submit();
  }
  
  </script>
    <title>Регистрация пользователя</title>
   <link rel="stylesheet" type="text/css" href="index.css">
<META content="text/html; charset=windows-1251" http-equiv=Content-Type> 
  </HEAD>
  <BODY background="bg.jpg" alink="#0000ff" vlink="#0000ff" link="#0000ff">

        <FONT size="3" color="#0000ff" style="FONT-WEIGHT: bold">

<center>
     <td align=center>
<img src="relationshipsromance.jpeg"></img>
</td></center><br>


<br>
<center>
<table>
<td width=60% valign=top>
<table border=0 >
<td>
<tr>		
		<td valign=top>
		<a href="search.asp"><font size=3>Найти заявки</font></a>
		<hr>
		</td>
</tr>
<tr>
		<td>
		<a href="default1.asp"><font size=3>Внутренняя уборка территории</font></a>
		<hr>
		</td>
</tr>
<tr>
		<td>
		<a href="history.asp"><font size=3>Уборка территории</font></a>
		<hr>
		</td>
</tr>
<tr>
		<td>
		<a href="history.asp"><font size=3>Дополнительные услуги</font></a>
		<hr>
		</td>
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
<tr>
<td>&nbsp;&nbsp;</td>
</tr>

</td>
</table>
</td> 
<td  width=30%>   
<%if request("wrong_picture")="1" then%>
<script language="javascript"> 
alert("Очень большая фотография, выберите, пожалуйста поменьше, до 50 Килобайт")
</script> 
<%else%>
	<%Session("Nick")=""
	Session("Password")=""
	Session("PersonName")=""
	Session("Sex")=""
	Session("age")=""
	Session("year")=""
	Session("month")=""
	Session("monthname")=""
	Session("day")=""
	Session("weight")=""
	Session("height")=""
	Session("aboutme")=""
	Session("aboutyou")=""
	Session("phone")=""
	Session("html_page")=""
	Session("find_sex")=""
	Session("pattern")=""
	Session("dream_life")=""
	Session("living")=""
	Session("find_sex")=""
	Session("purpose_id")=""
	Session("favorite_eat")=""
	Session("favorite_drink")=""
	Session("interest")=""	      	      	      
	%>

<%end if%>  
<form name="f1" method="POST" enctype="multipart/form-data" action="Insert.asp">


        
<FONT color="#0000ff" style="FONT-WEIGHT: bold">Если Вы еще не зарегистрированы, то з&#1072;&#1087;&#1086;&#1083;&#1085;&#1080;&#1090;&#1077;, &#1087;&#1086;&#1078;&#1072;&#1083;&#1091;&#1081;&#1089;&#1090;&#1072;, &#1089;&#1074;&#1086;&#1102; &#1072;&#1085;&#1082;&#1077;&#1090;&#1091; :</font><br>
<center>
<table>
   <tr> 
	 <td><FONT   color="#0000ff" style="FONT-WEIGHT: bold">&#1053;&#1080;&#1082;:</font> </td><td><input name=Nick type=text value="<%=Session("Nick")%>"></td>
	</tr> 
   <tr>
	 <td><FONT color="#0000ff" style="FONT-WEIGHT: bold">Пароль:</font></td><td><input type="password"  name=pass1 value="<%=Session("Password")%>"></td>
	</tr>
    <tr>
	 <td><FONT color="#0000ff" style="FONT-WEIGHT: bold"></font></td><td><input type="hidden" id=pass2 name=pass2></td>
	</tr>
<br>
<br>
	<tr>
	 <td><FONT color="#0000ff" style="FONT-WEIGHT: bold">&#1048;&#1084;&#1103;:</font></td><td><input name=PersonName 

type=text value="<%=Session("PersonName")%>"></td>
	</tr>
	
	<tr>
	 <td><FONT  color="#0000ff" style="FONT-WEIGHT: bold">&#1043;&#1086;&#1088;&#1086;&#1076;:</font></td><td>
	 <select name="City"><option value="&#1052;&#1086;&#1089;&#1082;&#1074;&#1072;" selected>&#1052;&#1086;&#1089;&#1082;&#1074;&#1072;</option>
		<% rs1.movefirst
                 while not rs1.eof
                 %>
                 <option value="<%=rs1(0)%>"><%=rs1(0)%></option>         
                 <%
                 rs1.movenext
                 wend
                 %>						


	 </select></td>
     </tr> 
     <tr>
	 <td><FONT  color="#0000ff" style="FONT-WEIGHT: bold">Пол:</font></td><td>
	 <%dim i%>
	 <select name="Sex">
	                   <%if Session("Sex")="f" then%>
	                    <option value="f" selected>Женcкий</option>
	                    <option value="m">Мужской</option>
	                    <%else%>
						<option value="f" >Женcкий</option>
	                    <option value="m" selected>Мужской</option>
						<%end if%>
	 </select></td>
     </tr> 
     <tr>
	 <td><FONT  color="#0000ff" style="FONT-WEIGHT: bold">Дата рождения:(день/месяц/год)</font> </td><td>
	 <select name=day  value="<%=Session("day")%>">
	 	 	 
	 	 	 <%if   Session("day")<>"" then%>
	 	 	 <option value="<%=Session("day")%>" selected><%=Session("day")%></option>
	 	 	 <%else%>
		  <option value="1" selected>1</option>
		  <%end if%>
		 <%for i=1 to 31%>
		 <option value="<%=i%>"><%=i%></option>
		 <%next%>
       </select>
	 	 	 
	 	 	 
	 	 <select name=month  value="<%=Session("month")%>">
	         <%if   Session("year")<>"" then%>
	 	 	 <option value="<%=Session("month")%>" selected><%=Session("monthname")%></option>
	 	 	 <%else%>
		 <option value="1">Январь</option>
		 <option value="2">Февраль</option>
		 <option value="3">Март</option>
		 <option value="4">Апрель</option>
		 <option value="5">Май</option>
		 <option value="6">Июнь</option>
		 <option value="7">Июль</option>
		 <option value="8">Август</option>
		 <option value="9">Сентябрь</option>
		 <option value="10">Октябрь</option>
		 <option value="11">Ноябрь</option>
		 <option value="12">Декабрь</option>
		  <%end if%>
		 
		 
       </select>
	 <select name=year> 
	 
	     <%if   Session("year")<>"" then%>
		 <option value="<%=Session("year")%>" selected><%=Session("year")%></option>
		 
		 <%else%>
		  <option value="1980" selected>1980</option>
		  <%end if%>
		 <%for i=1931 to 1995%>
		 <option value="<%=i%>"><%=i%></option>
		 <%next%>
       </select></td>
	 </tr>
	 

	 
	 <tr>
	 <td valign=top><FONT  color="#0000ff" style="FONT-WEIGHT: bold">Должность и должностные обязанности:</font></td><td><textarea name=aboutme rows=10 cols=20><%=Session("aboutme")%></textarea></td>
	 </textarea></tr>
	 <tr>
	 <td></td>
	 
	 </tr>
	 
	 
     <tr>
	 <td width="100%">
	 <FONT color="red"  style="FONT-WEIGHT: bold">Фото до 50Кб.</font><br>
	 </td>
	 </tr>
	 <tr>
	 <td>
	 <FONT  color="#0000ff" style="FONT-WEIGHT: bold">Фото :</td></td><td>
		<input type="file" name="Photo1" size="40"></td></tr>
	<td> </td><td>

	 <tr>
	 <td><input type="button" value="&#1044;&#1086;&#1073;&#1072;&#1074;&#1080;&#1090;&#1100;" onclick="javascript:submit_form();"></td><td><input type="reset" value="&#1054;&#1095;&#1080;&#1089;&#1090;&#1080;&#1090;&#1100;"></td>
   
	 </tr>



	
</table>
</center>
</form>
</td>
<td  width=30%>
<table>

</table>
</td>
</table>

</FONT>
<br>
<br>
<br>
<br>
<br>
<br>

<center><FONT size="1" color="#0000ff" style="FONT-WEIGHT: bold">&copy;2005, Solva SoftWare inc. All rights Reserved 

</font></center>
  </BODY>
</HTML>