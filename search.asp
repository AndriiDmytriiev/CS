<HTML>
  <HEAD>
    <title>Окунитесь в роскошь человеческого общения</title>
<link rel="stylesheet" type="text/css" href="index.css">
<META content="text/html; charset=windows-1251" http-equiv=Content-Type> 
<META NAME="Keywords" CONTENT="">

  </HEAD>
  <BODY background="bg.jpg" alink="#32FFFF" vlink="#32FFFF" link="#32FFFF">
      
      <center>
     <td align=center>
<img src="relationshipsromance.jpeg"></img>

</td></center><br>
<%
Session("qs")=""
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
     ' set rs1=conn.Execute("select distinct [City] from [Persons]")
    

%>

<br>
</center> 
  
<form name=input action=show.asp id="input" method="get">

<table>
<tr>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                                  <FONT size=3 color="#0000ff" style="FONT-WEIGHT: bold">&#1055;&#1086;&#1080;&#1089;&#1082;:</font> </tr>

<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                     
</td>
</tr>
<tr>
<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                     </td>
<td>
<table>
	 <td><FONT  color="#0000ff" style="FONT-WEIGHT: bold">Инициатор:</font>  </td><td><input name=sale ></td>
	</tr>
	<tr> 
	 <td><FONT  color="#0000ff" style="FONT-WEIGHT: bold">Клиент:</font>  </td><td><input name=client ></td>
	</tr> 
	<tr>
	 
     </tr> 
      <tr> <td>Дата начала работ</td>  </tr>
      <tr>
         <td>c &nbsp;&nbsp;&nbsp;<select name=day  value="<%=Session("day")%>">
	 	 	 
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
		  <option value="2006" selected>2006</option>
		  <%end if%>
		 <%for i=2006 to 2050%>
		 <option value="<%=i%>"><%=i%></option>
		 <%next%>
       </select>
</td>	 
         </tr>
      <tr>
         <td>по &nbsp;<select name=day1  value="<%=Session("day")%>">
	 	 	 
	 	 	 <%if   Session("day")<>"" then%>
	 	 	 <option value="<%=Session("day")%>" selected><%=Session("day")%></option>
	 	 	 <%else%>
		  <option value="1" selected>1</option>
		  <%end if%>
		 <%for i=1 to 31%>
		 <option value="<%=i%>"><%=i%></option>
		 <%next%>
                 </select>	 	 	 
	 	 	 
	 	 <select name=month1  value="<%=Session("month")%>">
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
	 <select name=year1> 
	 
	     <%if   Session("year")<>"" then%>
		 <option value="<%=Session("year")%>" selected><%=Session("year")%></option>
		 
		 <%else%>
		  <option value="2006" selected>2006</option>
		  <%end if%>
		 <%for i=2006 to 2050%>
		 <option value="<%=i%>"><%=i%></option>
		 <%next%>
       </select>
</td>	 
         </tr>

	 <tr>
	 <td><input type="submit" value="&#1048;&#1089;&#1082;&#1072;&#1090;&#1100;"></td><td><input type="reset" value="&#1054;&#1095;&#1080;&#1089;&#1090;&#1080;&#1090;&#1100;"></td>
	 </tr></TR>
</table>

</form>

</td>
</table>

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