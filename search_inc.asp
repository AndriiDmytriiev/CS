
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
      set rs1=conn.Execute("select distinct [City] from [Persons]")
    

%>



<form name=input action=show.asp id="input" method="get">

<table border=1>

<tr>
<table>
	 <td><FONT  color="#0000ff" style="FONT-WEIGHT: bold">&#1048;&#1084;&#1103;:</font>  </td><td><input name=personname ></td>
	</tr>
	<tr> 
	 <td><FONT  color="#0000ff" style="FONT-WEIGHT: bold">&#1053;&#1080;&#1082;:</font>  </td><td><input name=nick ></td>
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
	 <td><FONT  color="#0000ff" style="FONT-WEIGHT: bold">&#1055;&#1086;&#1083;:</font></td><td>
	 <select name="Sex"><option value="m" >&#1052;&#1091;&#1078;&#1089;&#1082;&#1086;&#1081;</option>
						<option value="f" selected >&#1046;&#1077;&#1085;&#1089;&#1082;&#1080;&#1081;</option>
	 </select></td>
     </tr> 
     <tr>
	 <td><FONT  color="#0000ff" style="FONT-WEIGHT: bold">&#1042;&#1086;&#1079;&#1088;&#1072;&#1089;&#1090;:</font></td><td> <FONT  color="#0000ff" style="FONT-WEIGHT: bold">&#1086;&#1090;</font> <input name=age1 width="10" style="WIDTH: 29px; HEIGHT: 22px" size=3 
     > <FONT  color="#0000ff" style="FONT-WEIGHT: bold">&#1076;&#1086;</font> <input name=age2 width="10" style="WIDTH: 

29px; HEIGHT: 22px" size=3 
     ></td>
	 </tr>
<tr>
	 <td><FONT  color="#0000ff" style="FONT-WEIGHT: bold">Рост:</font>  </td><td><input name=height ></td>
	</tr>
<tr>
	 <td><FONT  color="#0000ff" style="FONT-WEIGHT: bold">Вес:</font>  </td><td><input name=weight ></td>
	</tr>

	 <tr>
	 <td><FONT  color="#0000ff" style="FONT-WEIGHT: bold"></font></td><td></td>
	 </tr>
	 <tr>
	 <td><FONT  color="#0000ff" style="FONT-WEIGHT: bold">&#1045;&#1089;&#1090;&#1100; &#1092;&#1086;&#1090;&#1086;:</font></td><td><input name=photo 

type=checkbox></td>
	 </tr>
	 
	 
	 <tr>
	 <td><input type="submit" value="&#1048;&#1089;&#1082;&#1072;&#1090;&#1100;"></td><td><input type="reset" value="&#1054;&#1095;&#1080;&#1089;&#1090;&#1080;&#1090;&#1100;"></td>
	 </tr></TR>
</table>
</form>
</td>
</table>

