<%

on error resume next

set conn2 = server.CreateObject("Adodb.connection")
      
        connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
         Server.MapPath("missyou.mdb")
         conn2.Open connStr
      ' Recordset object
      
dim id

id =request("ID")
if left(id,1)="'" then
id=right(left(id,len(id)-1),len(id)-2 )
end if

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql1="select [SaleID],[ClientID],[datebeg],[dateend],[status],[Address1],[Address2],[a1],[a2],[a3],[a4],[a5],[a6],[a7],[a8],[a9],[a10],[a11],[a12],[a13],[a14],[a15]," 
sql1= sql1 &"[a16],[a17],[a18],[a19],[a20],[a21],[a22],[a23],[a24],[a25],[a26],[a27],[a28],[a29],[a30],[a31],[a32],[a33],[a34],[a35],[a36],[a37],[a38],[a39],[a40],[a41],[a42],[a43]"
sql1= sql1 & ",[a44],[a45],[a46],[a47],[a48],[a49],[a50],[a51],[a52],[a53],[a54],[a55],[a56],[a57],[a58],[a59],[a60],[a61],[a62],[a63],[a64],[a65],[a66],[a67],[a68],[a69],[a70],[a71],[a72],[a73]"
sql1= sql1 & ",[a74],[a75],[a76],[a77],[a78],[a79],[a80],[a81],[a82],[a83],[a84],[a85],[a86],[a87],[a88],[a89],[a90],[a91],[a92],[a93],[a94],[a95],[a96],[a97],[a98],[a99],[a100]"
sql1= sql1 & ",[a101],[a102],[a103],[a104],[a105],[a106],[a107],[a108],[a109],[a110],[a111],[a112],[a113],[a114],[a115],[a116],[a117],[a118],"
sql1= sql1 & "[b1],[b2],[b3],[b4],[b5],[b6],[b7],[b8],[b9],[b10],[b11],[b12],[b13],[b14],[b15],[b16],[b17],[b18],[b19],[b20],[b21],[b22],[b23],[b24],[b25],[b26],[b27],[b28],[b29],[b30],[b31],[b32],[b33]"
sql1= sql1 & ",[b34],[b35],[b36],[b37],[b38],[b39],[b40],[b41],[b42],[b43],[b44]"
sql1= sql1 & ",[c1],[c2],[c3],[c4],[c5],[c6],[c7],[c8],[c9],[c10],[c11],[c12],[c13],[c14],[c15],[c16],[c17],[c18],[c19],[c20],[c21],[c22],[c23],[c24],[c25],[c26],[c27],[c28],[c29],[c30],[c31],[c32],[c33],[NickName],[Comments] from InternClean " 
sql1= sql1 & " where ID = " & id & ""
       
        Set rs = conn2.Execute(sql1)

'response.write sql1        
        rs.MoveFirst

if rs("datebeg")<>"" then
Session("datebeg")=rs("datebeg")
else
Session("datebeg")=date()
end if

if rs("dateend")<>"" then
Session("dateend")=rs("dateend")
else
Session("dateend")=date()
end if




'response.write sql1

'response.write rs(0)
'response.write err.description

Session("Client")=rs("ClientID")
Session("NickName")=rs("NickName")
Session("Comments")=rs("Comments")
Session("datebeg")=rs("datebeg")
Session("dateend")=rs("dateend")
Session("Address")=rs("Address")
Session("Address1")=rs("Address1")
Session("Address2")=rs("Address2")
session("a1")=rs("a1")
session("a2")=rs("a2")
session("a3")=rs("a3")
session("a4")=rs("a4")
session("a5")=rs("a5")
session("a6")=rs("a6")
session("a7")=rs("a7")
session("a8")=rs("a8")
session("a9")=rs("a9")
session("a10")=rs("a10")
session("a11")=rs("a11")
session("a12")=rs("a12")
session("a13")=rs("a13")
session("a14")=rs("a14")
session("a15")=rs("a15")
session("a16")=rs("a16")
session("a17")=rs("a17")
session("a18")=rs("a18")
session("a19")=rs("a19")
session("a20")=rs("a20")
session("a21")=rs("a21")
session("a22")=rs("a22")
session("a23")=rs("a23")
session("a24")=rs("a24")
session("a25")=rs("a25")
session("a26")=rs("a26")
session("a27")=rs("a27")
session("a28")=rs("a28")
session("a29")=rs("a29")
session("a30")=rs("a30")
session("a31")=rs("a31")
session("a32")=rs("a32")
session("a33")=rs("a33")
session("a34")=rs("a34")
session("a35")=rs("a35")
session("a36")=rs("a36")
session("a37")=rs("a37")
session("a38")=rs("a38")
session("a39")=rs("a39")
session("a40")=rs("a40")
session("a41")=rs("a41")
session("a42")=rs("a42")
session("a43")=rs("a43")
session("a44")=rs("a44")
session("a45")=rs("a45")
session("a46")=rs("a46")
session("a47")=rs("a47")
session("a48")=rs("a48")
session("a49")=rs("a49")
session("a50")=rs("a50")
session("a51")=rs("a51")
session("a52")=rs("a52")
session("a53")=rs("a53")
session("a54")=rs("a54")
session("a55")=rs("a55")
session("a56")=rs("a56")
session("a57")=rs("a57")
session("a58")=rs("a58")
session("a59")=rs("a59")
session("a60")=rs("a60")
session("a61")=rs("a61")
session("a62")=rs("a62")
session("a63")=rs("a63")
session("a64")=rs("a64")
session("a65")=rs("a65")
session("a66")=rs("a66")
session("a67")=rs("a67")
session("a68")=rs("a68")
session("a69")=rs("a69")
session("a70")=rs("a70")
session("a71")=rs("a71")
session("a72")=rs("a72")
session("a73")=rs("a73")
session("a74")=rs("a74")
session("a75")=rs("a75")
session("a76")=rs("a76")
session("a77")=rs("a77")
session("a78")=rs("a78")
session("a79")=rs("a79")
session("a80")=rs("a80")
session("a81")=rs("a81")
session("a82")=rs("a82")
session("a83")=rs("a83")
session("a84")=rs("a84")
session("a85")=rs("a85")
session("a86")=rs("a86")
session("a87")=rs("a87")
session("a88")=rs("a88")
session("a89")=rs("a89")
session("a90")=rs("a90")
session("a91")=rs("a91")
session("a92")=rs("a92")
session("a93")=rs("a93")
session("a94")=rs("a94")
session("a95")=rs("a95")
session("a96")=rs("a96")
session("a97")=rs("a97")
session("a98")=rs("a98")
session("a99")=rs("a99")
session("a100")=rs("a100")
session("a101")=rs("a101")
session("a102")=rs("a102")
session("a103")=rs("a103")
session("a104")=rs("a104")
session("a105")=rs("a105")
session("a106")=rs("a106")
session("a107")=rs("a107")
session("a108")=rs("a108")
session("a109")=rs("a109")
session("a110")=rs("a110")
session("a111")=rs("a111")
session("a112")=rs("a112")
session("a113")=rs("a113")
session("a114")=rs("a114")
session("a115")=rs("a115")
session("a116")=rs("a116")
session("a117")=rs("a117")
session("a118")=rs("a118")

session("b1")=rs("b1")
session("b2")=rs("b2")
session("b3")=rs("b3")
session("b4")=rs("b4")
session("b5")=rs("b5")
session("b6")=rs("b6")
session("b7")=rs("b7")
session("b8")=rs("b8")
session("b9")=rs("b9")
session("b10")=rs("b10")
session("b11")=rs("b11")
session("b12")=rs("b12")
session("b13")=rs("b13")
session("b14")=rs("b14")
session("b15")=rs("b15")
session("b16")=rs("b16")
session("b17")=rs("b17")
session("b18")=rs("b18")
session("b19")=rs("b19")
session("b20")=rs("b20")
session("b21")=rs("b21")
session("b22")=rs("b22")
session("b23")=rs("b23")
session("b24")=rs("b24")
session("b25")=rs("b25")
session("b26")=rs("b26")
session("b27")=rs("b27")
session("b28")=rs("b28")
session("b29")=rs("b29")
session("b30")=rs("b30")
session("b31")=rs("b31")
session("b32")=rs("b32")
session("b33")=rs("b33")
session("b34")=rs("b34")
session("b35")=rs("b35")
session("b36")=rs("b36")
session("b37")=rs("b37")
session("b38")=rs("b38")
session("b39")=rs("b39")
session("b40")=rs("b40")
session("b41")=rs("b41")
session("b42")=rs("b42")
session("b43")=rs("b43")
session("b44")=rs("b44")

session("c1")=rs("c1")
session("c2")=rs("c2")
session("c3")=rs("c3")
session("c4")=rs("c4")
session("c5")=rs("c5")
session("c6")=rs("c6")
session("c7")=rs("c7")
session("c8")=rs("c8")
session("c9")=rs("c9")
session("c10")=rs("c10")
session("c11")=rs("c11")
session("c12")=rs("c12")
session("c13")=rs("c13")
session("c14")=rs("c14")
session("c15")=rs("c15")
session("c16")=rs("c16")
session("c17")=rs("c17")
session("c18")=rs("c18")
session("c19")=rs("c19")
session("c20")=rs("c20")
session("c21")=rs("c21")
session("c22")=rs("c22")
session("c23")=rs("c23")
session("c24")=rs("c24")
session("c25")=rs("c25")
session("c26")=rs("c26")
session("c27")=rs("c27")
session("c28")=rs("c28")
session("c29")=rs("c29")
session("c30")=rs("c30")
session("c31")=rs("c31")
session("c32")=rs("c32")
session("c33")=rs("c33")

'response.write session("a1")

Set rs = Server.CreateObject("ADODB.Recordset")

sql1 = "select [ClientID],[ClientName],[Address],[Phone] from Client "

Set rs = conn2.Execute(sql1)
rs.MoveFirst


Set rs1 = Server.CreateObject("ADODB.Recordset")

sql1 = "select [ContactPerson] from Client where ClientID=" & Session("Client") & "" 

'response.write sql1

Set rs1 = conn2.Execute(sql1)
rs1.MoveFirst


Set rs2 = Server.CreateObject("ADODB.Recordset")

sql1 = "select [ClientID],[ClientName],[Address],[Phone] from Client where ClientID = " & Session("Client") & ""
Set rs2 = conn2.Execute(sql1)
rs2.MoveFirst

'response.write sql1

Session("Contact") = rs1("ContactPerson")
Session("ClientName") = rs2("ClientName")
Session("Address") = rs2("Address")



%>
<HTML>
  <HEAD>
  
    <title>Default page</title>
     <link rel="stylesheet" type="text/css" href="index.css">
<META content="text/html; charset=windows-1251" http-equiv=Content-Type> 
<META NAME="Keywords" CONTENT="">

  </HEAD>
  <BODY background="bg.jpg" alink="#0000ff" vlink="#0000ff" link="#0000ff">


        
<center>
<img src="relationshipsromance.jpeg"></img>
</center>


<br>
<br>
<br>
<table>

<tr>
<td><font size=3>Название проекта: <%=Session("pr")%></font></td>
</tr>

<tr>
<td><font size=3>Инициатор: <%=Session("login")%></font></td>
</tr>
<tr>
 <table>
  
 <td valign="top">
<a href="common.asp?ID=<%=Session("ID")%>"><font  size=1>Общие данные(редактировать)</font></a>
		<hr>
		<a href="vnutr.asp?ID=<%=Session("ID")%>"><font  size=1>Заявка (редактировать)</font></a>
		<hr>
		<a href="spec.asp?ID=<%=Session("ID")%>"><font  size=1>ДОПОЛНИТЕЛЬНЫЕ УСЛУГИ (редактировать)</font></a>
		<hr>
                <a href="default1_0_nonedit.asp?ID=<%=Session("ID")%>"><font  size=1>Заявка (Вид для печати)</font></a>
		<hr> 
                               <a href="uploadform.asp"><font  size=1>Загрузить файл</font></a>
				
				<hr> 
                               <a href="show.asp"><font  size=1>Показать все заявки</font></a>
				<hr>

 </td>
<td>


<form method=get action='save5.asp'>

<!---------
<table border=0 width=100%>
<tr><td align=left>ВНИМАНИЕ: Заполнению подлежат ячейки, выделенные цветом. Данные, полученные в ходе исследования строго конфиденциальны!
</td></tr>
<tr><td align=left>1.<b><u>ВНУТРЕННЯЯ УБОРКА ПОМЕЩЕНИЙ:</u></b></u></b></td></tr>
</table>
<br>


----->
<br>
Предоставьте, пожалуйста, общую информацию

<table border=1>
<tr>
<td>Клиент:</td><td>	<select name=d1  value="<%=rs("ClientName")%>">
                 <%dim j

                   j=1 
                  %> 
                <option value="<%=Session("Client")%>" selected ><%=Session("ClientName") %></option>
                 <% while not rs.eof %>
		 <option value="<%=j%>"><%=rs("ClientName") %></option>
                 <% rs.MoveNext %>
                 <% j = j+1 %>
		 <% wend %>
                 </select></td>
</tr>
<tr>
<td>Срок исполнения:</td><td>
<input type=text name=d2 value='<%=Session("dateend")%>' size="20"></td>
</tr>
<tr>
<td>Project NickName (Название проекта):</td><td>
<input type=text name=d3 value='<%=Session("NickName")%>' size="20"></td>
</tr>
<tr>
<td>Контактное лицо:</td><td>
<input type=text name=d4 value='<%=Session("Contact")%>' size="20"></td>
</tr>
<tr>
<td>Начало работ:</td>
<td>
<!------
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
		  <option value="2007" selected>2007</option>
		  <%end if%>
		 <%for i=2007 to 2050%>
		 <option value="<%=i%>"><%=i%></option>
		 <%next%>
       </select>
</td>
---->
<input type=text name=d2 value='<%=Session("datebegin")%>' size="20"></td>
</tr>
<tr>
<%rs.MoveFirst%>
<td>Адрес клиента:</td><td><select name=d6  value="<%=rs("Address")%>">
<option value="<%=Session("Address")%>" selected><%=Session("Address") %></option>
                  
                 <% while not rs.eof %>
		 <option value="<%=rs("Address")%>"><%=rs("Address") %></option>
                 <% rs.MoveNext %>
		 <% wend %>
                 </select></td>
</tr>
<tr>
<td>Место откуда:</td><td><select name=d8  value="<%=rs("Address")%>">
<option value="<%=Session("Address1")%>" selected><%=Session("Address") %></option>
<option value="Днепропетровск" selected>Днепропетровск</option>
<option value="Винница" >Винница</option>
<option value="Луцк">Луцк</option>
<option value="Горловка (Донецкая обл)">Горловка (Донецкая обл)</option>
<option value="Донецк">Донецк</option>
<option value="Артемовск (Донецкая обл)">Артемовск (Донецкая обл)</option>
<option value="Краматорск (Донецкая обл)">Краматорск (Донецкая обл)</option>
<option value="Мариуполь (Донецкая обл)">Мариуполь (Донецкая обл)</option>
<option value="Бердичев (Житомирская обл)">Бердичев (Житомирская обл)</option>
<option value="Житомир">Житомир</option>
<option value="Мелитополь (Запорожская обл)">Мелитополь (Запорожская обл)</option>
<option value="Запорожье">Запорожье</option>
<option value="Ивано-Франковск">Ивано-Франковск</option>
<option value="Киев">Киев</option>
<option value="Симферополь">Симферополь</option>
<option value="Луганск">Луганск</option>
<option value="Львов">Львов</option>
<option value="Николаев">Николаев</option>
<option value="Одесса">Одесса</option>
<option value="Полтава">Полтава</option>
<option value="Ровно">Ровно</option>
<option value="Сумы">Сумы</option>
<option value="Тернополь">Тернополь</option>
<option value="Харьков">Харьков</option>
<option value="Хмельницкий">Хмельницкий</option>
<option value="Черкассы">Черкассы</option>
<option value="Чернигов">Чернигов</option>
<option value="Черновцы">Черновцы</option>
</select></td>
</tr>
<td>Место куда:</td><td><select name=d9  value="<%=rs("Address2")%>">
<option value="<%=Session("Address2")%>" selected><%=Session("Address2") %></option>
<option value="Винница" >Винница</option>
<option value="Луцк">Луцк</option>
<option value="Горловка (Донецкая обл)">Горловка (Донецкая обл)</option>
<option value="Донецк">Донецк</option>
<option value="Артемовск (Донецкая обл)">Артемовск (Донецкая обл)</option>
<option value="Краматорск (Донецкая обл)">Краматорск (Донецкая обл)</option>
<option value="Мариуполь (Донецкая обл)">Мариуполь (Донецкая обл)</option>
<option value="Бердичев (Житомирская обл)">Бердичев (Житомирская обл)</option>
<option value="Житомир">Житомир</option>
<option value="Мелитополь (Запорожская обл)">Мелитополь (Запорожская обл)</option>
<option value="Запорожье">Запорожье</option>
<option value="Днепропетровск">Днепропетровск</option>
<option value="Ивано-Франковск">Ивано-Франковск</option>
<option value="Киев">Киев</option>
<option value="Симферополь">Симферополь</option>
<option value="Луганск">Луганск</option>
<option value="Львов">Львов</option>
<option value="Николаев">Николаев</option>
<option value="Одесса">Одесса</option>
<option value="Полтава">Полтава</option>
<option value="Ровно">Ровно</option>
<option value="Сумы">Сумы</option>
<option value="Тернополь">Тернополь</option>
<option value="Харьков">Харьков</option>
<option value="Хмельницкий">Хмельницкий</option>
<option value="Черкассы">Черкассы</option>
<option value="Чернигов">Чернигов</option>
<option value="Черновцы">Черновцы</option>
</select></td>
               
                 
</tr>




<tr>
<td>Комментарии:</td><td><textarea name=d7 rows="5" cols="30" size = 20><%=Session("Comments")%></textarea></td>
</tr>

<tr>

</tr>

<tr>

</tr>
<tr>

</tr>

<tr>

</tr>
<tr>

</tr>

<tr>

</tr>
<tr>

</tr>

<tr>

</tr>
<tr>

</tr>

<tr>
<td></td><td></td><td><input type=submit value="Сохранить"></td>
</tr>
</table>
</align>
</form>
<center>
<table border=0 >
<td>

</table>
</center>
</tr>
</table>
</center>


</td>
</tr>
		
</table>		
</td>
</tr>
</center>		

</center>
</FONT>
<br>
<br>
<br>
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