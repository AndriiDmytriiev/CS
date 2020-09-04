<%
on error resume next
set conn2 = server.CreateObject("Adodb.connection")
      
        connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
         Server.MapPath("missyou.mdb")
         conn2.Open connStr
      ' Recordset object
      
        
        
'-Сохраняем данные из первой формы---------
session("a1")=request.querystring("t1")
session("a2")=request.querystring("ch1")
session("a3")=request.querystring("Period1")
session("a4")=request.querystring("t2")
session("a5")=request.querystring("ch2")
session("a6")=request.querystring("Period2")
session("a7")=request.querystring("t3")
session("a8")=request.querystring("ch3")
session("a9")=request.querystring("Period3")
session("a10")=request.querystring("t4")
session("a11")=request.querystring("ch4")
session("a12")=request.querystring("Period4")
session("a13")=request.querystring("t5")
session("a14")=request.querystring("ch5")
session("a15")=request.querystring("Period5")
session("a16")=request.querystring("t6_1")
session("a17")=request.querystring("ch6")
session("a18")=request.querystring("Period6")
session("a19")=request.querystring("t6_2")
session("a20")=request.querystring("t7_1")
session("a21")=request.querystring("ch7")

session("a22")=request.querystring("Period7")

session("a23")=request.querystring("t7_2")
session("a24")=request.querystring("t8_1")
session("a25")=request.querystring("ch8")
session("a26")=request.querystring("Period8")

session("a27")=request.querystring("t8_2")
session("a28")=request.querystring("t9_1")
session("a29")=request.querystring("ch9")
session("a30")=request.querystring("Period9")
session("a31")=request.querystring("t9_2")
session("a32")=request.querystring("t10_1")
session("a33")=request.querystring("ch10")
session("a34")=request.querystring("Period10")
session("a35")=request.querystring("t10_2")
session("a36")=request.querystring("t11_1")
session("a37")=request.querystring("ch11")
session("a38")=request.querystring("Period11")
session("a39")=request.querystring("t11_2")
session("a40")=request.querystring("t12_1")
session("a41")=request.querystring("ch12")
session("a42")=request.querystring("Period12")
session("a43")=request.querystring("t12_2")
session("a44")=request.querystring("t13_1")
session("a45")=request.querystring("ch13")

session("a46")=request.querystring("Period13")

session("a47")=request.querystring("t13_2")
session("a48")=request.querystring("t14_1")
session("a49")=request.querystring("ch14")
session("a50")=request.querystring("Period14")
session("a51")=request.querystring("t14_2")
session("a52")=request.querystring("t15_1")
session("a53")=request.querystring("ch15")
session("a54")=request.querystring("Period15")
session("a55")=request.querystring("t15_2")
session("a56")=request.querystring("t16_1")
session("a57")=request.querystring("ch16")
session("a58")=request.querystring("Period16")
session("a59")=request.querystring("t16_2")
session("a60")=request.querystring("t17_0")
session("a61")=request.querystring("t17_1")
session("a62")=request.querystring("ch17")

session("a63")=request.querystring("Period7")

session("a64")=request.querystring("t17_2")
session("a65")=request.querystring("t18")
session("a66")=request.querystring("ch18")
session("a67")=request.querystring("Period19")
session("a68")=request.querystring("t19_1")
session("a69")=request.querystring("ch19_1")
session("a70")=request.querystring("t19_2")
session("a71")=request.querystring("ch19_2")
session("a72")=request.querystring("t19_3")
session("a73")=request.querystring("ch19_3")
session("a74")=request.querystring("t19_4")
session("a75")=request.querystring("ch19_4")
session("a76")=request.querystring("ch20")
session("a77")=request.querystring("Period20")
session("a78")=request.querystring("t20_1")
session("a79")=request.querystring("ch20_1")
session("a80")=request.querystring("t20_2")
session("a81")=request.querystring("ch20_2")
session("a82")=request.querystring("t20_3")
session("a83")=request.querystring("ch20_3")
session("a84")=request.querystring("t20_4")
session("a85")=request.querystring("ch20_4")
session("a86")=request.querystring("t20_5")
session("a87")=request.querystring("ch20_5")
session("a88")=request.querystring("t20_6")
session("a89")=request.querystring("ch20_6")
session("a90")=request.querystring("t20_7")
session("a91")=request.querystring("ch20_7")
session("a92")=request.querystring("t20_8")
session("a93")=request.querystring("ch20_8")
session("a94")=request.querystring("t20_9")
session("a95")=request.querystring("ch20_9")
session("a96")=request.querystring("t20_10")
session("a97")=request.querystring("ch20_10")
session("a98")=request.querystring("t20_11")
session("a99")=request.querystring("ch20_11")
session("a100")=request.querystring("ch21")
session("a101")=request.querystring("Period21")
session("a102")=request.querystring("t22")
session("a103")=request.querystring("t23")
session("a104")=request.querystring("t24")
session("a105")=request.querystring("t25")
session("a106")=request.querystring("t26")
session("a107")=request.querystring("t27")
session("a108")=request.querystring("t28")
session("a109")=request.querystring("ch22")
session("a110")=request.querystring("Period22")
session("a111")=request.querystring("choice1")
session("a112")=request.querystring("choice2")
session("a113")=request.querystring("t29")
session("a114")=request.querystring("ch23")
session("a115")=request.querystring("Period23")
session("a116")=request.querystring("choice3")
session("a117")=request.querystring("t30")
session("a118")=request.querystring("t31")

'------------------------------------------


        
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
<%if Session("login")<>"" then%>
<%
'Session("day")=Request.Form("day")
'        Session("month")=Request.Form("month")
'        Session("year")=Request.Form("year")
%>


<table>
<tr>
<td>Инициатор:<%=Session("Login")%></td><td>|</td><td>Клиент:<%=Session("client")%></td>
</tr>
<tr>
<td>Дата:<%=date()%></td><td>|</td><td>Контактное лицо, тел:<%=Session("contact")%></td>
</tr>
<tr>
<td>Срок ответа:<%=Session("day") & "." & Session("month") & "." & Session("year") %></td><td>|</td><td>Дата начала работ:<%=date()%></td>
</tr>
</table>

<%end if%>

<td>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

</td>
<td>
<table>
<tr>
<td align=right valign=top>




<form method=get action='default_next2.asp'>
<!--------------
<table border=0 width=100%>
<tr><td align=left>ВНИМАНИЕ: Заполнению подлежат ячейки, выделенные цветом. Данные, полученные в ходе исследования строго конфиденциальны!
</td></tr>
<tr><td align=left>1.<b><u>ВНУТРЕННЯЯ УБОРКА ПОМЕЩЕНИЙ:</u></b></u></b></td></tr>
</table>
<br>


Предоставьте, пожалуйста, информацию по всем зданиям, в которых необходимо осуществить уборку.



<table border=1>
<tr>
<td>&nbsp;</td><td>Показатель</td><td>Ед.изм.</td><td>Здание № 1</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>1.</td><td><b>Краткие данные о здании/объекте</b></td><td></td><td>офисы</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>1.1</td><td>этажность</td><td>эт.</td><td>
<input type=textbox name=t1 size="20"></td><td>
<input type=checkbox name = ch1 ></td><td>
<select name="Period1">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>1.2</td><td>год последнего ремонта</td><td>год</td><td>
<input type=textbox name=t2 size="20"></td><td>
<input type=checkbox name = ch2 ></td><td>
<select name="Period2">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>1.3</td><td>высота потолков</td><td>м.</td><td>
<input type=textbox name=t3 size="20"></td><td>
<input type=checkbox name = ch3 ></td><td>
<select name="Period3">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>2.</td><td><b>Общая площадь</b></td><td>кв.м</td><td>
<input type=textbox name=t4 size="20"></td><td>
<input type=checkbox name = ch4 ></td><td>
<select name="Period4">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>2.1</td><td>Площадь каждого этажа</td><td>кв.м</td><td>
<input type=textbox name=t5 size="20"></td><td>
<input type=checkbox name = ch5 ></td><td>
<select name="Period5">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>2.2</td><td>Кабинеты VIP</td><td>кол-во</td><td>
<input type=textbox name=t6_1 size="20"></td><td>
<input type=checkbox name = ch6 ></td><td>
<select name="Period6">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t6_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.3</td><td>Офисные помещения</td><td>кол-во</td><td>
<input type=textbox name=t7_1 size="20"></td><td>
<input type=checkbox name = ch7 ></td><td>
<select name="Period7">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t7_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.4</td><td>Складские помещения</td><td>кол-во</td><td>
<input type=textbox name=t8_1 size="20"></td><td>
<input type=checkbox name = ch8 ></td><td>
<select name="Period8">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t8_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.5</td><td>Складские помещения</td><td>кол-во</td><td>
<input type=textbox name=t9_1 size="20"></td><td>
<input type=checkbox name = ch9 ></td><td>
<select name="Period9">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t9_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.6</td><td>Технические помещения, подвалы</td><td>кол-во</td><td>
<input type=textbox name=t10_1 size="20"></td><td>
<input type=checkbox name = ch10 ></td><td>
<select name="Period10">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t10_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>2.7</td><td>Коридоры</td><td>кол-во</td><td>
<input type=textbox name=t11_1 size="20"></td><td>
<input type=checkbox name = ch11 ></td><td>
<select name="Period11">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t11_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.8</td><td>Лестницы</td><td>кол-во</td><td>
<input type=textbox name=t12_1 size="20"></td><td>
<input type=checkbox name = ch12 ></td><td>
<select name="Period12">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t12_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.9</td><td>Лифты</td><td>кол-во</td><td>
<input type=textbox name=t13_1 size="20"></td><td>
<input type=checkbox name = ch13 ></td><td>
<select name="Period13">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t13_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.10</td><td>Эскалаторы</td><td>кол-во</td><td>
<input type=textbox name=t14_1 size="20"></td><td>
<input type=checkbox name = ch14 ></td><td>
<select name="Period14">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t14_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.11</td><td>Санузлы</td><td>кол-во</td><td>
<input type=textbox name=t15_1 size="20"></td><td>
<input type=checkbox name = ch15 ></td><td>
<select name="Period15">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t15_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.12</td><td>холлы, вестибюли</td><td>кол-во</td><td>
<input type=textbox name=t16_1 size="20"></td><td>
<input type=checkbox name = ch16 ></td><td>
<select name="Period16">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t16_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>2.13</td><td>Другие площади (по возм.расшифровать)</td><td></td><td>
<input type=textbox name=t17_0 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>кол-во</td><td>
<input type=textbox name=t17_1 size="20"></td><td>
<input type=checkbox name = ch17 ></td><td>
<select name="Period17">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>площадь, кв.м</td><td>
<input type=textbox name=t17_2 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>3.</td><td><b>К-во сотрудников Вашей компании и посетителей, чел. в среднем в день</b></td><td>чел.</td><td>
<input type=textbox name=t18 size="20"></td><td>
<input type=checkbox name = ch18 ></td><td>&nbsp;</td>
</tr>


<tr>
<td>4.</td><td><b>Необходимость в обеспечении с/у расходными материалами (примерный расход в месяц)</b></td><td>кол-во</td><td>&nbsp;</td><td>&nbsp;</td><td>
<select name="Period19">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>&nbsp;</td><td>туалетная бумага</td><td>рул./мес.</td><td>
<input type=textbox name=t19_1 size="20"></td><td>
<input type=checkbox name = ch19_1 ></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>жидкое мыло</td><td>литр/мес.</td><td>
<input type=textbox name=t19_2 size="20"></td><td>
<input type=checkbox name = ch19_2 ></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>бум полотенца</td><td>лист/мес.</td><td>
<input type=textbox name=t19_3 size="20"></td><td>
<input type=checkbox name = ch19_3 ></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>бумажные сидения д/унитаза</td><td>шт./мес.</td><td>
<input type=textbox name=t19_4 size="20"></td><td>
<input type=checkbox name = ch19_4 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.</td><td><b>Поверхности</b></td><td>чел.</td><td>&nbsp;</td><td>
<input type=checkbox name = ch20 ></td><td>
<select name="Period20">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>

<tr>
<td>5.1</td><td>Мягкие покрытия (ковролин)</td><td>кв.м</td><td>
<input type=textbox name=t20_1 size="20"></td><td>
<input type=checkbox name = ch20_1 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.2</td><td>Полутвердые покрытия (линолеум, паркет, мармолеум, наливной пол, ламинат)</td><td>кв.м</td><td>
<input type=textbox name=t20_2 size="20"></td><td>
<input type=checkbox name = ch20_2 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.3</td><td>Твердые покрытия (плитка, мрамор, гранит) пол/стены</td><td>кв.м</td><td>
<input type=textbox name=t20_3 size="20"></td><td>
<input type=checkbox name = ch20_3 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.4</td><td>Стеклянные поверхности</td><td>кв.м</td><td>
<input type=textbox name=t20_4 size="20"></td><td>
<input type=checkbox name = ch20_4 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.5</td><td>Металлические поверхности</td><td>кв.м</td><td>
<input type=textbox name=t20_5 size="20"></td><td>
<input type=checkbox name = ch20_5 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.6</td><td>Офисные перегородки</td><td>шт.</td><td>
<input type=textbox name=t20_6 size="20"></td><td>
<input type=checkbox name = ch20_6 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.7</td><td>Офисные места</td><td>шт.</td><td>
<input type=textbox name=t20_7 size="20"></td><td>
<input type=checkbox name = ch20_7 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.8</td><td>Кожаная мебель</td><td>шт.</td><td>
<input type=textbox name=t20_8 size="20"></td><td>
<input type=checkbox name = ch20_8 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.9</td><td>Пластиковая мебель</td><td>шт.</td><td>
<input type=textbox name=t20_9 size="20"></td><td>
<input type=checkbox name = ch20_9 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.10</td><td>Деревянная мебель</td><td>шт.</td><td>
<input type=textbox name=t20_10 size="20"></td><td>
<input type=checkbox name = ch20_10 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>5.11</td><td>Другие поверхности</td><td>шт./кв.м</td><td>
<input type=textbox name=t20_11 size="20"></td><td>
<input type=checkbox name = ch20_11 ></td><td>&nbsp;</td>
</tr>

<tr>
<td>6.</td><td><b>Существующая у Вас уборка</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch21 ></td><td>
<select name="Period21">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>


<tr>
<td>6.1</td><td>График проведения основной комплексной уборки</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>с:</td><td>час</td><td>
<input type=textbox name = t22 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>до:</td><td>час</td><td>
<input type=textbox name = t23 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>дней в году:</td><td>дни</td><td>
<input type=textbox name = t24 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>6.2</td><td>График проведения поддерживающей уборки</td><td></td><td></td><td></td><td></td>
</tr>

<tr>
<td>&nbsp;</td><td>с:</td><td>час</td><td>
<input type=textbox name = t25 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>до:</td><td>час</td><td>
<input type=textbox name = t26 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>дней в году:</td><td>дни</td><td>
<input type=textbox name = t27 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>



<tr>
<td>6.3</td><td>Количество уборщиков</td><td>чел.</td><td>
<input type=textbox name = t28 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.</td><td><b>Помещения для размещения производственного персонала и оборудования</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch22 ></td><td>
<select name="Period22">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>

<tr>
<td>7.1</td><td>Наличие помещения</td><td>да/нет</td><td><select name="choice1">
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.2</td><td>Наличие телефона в помещении</td><td>да/нет</td><td><select name="choice2">
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.3</td><td>площадь помещения</td><td>кв.м</td><td>
<input type=textbox name = t29 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>8.</td><td><b>Помещения для размещения управленческого персонала и автоматизированных рабочих мест (для объектов более 10000 кв.м)</b></td><td>&nbsp;</td><td>&nbsp;</td><td>
<input type=checkbox name = ch23 ></td><td>
<select name="Period23">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>

<tr>
<td>8.1</td><td>Наличие помещения</td><td>да/нет</td><td><select name="choice3">
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td></td><td></td>
</tr>
<tr>
<td>8.2</td><td>Площадь помещения</td><td>кв.м</td><td>
<input type=textbox name = t30 size="20"></td><td></td><td></td>
</tr>
<tr>
<td>8.3</td><td>кол-во автоматизированных рабочих мест</td><td>шт.</td><td><input type=textbox name = t31 size="20"></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>



</table>
------------->
<br>

<table border=0 width=100%>
<tr><td align=left>ВНИМАНИЕ: Заполнению подлежат ячейки, выделенные цветом. Данные, полученные в ходе исследования строго конфиденциальны!
</td></tr>
<tr><td align=left>2.<b><u>УБОРКА ТЕРРИТОРИИ:</u></b></td></tr>
</table>


<br>

Предоставьте, пожалуйста, информацию по всем зданиям, в которых необходимо осуществить уборку.
<br>
<table border=1>
<tr>
<td>&nbsp;</td><td>Показатель</td><td>Ед.изм.</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>1.</td><td><b>Общая убираемая территория</b></td><td>кв.м</td><td><input type=textbox name=b1></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>1.1</td><td>Площади с указанием покрытия (тротуары, стоянки, газоны, проезды, другое)</td><td>&nbsp;</td><td><input type=textbox name=b2></td><td><input type=checkbox name = k1></td><td>
<select name="b3">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>1.1.1</td><td>Тротуары</td><td>год</td><td><input type=textbox name=b4></td><td><input type=checkbox name = k2></td><td>
<select name="b5">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>1.1.2</td><td>Стоянки</td><td>м.</td><td><input type=textbox name=b6></td><td><input type=checkbox name = k3></td><td>
<select name="b7">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>1.1.3</td><td>Газоны</td><td>м.</td><td><input type=textbox name=b8></td><td><input type=checkbox name = k4></td><td>
<select name="b9">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>1.1.4</td><td>Проезды</td><td>м.</td><td><input type=textbox name=b10></td><td><input type=checkbox name = k5></td><td>
<select name="b11">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>2.</td><td><b>Существующая у Вас уборка территории</b></td><td>кв.м</td><td><input type=textbox name=b12></td><td><input type=checkbox name = k6></td><td>
<select name="b13">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>

<tr>
<td>2.1</td><td>График уборки</td><td></td><td></td><td></td><td></td>
</tr>

<tr>
<td>&nbsp;</td><td>с:</td><td>час</td><td><input type=textbox name = b14></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>до:</td><td>час</td><td><input type=textbox name = b15></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>дней в году:</td><td>дни</td><td><input type=textbox name = b16></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>2.2</td><td>Количество уборщиков</td><td>чел.</td><td><input type=textbox name = b17></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>


<tr>
<td>3.</td><td><b>Инвентаризационная ведомость техники (оборудования) на</b></td><td>чел.</td><td><input type=textbox name=b18></td><td><input type=checkbox name = k7></td><td>&nbsp;</td>
</tr>
<tr>
<td>3.1</td><td>Офисные помещения</td><td>мес</td><td><input type=textbox name=b19></td><td><input type=checkbox name = k8></td><td>
<select name="b20">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>3.2</td><td>Офисные помещения</td><td>мес</td><td><input type=textbox name=b21></td><td><input type=checkbox name = k9></td><td>
<select name="b22">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>3.3</td><td>Офисные помещения</td><td>мес</td><td><input type=textbox name=b23></td><td><input type=checkbox name = k10></td><td>
<select name="b24">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>
<tr>
<td>3.4</td><td>Офисные помещения</td><td>мес</td><td><input type=textbox name=b25></td><td><input type=checkbox name = k11></td><td>
<select name="b26">
<option value="разовая" selected>разовая</option>
<option value="часто" >часто</option>
<option value="раз в год" >раз в год</option>
</select>
</td>
</tr>

<tr>
<td>4.</td><td><b>Наличие места для стоянки уборочной техники,</b> оборудованного отоплением и складом для хранения расходных материалов и запчастей, а также проведения ремонта техники, оборудования и инвентаря</td><td>да/нет</td><td><select name="b27">
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>5.</td><td><b>Потребность в стрижке газона, покосе травы (площадь и кол-во раз)</b></td><td>да/нет</td><td><select name="b28">
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>5.1</td><td>площадь</td><td>кв.м</td><td><input type=textbox name=b29></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td></td><td>количество стрижек газонов и покосов травы</td><td>кол-во</td><td><input type=textbox name=b30></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>6.</td><td><b>Качество уборки территории зимой </b>(A - до утрамбованного слоя снега/ Б - до покрытия)</td><td>А/Б</td><td><select name="b31">
<option value="1" selected>А</option>
<option value="2" >Б</option>
</select></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

<tr>
<td>7.</td><td><b>Вывоз снега в зимний период (объем вывоза за зиму)</b></td><td>куб.м.</td><td><input type=textbox name=b32></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>

<tr>
<td>8.</td><td><b>Вывоз ТБО, примерный месячный объем мусора</b></td><td>куб.м.</td><td><input type=textbox name=b33></td><td>&nbsp;</td><td>&nbsp;</td>
</tr>

</table>

<!-------------------

<align=left>
<br>
<table border=0 width=100%>
<tr><td align=left>ВНИМАНИЕ: Заполнению подлежат ячейки, выделенные цветом. Данные, полученные в ходе исследования строго конфиденциальны!</td></tr>
<tr><td align=left>3.<b><u>ДОПОЛНИТЕЛЬНЫЕ УСЛУГИ:</u></b></td></tr>
</table>
<table border=1 width=100%>
<tr>
<td>&nbsp;</td><td>Показатель</td><td>Ед.изм.</td><td>&nbsp;</td><td>&nbsp;</td>
</tr>
<tr>
<td><b>1.</b></td><td><b>Организация работы прачечных</b></td><td>да/нет</td><td><select name="c1">
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td><input type=checkbox name = c2></td>
</tr>
<tr>
<td>1.1</td><td>Наличие химчистки-прачечной</td><td>да/нет</td><td><select name="c3">
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td><input type=checkbox name = c4></td>
</tr>
<tr>
<td>1.1.1</td><td>количество обслуживающего персонала</td><td>чел.</td><td><input type=text name = c5></td><td><input type=checkbox name = c6></td>
</tr>
<tr>
<td>1.2</td><td>Потребность в стирке спецодежды, объем в месяц</td><td>кг</td><td><input type=text name = c7></td><td><input type=checkbox name = c8></td>
</tr>
<tr>
<td><b>2.</b></td><td><b>Глубокая химическая чистка ковровых покрытий</b></td><td>да/нет</td><td><select name="c9">
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td><input type=checkbox name = c10></td>
</tr>
<tr>
<td>2.1</td><td>Площадь ковровых покрытий</td><td>кв.м</td><td><input type=text name = c11></td><td><input type=checkbox name = c12></td>
</tr>
<tr>
<td>2.2</td><td>Периодичность проведения</td><td>&nbsp;</td><td><input type=text name = c13></td> <td>&nbsp;</td>
</tr>
<tr>
<td><b>3.</b></td><td><b>Нанесение полимерного лака на линолеумные полы</b></td><td>да/нет</td><td><select name="c14">
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td><input type=checkbox name = c15></td>
</tr>
<tr>
<td>3.1</td><td>Площадь линолеума</td> <td>кв.м</td><td><input type=text name = c16></td>
</tr>
<tr>
<td>3.2</td><td>Периодичность нанесения лака</td> <td>&nbsp;</td><td><input type=text name = c16></td><td>&nbsp;</td>
</tr>
<tr>
<td><b>4.</b></td><td><b>Мытье окон (площадь окон с одной стороны)</b></td> <td>да/нет</td><td><select name="c17">
<option value="да" selected>да</option>
<option value="нет" >нет</option>
</select></td><td><input type=checkbox name = c18></td>
</tr>
<tr>
<td>4.1</td><td>Легкий доступ к окнам (с пола)</td> <td>кв.м</td><td><input type=text name = c19></td><td><input type=checkbox name = c20></td>
</tr>
<tr>
<td>4.2</td><td>Затрудненный (со стремянки)</td> <td>кв.м</td><td><input type=text name = c21></td><td><input type=checkbox name = c22></td>
</tr>
<tr>
<td>4.3</td><td> С помощью промышленных альпинистов </td> <td>кв.м</td><td><input type=text name = c23></td><td><input type=checkbox name = c24></td>
</tr>
<tr>
<td><b>5.</b></td><td><b>Другие услуги, которые вы хотели бы получать</b></td><td>&nbsp;</td> <td>&nbsp;</td><td><input type=checkbox name = c25></td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td><input type=text name = c26></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td><input type=text name = c27></td><td>&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td><input type=text name = c28></td><td>&nbsp;</td>
</tr>
----->
<tr>
<td></td><td></td><td>&nbsp;</td><td>&nbsp;</td><td><input type=submit value="Дальше >>"></td>
</tr>
</table>
</align>
</form>
<center>
<table border=0 >
<td>
<tr>		
		<td>
		<a href="search.asp"><font  size=3>Найти заявки от клиентов</font></a>
		<hr>
		</td>
		<td>
		<a href="finansi.asp"><font  size=3>Финансы</font></a>
		<hr>
		</td>
		<td>
		<a href="history.asp"><font  size=3>Журнал</font></a>
		<hr>
		</td>
		<td>
		<a href="reg.asp"><font  size=3>Зарегистрироваться</font></a>
		<hr>
		</td>
</tr>
</table>
</center>


</td>
</tr>
		
</table>		

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