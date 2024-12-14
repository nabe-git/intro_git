<!-- #include file="plat.asp" -->
<!-- #include file="common.asp" -->
<%

'テーブル一つだけを利用したテーブル
dim xi
dim where_string

call makesql

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
		<title>NISSIN WEB SYSTEM</title>
		<link rel="stylesheet" href="style/t-plat.css" type="text/css">
		<script language="JavaScript" src="image\calendarlay.js"></script>
		<script language="javascript">
		<!--
		function changeCSS(CSS_NO) {
			if (document.getElementById(CSS_NO).style.color == "red") {
				document.getElementById(CSS_NO).style.color = "";
				document.getElementById(CSS_NO).style.fontWeight = "";
			} else {
				document.getElementById(CSS_NO).style.color = "red";
				document.getElementById(CSS_NO).style.fontWeight = 400;
			}
		}
		//-->
		</script>
	</head>
<body>
	<table width="100%" bgcolor="#5A90D3">
		<tr>
			<td class="top_td" width="30" ><img src="image\nissin_mark.png"></td>
			<td class="top_td" ><font color="#FFFFFF">&nbsp;&nbsp;Infomation Platform - </font><font size="3" color="#FFFFFF"><%= session("REPORTNM") %></font><font size="2" color="5A90D3">-<%= session("REPORTID") %></font></td>
			<td width="600" class="top_td" align="center"><%= fun_condition %></td>
			<td width="100" class="top_td" align="center"><a href="condition.asp"><font color="#FFFFFF">SEARCH</font></a></td>
			<td width="100" class="top_td" align="center"><a href="logout.asp"><font color="#FFFFFF">LOGOUT</font></a></td>
		</tr>
	</table>

	<%= report_list %></p>
	<% if session("error") <> "" then %>
		<font size="3" color="RED"><%= session("error") %></font>
		<% session("error") = "" %>
	<% end if %>

	<div class="form_margin">
	<% if session("INSERT_GRANT") = "Y" Then %>
		<p><a href="ipinst_type00.asp">INSERT</a></p>
	<% end if %>
	</div>

	<% if session("UPDATE_GRANT") = "Y" then %>
		<form name = "f1" method="POST" action="ipchk_type00.asp" style="margin: 0 0 0 5px;">
	<% end if %>

	<% if session("UPDATE_GRANT") = "Y" and ubound(up_any_fldlist) >= 0 then %>
		<p>
		<input type="hidden" name="PFLG" value="AS">
		<% for xi = 0 to ubound(up_any_fldlist ) %>
			<%= replace(alias.item(up_any_fldlist(xi)),"<BR>"," ") %>:
			<% if attribute.item(up_any_fldlist(xi)) = "DATE" Then %>
				<input type=text size="<%= length.item(up_any_fldlist(xi)) + 4 %>" maxlength="<%= length.item(up_any_fldlist(xi)) %>" name="<%= up_any_fldlist(xi) %>" onFocus="wrtCalendarLay(this,event)">
			<% Else %>
				<input type=text size="<%= length.item(up_any_fldlist(xi)) + 4 %>" maxlength="<%= length.item(up_any_fldlist(xi)) %>" name="<%= up_any_fldlist(xi) %>">
			<% End If %>
		<% next %>
		<input type="submit" value="UPDATE">
		</p>
	<% end if %>

	<% If not db.eof Then %>
		<table>
			<tr>
			<% 
			dim w_title

			w_title = ""
			if session("UPDATE_GRANT") = "Y" then 
				if w_title = "" then
					w_title = w_title & "U" 
				else
					w_title = w_title & ",U" 
				end if
			end if

			if session("DELETE_GRANT") = "Y" then 
				if w_title = "" then
					w_title = w_title & "D" 
				else
					w_title = w_title & ",D" 
				end if
			end if

			if session("COPY_GRANT") = "Y" then 
				if w_title = "" then
					w_title = w_title & "C" 
				else
					w_title = w_title & ",C" 
				end if
			end if
			%>

			<% if w_title <> "" then %>
				<th nowrap><%= w_title %></th>
			<% end if %>

			<% For Xi = 1 To db.recordset.fields.Count - 1  %><th nowrap><%= alias.item(db.name(Xi)) %></th><% Next %>
				</tr>
				<% line_no = 0 %>
				<% Do Until db.Eof %>
					<% line_no = line_no + 1 %>
					<% If Db.Recordset.RowPosition >= From_Rset And Db.Recordset.RowPosition <= To_Rset Then %>
						<tr bgcolor="<%= setbgcolor(db.data(0)) %>" id="CSS<%= line_no %>" onClick="changeCSS('CSS<%= line_no %>')">
						<%
						w_title = ""

						if session("UPDATE_GRANT") = "Y" then 
							if w_title = "" then
								w_title = w_title & "<a href=""ipedit_type00.asp?SEQ=" & db.data(0) & """>U</a>" 
							else
								w_title = w_title & ",<a href=""ipedit_type00.asp?SEQ=" & db.data(0) & """>U</a>" 
							end if
						end if

						if session("DELETE_GRANT") = "Y" then 
							if w_title = "" then
								w_title = w_title & "<a href=""ipdele_type00.asp?SEQ=" & db.data(0) & """>D</a>" 
							else
								w_title = w_title & ",<a href=""ipdele_type00.asp?SEQ=" & db.data(0) & """>D</a>" 
							end if
						end if

						if session("COPY_GRANT") = "Y" then 
							if w_title = "" then
								w_title = w_title & "<a href=""ipcopy_type00.asp?SEQ=" & db.data(0) & """>C</a>" 
							else
								w_title = w_title & ",<a href=""ipcopy_type00.asp?SEQ=" & db.data(0) & """>C</a>" 
							end if
						end if

						if session("UPDATE_GRANT") = "Y" And ubound(up_any_fldlist) >= 0 then 
							w_title = w_title & ",<input type=""checkbox"" name=""" & db.name(0)  & """ value=""" & db.data(0)  & """>"
						end if
						%>

						<% if w_title <> "" then %>
							<td nowrap><%= w_title %></td>
						<% end if %>

						<% For Xi = 1 To db.recordset.Fields.Count - 1 %>
							<td nowrap><div align="<%= setdiv(attribute(db.name(Xi))) %>"><%= db.data(xi) %></div></td>
						<% next %>
					<% End If %>
				<% if db.nextdata = false then %>
				<% end if %>
				</tr>
			<% loop %>
		</table>
		<%
		dim cCond 
		dim cCondItem
		set ccond = session("colcond_app")
		For xi = 1 to session("intcond") 
			response.Write("<a href=""chgcondition.asp?int=" & xi & """>" & ccond.item(cstr(xi)) & "</a>&nbsp;" & "&nbsp;" & "&nbsp;")
		Next
		%>
	<% End If %>

	<% if session("UPDATE_GRANT") = "Y" then %>
		</form>
	<% end if %>
	<br>
	<% if total_page > 1 Then %>
		<% If Session("Current_page") = 1 Then %>
			<a href="<%= session("link_report") %>?PG=N"><img border=0 src="image\dn.gif"></a>&nbsp;&nbsp;PAGE:&nbsp;<%= Session("Current_page") %>/<%= total_page %>
		<% ElseIf Session("Current_page") = Total_Page Then %>
			<a href="<%= session("link_report") %>?PG=P"><img border=0 src="image\up.gif"></a>&nbsp;&nbsp;PAGE:&nbsp;<%= Session("Current_page") %>/<%= total_page %>
		<% Else %>
			<a href="<%= session("link_report") %>?PG=P"><img border=0 src="image\up.gif"></a>&nbsp;&nbsp;<a href="<%= session("link_report") %>?PG=N"><img border=0 src="image\dn.gif"></a>&nbsp;&nbsp;PAGE:&nbsp;<%= Session("Current_page") %>/<%= total_page %>
		<% End If %>
	<% end if %>
	<hr width="100%" size="1">
	<p align="right"><font size="2">ALL Rights Reserved Copyright (C) 2008. NISSIN Corporation</font></p>
</body>
</html>
<% 
set db = nothing 
%>
