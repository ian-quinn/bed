<%
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("data.mdb")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath



Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from abc "  
rs.open sql,conn,1,3
        rs.addnew
		rs("name")=Request.Form("name")
		rs("com")=Request.Form("com")
		rs("mail")=Request.Form("mail")
		rs.update
		response.Write "<script language='JavaScript'>{window.alert( '欢迎参会！我们将通过邮件与您联系！');window.location.href= 'index.html';}</script> "
        response.end
rs.close
set rs=nothing
%>