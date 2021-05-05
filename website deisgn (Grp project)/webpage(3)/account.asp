<%@ Language="JavaScript" %>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title></title>
</head>
<body>
 <%
 var CName = Request.Form("CName");
 var CProduct = Request.Form("CProduct");
 var CContact = Request.Form("CContact");
 var conn = new ActiveXObject("ADODB.Connection");
 conn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source='P:/Documents/Report/account.mdb'");
 var SqlString = "INSERT INTO Orders(ClientName, ProductSelect, Contact) VALUES('"+CName+"', '"+CProduct+"', '"+CContact+"')";
 conn.Execute(SqlString);
 conn.Close;
 %>
</body>
</html>
