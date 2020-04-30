<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/myConnection.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="1,2,3"
MM_authFailedURL="login.asp?msg=2"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (false Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<%
Dim Rset_products
Dim Rset_products_cmd
Dim Rset_products_numRows

Set Rset_products_cmd = Server.CreateObject ("ADODB.Command")
Rset_products_cmd.ActiveConnection = MM_myConnection_STRING
Rset_products_cmd.CommandText = "SELECT * FROM dbo.VIEW_prod_place_user" 
Rset_products_cmd.Prepared = true

Set Rset_products = Rset_products_cmd.Execute
Rset_products_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Rset_products_numRows = Rset_products_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/myproj/templates/adminTmp.dwt.asp" codeOutsideHTMLIsLocked="false" -->

<head>
<link rel="stylesheet" type="text/css" href="../css/adminCSS.css">
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "login.asp"
  ' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_newQS = "?"
    For Each Item In Request.QueryString
      If (Item <> "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(Request.QueryString(Item))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Untitled Document</title>
<!-- InstanceEndEditable -->
<!-- InstanceBeginEditable name="head" -->
<style type="text/css">
<!--
.style5 {color: #CFDCE6; font-size: 12px; }
-->
</style>
<!-- InstanceEndEditable -->
</head>

<body>
<div id="layer0">
  <div id="layer1">Welcome <%= Session("MM_Username") %> / <a href="<%= MM_Logout %>">logout</a></div>
  <div id="layer2"><!-- InstanceBeginEditable name="EditRegion3" -->
  
  <table width="95%" border="1" align="center">
    <tr bgcolor="#003853">
      <td width="8%" height="30"><div align="center" class="style5">item name</div></td>
      <td width="11%" height="30"><div align="center" class="style5">product number</div></td>
      <td width="12%" height="30"><div align="center" class="style5">company name</div></td>
      <td width="9%" height="30"><div align="center" class="style5">buyer's firstname</div></td>
      <td width="9%" height="30"><div align="center" class="style5">buyer's lasttname</div></td>
      <td width="9%" height="30"><div align="center" class="style5">price of unit</div></td>
      <td width="10%" height="30"><div align="center" class="style5">quantity</div></td>
      <td width="8%" height="30"><div align="center" class="style5">date of purchase</div></td>
      <td width="9%" height="30"><div align="center" class="style5">place of use</div></td>
      <td width="15%" height="30" nowrap="nowrap"><div align="center" class="style5">description</div></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT Rset_products.EOF)) %>
      <tr bgcolor="#FFFFFF">
        <td><div align="center"><%=(Rset_products.Fields.Item("item").Value)%></div></td>
        <td><div align="center"><%=(Rset_products.Fields.Item("productNum").Value)%></div></td>
        <td><div align="center"><%=(Rset_products.Fields.Item("companyName").Value)%></div></td>
        <td><div align="center"><%=(Rset_products.Fields.Item("fName").Value)%></div></td>
        <td><div align="center"><%=(Rset_products.Fields.Item("lName").Value)%></div></td>
        <td><div align="center"><%=(Rset_products.Fields.Item("price").Value)%></div></td>
        <td><div align="center"><%=(Rset_products.Fields.Item("quantity").Value)%></div></td>
        <td><div align="center"><%=(Rset_products.Fields.Item("date").Value)%></div></td>
        <td><div align="center"><%=(Rset_products.Fields.Item("place").Value)%></div></td>
        <td><div align="center"><%=(Rset_products.Fields.Item("description").Value)%></div></td>
      </tr>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Rset_products.MoveNext()
Wend
%>
  </table>
  <div id="back"><a style="color: #003853; background-color:#CFDCE6; display:inline; padding-right:30px;" href="productMng.asp">Back to Product Management</a></div>
  <!-- InstanceEndEditable --></div>
  <div id="layer3"><span class="copyright">Made and Designed by Charsooghi &copy;2009</span></div>
</div>
</body>
<!-- InstanceEnd --></html>
<%
Rset_products.Close()
Set Rset_products = Nothing
%>
