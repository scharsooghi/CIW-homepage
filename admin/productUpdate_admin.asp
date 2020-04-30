<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/myConnection.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="1"
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
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_myConnection_STRING
    MM_editCmd.CommandText = "UPDATE dbo.tbl_Product SET item = ?, productNum = ?, companyName = ?, price = ?, quantity = ?, [date] = ?, placeOfUse_ref = ?, [description] = ? WHERE id = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("item")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("productNum")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 50, Request.Form("companyName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("price"), Request.Form("price"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("quantity"), Request.Form("quantity"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 135, 1, -1, MM_IIF(Request.Form("date"), Request.Form("date"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 5, 1, -1, MM_IIF(Request.Form("placeOfUse_ref"), Request.Form("placeOfUse_ref"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 203, 1, 1073741823, Request.Form("description")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "productEdit_admin.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>
<%
Dim Rset_places
Dim Rset_places_cmd
Dim Rset_places_numRows

Set Rset_places_cmd = Server.CreateObject ("ADODB.Command")
Rset_places_cmd.ActiveConnection = MM_myConnection_STRING
Rset_places_cmd.CommandText = "SELECT * FROM dbo.tbl_place" 
Rset_places_cmd.Prepared = true

Set Rset_places = Rset_places_cmd.Execute
Rset_places_numRows = 0
%>
<%
Dim Rset_product__MMColParam
Rset_product__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  Rset_product__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim Rset_product
Dim Rset_product_cmd
Dim Rset_product_numRows

Set Rset_product_cmd = Server.CreateObject ("ADODB.Command")
Rset_product_cmd.ActiveConnection = MM_myConnection_STRING
Rset_product_cmd.CommandText = "SELECT * FROM dbo.tbl_Product WHERE id = ?" 
Rset_product_cmd.Prepared = true
Rset_product_cmd.Parameters.Append Rset_product_cmd.CreateParameter("param1", 5, 1, -1, Rset_product__MMColParam) ' adDouble

Set Rset_product = Rset_product_cmd.Execute
Rset_product_numRows = 0
%>
<%
Dim Rset_prod_user__MMColParam
Rset_prod_user__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  Rset_prod_user__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim Rset_prod_user
Dim Rset_prod_user_cmd
Dim Rset_prod_user_numRows

Set Rset_prod_user_cmd = Server.CreateObject ("ADODB.Command")
Rset_prod_user_cmd.ActiveConnection = MM_myConnection_STRING
Rset_prod_user_cmd.CommandText = "SELECT * FROM dbo.VIEW_prod_place_user WHERE id = ?" 
Rset_prod_user_cmd.Prepared = true
Rset_prod_user_cmd.Parameters.Append Rset_prod_user_cmd.CreateParameter("param1", 5, 1, -1, Rset_prod_user__MMColParam) ' adDouble

Set Rset_prod_user = Rset_prod_user_cmd.Execute
Rset_prod_user_numRows = 0
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
.style1 {color: #FFFFFF}
-->
</style>
<!-- InstanceEndEditable -->
</head>

<body>
<div id="layer0">
  <div id="layer1">Welcome <%= Session("MM_Username") %> / <a href="<%= MM_Logout %>">logout</a></div>
  <div id="layer2"><!-- InstanceBeginEditable name="EditRegion3" -->
  <form action="<%=MM_editAction%>" method="POST" name="form1" id="form1">
    <table align="center" cellpadding="5" bgcolor="#FFFFFF">
      <tr valign="baseline" bgcolor="#003853">
        <td colspan="2" align="right" nowrap="nowrap"><div align="left" class="style1">Update Form:</div></td>
        </tr>
      <tr valign="baseline">
        <td nowrap="nowrap" align="right"><div align="left">Item name :</div></td>
        <td><input type="text" name="item" value="<%=(Rset_product.Fields.Item("item").Value)%>" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td nowrap="nowrap" align="right"><div align="left">Product number :</div></td>
        <td><input type="text" name="productNum" value="<%=(Rset_product.Fields.Item("productNum").Value)%>" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td nowrap="nowrap" align="right"><div align="left">Company name :</div></td>
        <td><input type="text" name="companyName" value="<%=(Rset_product.Fields.Item("companyName").Value)%>" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td nowrap="nowrap" align="right"><div align="left">Buyer_ref:</div></td>
        <td><%=(Rset_prod_user.Fields.Item("fName").Value)%> <%=(Rset_prod_user.Fields.Item("lName").Value)%></td>
      </tr>
      <tr valign="baseline">
        <td nowrap="nowrap" align="right"><div align="left">Price of unit:</div></td>
        <td><input type="text" name="price" value="<%=(Rset_product.Fields.Item("price").Value)%>" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td nowrap="nowrap" align="right"><div align="left">Quantity:</div></td>
        <td><input type="text" name="quantity" value="<%=(Rset_product.Fields.Item("quantity").Value)%>" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td nowrap="nowrap" align="right"><div align="left">Date:</div></td>
        <td><input type="text" name="date" value="<%=(Rset_product.Fields.Item("date").Value)%>" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td nowrap="nowrap" align="right"><div align="left">Place of use :</div></td>
        <td><select name="placeOfUse_ref">
            <%
While (NOT Rset_places.EOF)
%>
            <option value="<%=(Rset_places.Fields.Item("id").Value)%>" <%If (Not isNull(Rset_product.Fields.Item("placeOfUse_ref").Value)) Then If (CStr(Rset_places.Fields.Item("id").Value) = CStr(Rset_product.Fields.Item("placeOfUse_ref").Value)) Then Response.Write("selected='selected'") : Response.Write("")%> ><%=(Rset_places.Fields.Item("place").Value)%></option>
            <%
  Rset_places.MoveNext()
Wend
If (Rset_places.CursorType > 0) Then
  Rset_places.MoveFirst
Else
  Rset_places.Requery
End If
%>
          </select>        </td>
      </tr>
      <tr>
        <td nowrap="nowrap" align="right" valign="top"><div align="left">Description:</div></td>
        <td valign="baseline"><textarea name="description" cols="50" rows="5"><%=(Rset_product.Fields.Item("description").Value)%></textarea>        </td>
      </tr>
      <tr valign="baseline">
        <td nowrap="nowrap" align="right"><div align="left"></div></td>
        <td><input type="submit" value="Update record" />        </td>
      </tr>
    </table>
    <input type="hidden" name="MM_update" value="form1" />
    <input type="hidden" name="MM_recordId" value="<%= Rset_product.Fields.Item("id").Value %>" />
  </form>
  <div id="back"><a style="color: #003853; background-color:#CFDCE6; display:inline; padding-right:30px;" href="productMng.asp">Back to Product Management</a></div>
  <!-- InstanceEndEditable --></div>
  <div id="layer3"><span class="copyright">Made and Designed by Charsooghi &copy;2009</span></div>
</div>
</body>
<!-- InstanceEnd --></html>
<%
Rset_places.Close()
Set Rset_places = Nothing
%>
<%
Rset_product.Close()
Set Rset_product = Nothing
%>
<%
Rset_prod_user.Close()
Set Rset_prod_user = Nothing
%>
