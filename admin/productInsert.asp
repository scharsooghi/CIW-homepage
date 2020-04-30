<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/myConnection.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="1,2"
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
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_myConnection_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.tbl_Product (item, productNum, companyName, buyer_ref, price, quantity, [date], placeOfUse_ref, [description]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("item")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("productNum")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 50, Request.Form("companyName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("buyer_ref"), Request.Form("buyer_ref"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 5, 1, -1, MM_IIF(Request.Form("price"), Request.Form("price"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 5, 1, -1, MM_IIF(Request.Form("quantity"), Request.Form("quantity"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 135, 1, -1, MM_IIF(Request.Form("date"), Request.Form("date"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 5, 1, -1, MM_IIF(Request.Form("placeOfUse_ref"), Request.Form("placeOfUse_ref"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 203, 1, 1073741823, Request.Form("description")) ' adLongVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "productMng.asp?msg=1"
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
Dim Rset_place
Dim Rset_place_cmd
Dim Rset_place_numRows

Set Rset_place_cmd = Server.CreateObject ("ADODB.Command")
Rset_place_cmd.ActiveConnection = MM_myConnection_STRING
Rset_place_cmd.CommandText = "SELECT * FROM dbo.tbl_place" 
Rset_place_cmd.Prepared = true

Set Rset_place = Rset_place_cmd.Execute
Rset_place_numRows = 0
%>
<%
Dim Rset_buyers__MMColParam
Rset_buyers__MMColParam = "1"
If (Session("MM_Username") <> "") Then 
  Rset_buyers__MMColParam = Session("MM_Username")
End If
%>
<%
Dim Rset_buyers
Dim Rset_buyers_cmd
Dim Rset_buyers_numRows

Set Rset_buyers_cmd = Server.CreateObject ("ADODB.Command")
Rset_buyers_cmd.ActiveConnection = MM_myConnection_STRING
Rset_buyers_cmd.CommandText = "SELECT * FROM dbo.tbl_Users WHERE username = ?" 
Rset_buyers_cmd.Prepared = true
Rset_buyers_cmd.Parameters.Append Rset_buyers_cmd.CreateParameter("param1", 200, 1, 50, Rset_buyers__MMColParam) ' adVarChar

Set Rset_buyers = Rset_buyers_cmd.Execute
Rset_buyers_numRows = 0
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
<title>Product Manager</title>
<!-- InstanceEndEditable -->
<!-- InstanceBeginEditable name="head" -->
<style type="text/css">
<!--
.style1 {
	color: #FFFFFF;
	font-size: 14px;
}
-->
</style>
<script>
function clickClear(obj,text){
	if(obj.value == text)
		obj.value = '';
		}

function clickRecall(obj, text)
{
	if(obj.value == "")
		obj.value = text;
		}
</script>
<!-- InstanceEndEditable -->
</head>

<body>
<div id="layer0">
  <div id="layer1">Welcome <%= Session("MM_Username") %> / <a href="<%= MM_Logout %>">logout</a></div>
  <div id="layer2"><!-- InstanceBeginEditable name="EditRegion3" -->
  <form action="<%=MM_editAction%>" method="post" name="form1" id="form1">
    <table width="300" align="center" cellpadding="5" bgcolor="#FFFFFF">
      <tr valign="baseline">
        <td colspan="2" align="right" valign="middle" nowrap="nowrap" bgcolor="#003853"><div align="left" class="style1">Insert Form:</div></td>
        </tr>
      <tr valign="baseline">
        <td width="114" align="right" valign="middle" nowrap="nowrap" bgcolor="#FFFFFF"><div align="left">Item name :</div></td>
        <td width="200" bgcolor="#FFFFFF"><input name="item" type="text" class="textfields" value="" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap="nowrap" bgcolor="#FFFFFF"><div align="left">Product number :</div></td>
        <td width="200" bgcolor="#FFFFFF"><input name="productNum" type="text" class="textfields" value="" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap="nowrap" bgcolor="#FFFFFF"><div align="left">Company name :</div></td>
        <td width="200" bgcolor="#FFFFFF"><input name="companyName" type="text" class="textfields" value="" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap="nowrap" bgcolor="#FFFFFF"><div align="left">Price of unit :</div></td>
        <td width="200" bgcolor="#FFFFFF"><input name="price" type="text" class="textfields" value="" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap="nowrap" bgcolor="#FFFFFF"><div align="left">Quantity:</div></td>
        <td width="200" bgcolor="#FFFFFF"><input name="quantity" type="text" class="textfields" value="" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap="nowrap" bgcolor="#FFFFFF"><div align="left">Date :</div></td>
        <td width="200" bgcolor="#FFFFFF"><input name="date" type="text" class="textfields" value="yy-mm-dd" size="32" onclick="clickClear(this,'yy-mm-dd')" onblur="clickRecall(this, 'yy-mm-dd')"/>        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap="nowrap" bgcolor="#FFFFFF"><div align="left">Place of use :</div></td>
        <td width="200" bgcolor="#FFFFFF"><select name="placeOfUse_ref">
            <%
While (NOT Rset_place.EOF)
%>
            <option value="<%=(Rset_place.Fields.Item("id").Value)%>"><%=(Rset_place.Fields.Item("place").Value)%></option>
            <%
  Rset_place.MoveNext()
Wend
If (Rset_place.CursorType > 0) Then
  Rset_place.MoveFirst
Else
  Rset_place.Requery
End If
%>
          </select>        </td>
      </tr>
      <tr>
        <td align="right" valign="middle" nowrap="nowrap" bgcolor="#FFFFFF"><div align="left">Description:</div></td>
        <td width="200" valign="baseline" bgcolor="#FFFFFF"><textarea name="description" cols="50" rows="5" class="textfields"></textarea>        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap="nowrap" bgcolor="#FFFFFF"><div align="left"></div></td>
        <td width="200" bgcolor="#FFFFFF"><input type="submit" class="loginBT" value="Insert" />        </td>
      </tr>
    </table>
    <input type="hidden" name="buyer_ref" value="<%=(Rset_buyers.Fields.Item("id").Value)%>" />
    <input type="hidden" name="MM_insert" value="form1" />
  </form>
  <div id="back"><a style="color: #003853; background-color:#CFDCE6; display:inline; padding-right:30px;" href="productMng.asp">Back to Product Management</a></div>
  <!-- InstanceEndEditable --></div>
  <div id="layer3"><span class="copyright">Made and Designed by Charsooghi &copy;2009</span></div>
</div>
</body>
<!-- InstanceEnd --></html>
<%
Rset_place.Close()
Set Rset_place = Nothing
%>
<%
Rset_buyers.Close()
Set Rset_buyers = Nothing
%>
