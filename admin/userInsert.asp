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
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_myConnection_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.tbl_Users (accLevel_ref, username, password, fName, lName, email) VALUES (?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("accLevel_ref"), Request.Form("accLevel_ref"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("username")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 50, Request.Form("password")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 50, Request.Form("fName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 50, Request.Form("lName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 50, Request.Form("email")) ' adVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "userMng.asp?msg=1"
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
Dim Rset_accLevel
Dim Rset_accLevel_cmd
Dim Rset_accLevel_numRows

Set Rset_accLevel_cmd = Server.CreateObject ("ADODB.Command")
Rset_accLevel_cmd.ActiveConnection = MM_myConnection_STRING
Rset_accLevel_cmd.CommandText = "SELECT * FROM dbo.tbl_AccessLevel" 
Rset_accLevel_cmd.Prepared = true

Set Rset_accLevel = Rset_accLevel_cmd.Execute
Rset_accLevel_numRows = 0
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
<title>User Insert</title>
<!-- InstanceEndEditable -->
<!-- InstanceBeginEditable name="head" -->
<style type="text/css">
<!--
.style2 {
	font-size: 14px;
	color: #FFFFFF;
}
-->
</style>
<!-- InstanceEndEditable -->
</head>

<body>
<div id="layer0">
  <div id="layer1">Welcome <%= Session("MM_Username") %> / <a href="<%= MM_Logout %>">logout</a></div>
  <div id="layer2"><!-- InstanceBeginEditable name="EditRegion3" -->
  <form action="<%=MM_editAction%>" method="post" name="form1" id="form1">
    <table align="center" cellpadding="5" bgcolor="#FFFFFF">
      <tr valign="baseline">
        <td colspan="2" align="right" valign="middle" nowrap="nowrap" bgcolor="#003853"><div align="left" class="style2">Insert Form</div></td>
        </tr>
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap="nowrap"><div align="left">Access level :</div></td>
        <td><select name="accLevel_ref">
            <%
While (NOT Rset_accLevel.EOF)
%>
            <option value="<%=(Rset_accLevel.Fields.Item("id").Value)%>"><%=(Rset_accLevel.Fields.Item("accessLevel").Value)%></option>
            <%
  Rset_accLevel.MoveNext()
Wend
If (Rset_accLevel.CursorType > 0) Then
  Rset_accLevel.MoveFirst
Else
  Rset_accLevel.Requery
End If
%>
          </select>        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap="nowrap"><div align="left">Username :</div></td>
        <td><input name="username" type="text" class="textfields" value="" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap="nowrap"><div align="left">Password :</div></td>
        <td><input name="password" type="text" class="textfields" value="" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap="nowrap"><div align="left">First name :</div></td>
        <td><input name="fName" type="text" class="textfields" value="" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap="nowrap"><div align="left">Last name :</div></td>
        <td><input name="lName" type="text" class="textfields" value="" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap="nowrap"><div align="left">Email:</div></td>
        <td><input name="email" type="text" class="textfields" value="" size="32" />        </td>
      </tr>
      <tr valign="baseline">
        <td align="right" valign="middle" nowrap="nowrap"><div align="left"></div></td>
        <td><input type="submit" class="loginBT" value="Insert" />        </td>
      </tr>
    </table>
    <input type="hidden" name="MM_insert" value="form1" />
  </form>
  <div id="back"><a style="color: #003853; background-color:#CFDCE6; display:inline; padding-right:30px;" href="userMng.asp">Back to User Management</a></div>
  <!-- InstanceEndEditable --></div>
  <div id="layer3"><span class="copyright">Made and Designed by Charsooghi &copy;2009</span></div>
</div>
</body>
<!-- InstanceEnd --></html>
<%
Rset_accLevel.Close()
Set Rset_accLevel = Nothing
%>
