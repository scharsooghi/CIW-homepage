<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/myConnection.asp" -->
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
' *** Redirect if username exists
MM_flag = "MM_insert"
If (CStr(Request(MM_flag)) <> "") Then
  Dim MM_rsKey
  Dim MM_rsKey_cmd
  
  MM_dupKeyRedirect = "registration.asp?msg=1"
  MM_dupKeyUsernameValue = CStr(Request.Form("username"))
  Set MM_rsKey_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsKey_cmd.ActiveConnection = MM_myConnection_STRING
  MM_rsKey_cmd.CommandText = "SELECT username FROM dbo.tbl_Users WHERE username = ?"
  MM_rsKey_cmd.Prepared = true
  MM_rsKey_cmd.Parameters.Append MM_rsKey_cmd.CreateParameter("param1", 200, 1, 50, MM_dupKeyUsernameValue) ' adVarChar
  Set MM_rsKey = MM_rsKey_cmd.Execute
  If Not MM_rsKey.EOF Or Not MM_rsKey.BOF Then 
    ' the username was found - can not add the requested username
    MM_qsChar = "?"
    If (InStr(1, MM_dupKeyRedirect, "?") >= 1) Then MM_qsChar = "&"
    MM_dupKeyRedirect = MM_dupKeyRedirect & MM_qsChar & "requsername=" & MM_dupKeyUsernameValue
    Response.Redirect(MM_dupKeyRedirect)
  End If
  MM_rsKey.Close
End If
%>
<%
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_myConnection_STRING
    MM_editCmd.CommandText = "INSERT INTO dbo.tbl_Users (username, password, fName, lName, email) VALUES (?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("username")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("password")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 50, Request.Form("fName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 50, Request.Form("lName")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 50, Request.Form("email")) ' adVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "login.asp"
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link rel="stylesheet" type="text/css" href="../css/adminCSS.css">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
<style type="text/css">
<!--
.style1 {color: #FFFFFF}
-->
</style>
</head>

<body>
<div align="center">
          <% msg=request.QueryString("msg") %>
<% If msg=1 Then %>
          <br/><span style="color:#FF0000;">This user name already exists !! </span>
 
          <% End If %>
</div>

<form action="<%=MM_editAction%>" method="POST" name="form1" id="form1">
  <table width="350" align="center" cellpadding="5" bgcolor="#C8D7E1">
    <tr valign="middle" bgcolor="#003853">
      <td height="30" colspan="2" align="right" nowrap="nowrap"><div align="left" class="style1">Registration Form :</div></td>
    </tr>
    <tr valign="middle">
      <td align="right" nowrap="nowrap"><div align="left">Username :</div></td>
      <td><div align="center">
        <input name="username" type="text" class="textfields" value="" size="32" />      
      </div></td>
    </tr>
    <tr valign="middle">
      <td align="right" nowrap="nowrap"><div align="left">Password :</div></td>
      <td><div align="center">
        <input name="password" type="password" class="textfields" value="" size="32" />      
      </div></td>
    </tr>
    <tr valign="middle">
      <td align="right" nowrap="nowrap"><div align="left">First name :</div></td>
      <td><div align="center">
        <input name="fName" type="text" class="textfields" value="" size="32" />      
      </div></td>
    </tr>
    <tr valign="middle">
      <td align="right" nowrap="nowrap"><div align="left">Last name :</div></td>
      <td><div align="center">
        <input name="lName" type="text" class="textfields" value="" size="32" />      
      </div></td>
    </tr>
    <tr valign="middle">
      <td align="right" nowrap="nowrap"><div align="left">Email :</div></td>
      <td><div align="center">
        <input name="email" type="text" class="textfields" value="" size="32" />      
      </div></td>
    </tr>
    <tr valign="middle">
      <td colspan="2" align="right" nowrap="nowrap"><div align="left">
          <blockquote>
            <p align="center">
              <input type="submit" class="loginBT" value="register" />      
                  </p>
          </blockquote>
        </div></td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1" />
</form>
<p>&nbsp;</p>
</body>
</html>
