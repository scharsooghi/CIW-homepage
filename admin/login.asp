<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/myConnection.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("usennameTf"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = "accLevel_ref"
  MM_redirectLoginSuccess = "CPanel.asp"
  MM_redirectLoginFailed = "login.asp?msg=1"

  MM_loginSQL = "SELECT username, password"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM dbo.tbl_Users WHERE username = ? AND password = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_myConnection_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 50, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 50, Request.Form("passTf")) ' adVarChar
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link rel="stylesheet" type="text/css" href="../css/adminCSS.css">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Log in</title>
<style type="text/css">
<!--
.style2 {
	color: #93B1C6;
	font-size: 14px;
}
-->
</style>
</head>

<body>
<form id="form1" name="form1" method="POST" action="<%=MM_LoginAction%>">
  <table width="400" align="center" cellpadding="10" cellspacing="0">
    <tr>
      <td colspan="2" valign="middle" bgcolor="#003853"><span class="style2">Login Form</span></td>
    </tr>
    <tr>
      <td width="161" bgcolor="#FFFFFF"><div align="center">username :</div></td>
      <td width="297" bgcolor="#FFFFFF"><label>
        <input name="usennameTf" type="text" class="textfields" id="usennameTf" />
      </label></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><div align="center">password :</div></td>
      <td bgcolor="#FFFFFF"><label>
        <input name="passTf" type="password" class="textfields" id="passTf" />
      </label></td>
    </tr>
    <tr>
      <td colspan="2" valign="middle" bgcolor="#FFFFFF"><label>
        <div align="center">
          <input name="button" type="submit" class="loginBT" id="button" value="Submit" />
        or <a href="registration.asp">sign up!</a></div>
      </label></td>
    </tr>
    <tr>
      <td colspan="2" bgcolor="#FFFFFF">
        <div align="center">
          <% msg=request.QueryString("msg") %>
          <% If msg=1 Then %>
          <br/><span style="color:#FF0000;">Please check the username and password !! </span>
          <% ElseIf msg=2 Then %>
          <br/><span style="color:#FF0000;">The site admin haven't confirmed your registration yet. Please come back later. </span>    
          <% End If %>
        </div></td>
    </tr>
  </table>
</form>
</body>
</html>
