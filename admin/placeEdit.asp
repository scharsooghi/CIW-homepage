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
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_myConnection_STRING
    MM_editCmd.CommandText = "DELETE FROM dbo.tbl_place WHERE id = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "placeEdit.asp"
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
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Rset_places_numRows = Rset_places_numRows + Repeat1__numRows
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
<title>Place Edit</title>
<!-- InstanceEndEditable -->
<!-- InstanceBeginEditable name="head" -->
<style type="text/css">
<!--
.style1 {color: #CFDCE6}
.style2 {font-size: 14px}
.style3 {color: #CFDCE6; font-size: 14px; }
-->
</style><!-- InstanceEndEditable -->
</head>

<body>
<div id="layer0">
  <div id="layer1">Welcome <%= Session("MM_Username") %> / <a href="<%= MM_Logout %>">logout</a></div>
  <div id="layer2"><!-- InstanceBeginEditable name="EditRegion3" -->

  <div align="center">
    <table border="1" align="center" cellpadding="5" bgcolor="#FFFFFF">
        <tr bgcolor="#003853">
          <td height="30"><div align="center" class="style1 style2">place</div></td>
          <td height="30"><div align="center" class="style3">delete</div></td>
        </tr>
        <% While ((Repeat1__numRows <> 0) AND (NOT Rset_places.EOF)) %>
          <tr align="center" valign="middle">
            <td><div align="center"><%=(Rset_places.Fields.Item("place").Value)%></div></td>
            <td>          <form id="form1" name="form1" method="POST" action="<%=MM_editAction%>">
              <label>

                  <input name="button" type="submit" class="loginBT" id="button" value="delete" />
              </label>
                
                  <input type="hidden" name="MM_delete" value="form1" />
                  <input type="hidden" name="MM_recordId" value="<%= Rset_places.Fields.Item("id").Value %>" />
                  
            </form></td>
          </tr>
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Rset_places.MoveNext()
Wend
%>
      </table>
  </div>
  <div id="back"><a style="color: #003853; background-color:#CFDCE6; display:inline; padding-right:30px;" href="placeMng.asp">Back to Place Management</a></div>
  <!-- InstanceEndEditable --></div>
  <div id="layer3"><span class="copyright">Made and Designed by Charsooghi &copy;2009</span></div>
</div>
</body>
<!-- InstanceEnd --></html>
<%
Rset_places.Close()
Set Rset_places = Nothing
%>
