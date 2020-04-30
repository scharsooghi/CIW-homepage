<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
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
<title>Place Management</title>
<!-- InstanceEndEditable -->
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
</head>

<body>
<div id="layer0">
  <div id="layer1">Welcome <%= Session("MM_Username") %> / <a href="<%= MM_Logout %>">logout</a></div>
  <div id="layer2"><!-- InstanceBeginEditable name="EditRegion3" -->
  <div align="center">
          <% msg=request.QueryString("msg") %>
      <% If msg=1 Then %>
          <br/><span style="color:#FF0000;">insert is done !! </span><br />   
          <% End If %>
    </div>
  <% If Session("MM_UserAuthorization")=1 or Session("MM_UserAuthorization")=2 Then %>
  <a href="placeInsert.asp">Add new place</a>
  <% End If %>
  <% If Session("MM_UserAuthorization")=1 Then %>
  <a href="placeEdit.asp">delete places</a>
  <% End If %>
  
  <div id="back"><a style="color: #003853; background-color:#CFDCE6; display:inline; padding-right:30px;" href="CPanel.asp">back to CPanel</a></div><!-- InstanceEndEditable --></div>
  <div id="layer3"><span class="copyright">Made and Designed by Charsooghi &copy;2009</span></div>
</div>
</body>
<!-- InstanceEnd --></html>
