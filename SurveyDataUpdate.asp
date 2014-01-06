<% Response.Buffer=TRUE %>
<!-- #include file ='_include/dbconnOCS.asp' -->



<%

Dim strOper, strID, strUserID, strGender 
Dim fUserID, fGender

strUserID = Request("UserID") 

strGender = request("Gender")

fUserID = request.form("UserID")
fGender = request.form("Gender")

if fUserID <> "" then

  Response.write("fUserID = "&fUserID)
End if

'Response.write "Gender"

strAccountName =  Replace(Request("cAccountName"), "'", "''") 
if strAccountName = "" then
 strAccountName =  Replace(Request("AccountName"), "'", "''") 
end if
'straAccountID = Replace(Request("aAccountID"), "'", "''") 

strOper = "Add" 
strSQL = ""
if strUserID <> "" then
Select   Case strOper 
    Case  "Add": 'Add Record
      strSQL = "Insert dbo.[SurveyData] (UserID, Gender ) Values (" & strUserID & ", "& strGender  &")"
    Case "edit": 'Edit
      strSQL = "Update dbo.[SurveyData] Set Gender = "&strGender
      strSQL = strSQL & " Where UserID = "&strUserID 
    Case "Del": 'Delete
      strSQL =  "Delete From dbo.[SurveyData] Where UserID = " & strUserID
End Select       
end if   
  
  Call OpenDB()
  Response.write "&UserID=998"
  if strUserID <> "" then
   Set rs = dbConn.Execute(strSQL) 
  else
     strSQL = "Insert dbo.[SurveyData] (UserID, Gender ) Values (888,3)"
	 Set rs = dbConn.Execute(strSQL) 
  end if 
 Call CloseDB()
%>