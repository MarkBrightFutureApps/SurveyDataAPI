<script LANGUAGE="VBScript" RUNAT="Server">
Dim rs, dbConn, strConn

Function OpenDB()

  Set dbConn = Server.CreateObject("ADODB.Connection")
   strConn = "Provider=SQLOLEDB.1; Data Source=Denison8; Initial Catalog=SurveyDataTest;User ID=Docs_SoftwareUser; Password=6u3jes2goine; TRUSTED_CONNECTION=NO"
 

  dbConn.Open strConn
End Function

Function CloseDB()
	Set rs = Nothing
	If ucase(TypeName(dbConn)) = "OBJECT" Then
		dbConn.Close
		Set dbConn = Nothing
	End If
End Function


</script>