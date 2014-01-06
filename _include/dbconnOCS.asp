<script LANGUAGE="VBScript" RUNAT="Server">
Dim rs, dbConn, strConn

Function OpenDB()

  Set dbConn = Server.CreateObject("ADODB.Connection")
'   strConn = "Provider=SQLOLEDB.1; Data Source=Denison8; Initial Catalog=SurveyDataTest;User ID=Docs_SoftwareUser; Password=6u3jes2goine; TRUSTED_CONNECTION=NO"
 strConn = "Server=3c73e358-a699-4c5f-9d5b-a2ab01813fb5.sqlserver.sequelizer.com;Database=db3c73e358a6994c5f9d5ba2ab01813fb5;User ID=ygyhpgtwrkdrdqle;Password=em8DGgbh48eiP7zkcUiPofSay44iCD4VGjyVwu62oYwKjMnR75NvDU68A8eMk2Lx; "

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