<%

function GetDatabaseConnection()

	Dim connection
 
	SET connection = Server.CreateObject("ADODB.Connection")

	connection.Open Application("ConnStr_360") 	

	set GetDatabaseConnection = connection

end function

function CloseDatabaseConnection(connection)

	connection.close
	set connection = nothing

end function

sub CloseRecordset(rs)
	
	rs.Close
	set rs = nothing

end sub

function raGetTranslationsForPage(kiLanguage, sPageId)
 dim db
 dim sSQL
 dim rsL


    sSQL = "EXEC [ml_GetAllTranslationsForPage] "&cStr(kiLanguage)&",'" &sPageId& "'"
   'sSQL = "ml_GetAllTranslationsForPage_ByVersion "&nkiLanguage&",'" & sPageId & "'," & nPreviewVersion  & ""

	set rsL = server.CreateObject("adodb.recordset")	

    set db = GetDatabaseConnection()
	set rsL = db.execute(sSql)	
   
    
    raGetTranslationsForPage = rsL.getRows

    rsL.Close
	set rsL = nothing
    Call CloseDatabaseConnection(db)
    ' This doesn't work- >CloseRecordset rs1
   
   ' raGetTranslationsForPage = aData
    'rsGetTranslationsForPage = rs
 
end function

function UpdateTranslationUser( cUserName, cPassword, dExpiration, kiLanguage, kiTranslationUserId )

	dim sSQL

	sSQL = sSQL & "EXEC ml_UpdateTranslationUser '" & cUserName & "', "
	sSQL = sSQL & "'" & cPassword & "', "
	sSQL = sSQL & "'" & dExpiration & "', "
	sSQL = sSQL & kiLanguage & ", "
	sSQL = sSQL & kiTranslationUserId

	dim db
	set db = GetDatabaseConnection()
	
	db.execute(sSql)

	call  CloseDatabaseConnection(db)
		
end function


function GetTranslationUserByName(db, cUserName)

	dim sSQL
	dim rs
	
	sSQL = "EXEC ml_GetTranslationUserByName '" & cUserName & "'"
	set rs = db.execute(sSql)	

	set GetTranslationUserByName = rs

end function


function InsertTranslationUser(cUserName, cPassword, dExpiration, kiLanguage )

	dim sSQL

	sSQL = sSQL & "EXEC ml_InsertTranslationUser '" & cUserName & "', "
	sSQL = sSQL & "'" & cPassword & "', "
	sSQL = sSQL & "'" & dExpiration & "', "
	sSQL = sSQL & kiLanguage
	
	dim db
	set db = GetDatabaseConnection()
	
	db.execute(sSql)

	call  CloseDatabaseConnection(db)
		
end function

function DeleteTranslationUser(kiTranslationUserId)

	dim sSQL
	
	sSQL = "EXEC ml_DeleteTranslationUser " & kiTranslationUserId

	dim db
	set db = GetDatabaseConnection()
	
	db.execute(sSql)

	call  CloseDatabaseConnection(db)

end function

function UpdateThreeClickLabel(cLabelId, cPageId, kiThreeClickLabel)

	dim sSQL
	sSQL = "EXEC ml_UpdateThreeClickLabel '" & cLabelId & "', "
	sSQL = sSQL & "'" & cPageId & "', "
	sSQL = sSQL & kiThreeClickLabel

	dim db
	set db = GetDatabaseConnection()
	
	db.execute(sSql)

	call CloseDatabaseConnection(db)

end function

function GetThreeClickLabel( db, cLabelId )
	dim sSQL
	dim rs
	
	sSQL = "EXEC ml_GetThreeClickLabel '" & cLabelId & "'"
	set rs = db.execute(sSql)	

	set GetThreeClickLabel = rs

end function

function GetThreeClickLabels( db, cLabelId )
	
	dim sSQL
	dim rs
	
	sSQL = "EXEC ml_GetThreeClickLabels '" & cLabelId & "'"
	set rs = db.execute(sSql)	

	set GetThreeClickLabel = rs

end function

function InsertThreeClickLabel( cPageId, cLabelId )

	dim sSQL

	sSQL = "EXEC ml_InsertThreeClickLabel '" & cLabelId & "', "
	sSQL = sSQL & "'" & cPageId & "'"

	dim db
	set db = GetDatabaseConnection()
	
	db.execute(sSql)

	call CloseDatabaseConnection(db)

end function


function InsertThreeClickLabelLang(kiThreeClickLabel, kiLanguage, cLabel)

	dim sSQL
	
	sSQL = "EXEC ml_InsertThreeClickLabelLang " & kiThreeClickLabel & ", "
	sSQL = sSQL & kiLanguage & ", "
	sSQL = sSQL & "'" & cLabel & "'"

	
	dim db
	set db = GetDatabaseConnection()

	db.execute(sSql)

	call CloseDatabaseConnection(db)

end function

function InsertThreeClickLabelLang2(db, kiThreeClickLabel, kiLanguage, cLabel)

	dim sSQL
	
	sSQL = "EXEC ml_InsertThreeClickLabelLang " & kiThreeClickLabel & ", "
	sSQL = sSQL & kiLanguage & ", "
	sSQL = sSQL & "'" & cLabel & "'"

	
	dim rs
	set rs = server.CreateObject("adodb.recordset")

	set rs = db.execute(sSql)

	set InsertThreeClickLabelLang2 = rs

end function


function UpdateThreeClickLabelVersion( kiLanguage, iVersion, tNotes)

	dim sSQL

	sSQL = "EXEC ml_UpdateThreeClickLabelVersion " & kiLanguage & ", "
	sSQL = sSQL & iVersion & ", "
	sSQL = sSQL & "'" & tNotes & "'"

	dim db
	set db = GetDatabaseConnection()
	
	db.execute(sSql)

	call  CloseDatabaseConnection(db)

end function

function UpdateTranslation(kiLanguage, iVersion, sCurKiThreeClickLabel, sCurCLabel, sCurTNotes, TranslationUserName)

	dim sSQL

	sSQL = "EXEC ml_UpdateTranslation "
	sSQL = sSQL & "" & kiLanguage
	sSQL = sSQL & "," & iVersion
	sSQL = sSQL & "," & sCurKiThreeClickLabel
	sSQL = sSQL & ",'" & sCurCLabel   &"'"
	sSQL = sSQL & ",'" & sCurTNotes   &"'"
	sSQL = sSQL & ",'" & session(TranslationUserName)   &"'"

	dim db
	set db = GetDatabaseConnection()
	
	db.execute(sSql)

	call  CloseDatabaseConnection(db)

end function

function PublishTranslationVersion( kiLanguage, iVersion)

	dim sSQL

	sSQL = "exec ml_PublishTranslationVersion " & kiLanguage & "," & iVersion

	dim db
	set db = GetDatabaseConnection()
	
	db.execute(sSql)

	call  CloseDatabaseConnection(db)

end function

function DeleteTranslationVersion( kiLanguage, iVersion)

	dim sSQL

	sSQL = "exec ml_DeleteVersion " & kiLanguage & "," & iVersion

	dim db
	set db = GetDatabaseConnection()
	
	db.execute(sSql)

	call  CloseDatabaseConnection(db)

end function

function CopyLabelFromProdForEditing( kiLabel, kiLanguage, iVersion)

	dim sSQL

	sSQL = "exec ml_CopyLabelFromProdForEditing " & kiLabel & "," & kiLanguage & "," & iVersion

	dim db
	set db = GetDatabaseConnection()
	
	db.execute(sSql)

	call  CloseDatabaseConnection(db)

end function


function CreateNewTranslationVersion( kiLanguage )

	dim sSQL

	sSQL = "exec ml_CreateNewTranslationVersion " & kiLanguage

	dim db
	set db = GetDatabaseConnection()
	
	db.execute(sSql)

	call  CloseDatabaseConnection(db)


end function

function GetUserByCredentials(db, cUsername, cPassword) 

	dim sSQL
	dim rs	
	set rs = server.CreateObject("adodb.recordset")

	sSQL = "exec ml_CheckLogin '" & cUsername& "','" & cPassword & "'"

	set rs = db.execute(sSql)	
				
	set GetUserByCredentials = rs
	 
end function


function GetUserList(db)

	dim sSQL
	dim rs
	
	sSQL = "exec ml_GetUserList"
	set rs = db.execute(sSql)	

	set GetUserList = rs

end function

function GetPageList(db)

	dim sSQL
	dim rs
	
	sSQL = "exec ml_GetPageList"
	set rs = db.execute(sSql)	

	set GetPageList = rs


end function


function GetLabels(db, sEditLabelPageIDFilter)

	dim sSQL
	dim rs
	
	sSQL = "exec ml_GetThreeClickLabels '"  & sEditLabelPageIDFilter & "'"
	set rs = db.execute(sSql)	

	set GetLabels = rs

end function

function GetThreeClickLanguages(db)

	dim sSQL
	dim rs
	
	sSQL = "exec ml_GetThreeClickLanguages"
	set rs = db.execute(sSql)	

	set GetThreeClickLanguages = rs


end function

function GetThreeClickAvailableLanguages(db)

	dim sSQL
	dim rs
	
	sSQL = "exec ml_GetThreeClickAvailableLanguages"
	set rs = db.execute(sSql)	

	set GetThreeClickAvailableLanguages = rs

end function


function GetThreeClickVersionLanguages(db)

	dim sSQL
	dim rs
	
	sSQL = "exec ml_GetThreeClickVersionLanguages"
	set rs = db.execute(sSql)	

	set GetThreeClickVersionLanguages = rs

end function

function GetVersionTable(db, sEditVersionStatusFilter, sLanguageFilter)

	dim sSQL
	dim rs

	set rs = server.CreateObject("adodb.recordset")

	dim filterString
	
	sSQL = "exec ml_GetVersionTable"
	set rs = db.execute(sSql)	

	if sEditVersionStatusFilter <> "All" or sLanguageFilter <> "All Languages" then
		if sEditVersionStatusFilter <> "All" then
			filterString = "  VersionStatus='" & sEditVersionStatusFilter & "'"

			if sLanguageFilter <> "All Languages" then filterString = filterString & " and "
		end if

		if sLanguageFilter <> "All Languages" then
			filterString = filterString & " kiLanguage=" & sLanguageFilter
		end if	

		rs.filter = filterString

	end if
	 
'	rs.sort = "cLanguageNameEng, iVersion desc"

	set GetVersionTable = rs

end function

function GetThreeClickLabelVersionNotes(db, kiLanguage, iVersion)

	dim sSQL
	dim rs
	
	sSQL = "exec ml_GetThreeClickLabelVersionNotes " &  kiLanguage & "," & iVersion
	set rs = db.execute(sSql)	

	set GetThreeClickLabelVersionNotes = rs

end function

function GetAllTranslationsForEditor(db, kiLanguage, iVersion, filterString)

	dim sSQL
	dim rs
	
	sSQL = "exec ml_GetAllTranslationsForEditor " &  kiLanguage & "," & iVersion & ",'" & filterString & "'"
	set rs = db.execute(sSql)	

	set GetAllTranslationsForEditor = rs


end function

sub CreateNewTranslationLanguage(language, editor)

	dim sSQL

	sSQL = "exec ml_CreateNewTranslationLanguage " & language & ",'" & editor & "'"

	dim db
	set db = GetDatabaseConnection()
	
	db.execute(sSql)

	call  CloseDatabaseConnection(db)


end sub

 %>