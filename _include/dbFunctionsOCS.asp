<%

function GetDatabaseConnection()

	Dim connection
 
	SET connection = Server.CreateObject("ADODB.Connection")

	'connection.Open Application("ConnStr_360") 	
	connection.Open Application("ConnStr_OCS") 	

	set GetDatabaseConnection = connection

end function

function CloseDatabaseConnection(connection)

	connection.close
	set connection = nothing

end function


function CreateRecordSet(rs)

	set rs = server.CreateObject("adodb.recordset")	
	set CreateRecordSet= rs


end function

sub CloseRecordset(rs)
	
	rs.Close
	set rs = nothing

end sub

function CheckDbConnError(dbConn, ProcNumber, ProcName, ProcSP, ProcSQL)
dim SQLErrorCode
dim sErrorMsg
 
   If  (dbConn.errors.count > 0)   then
      isError = 1
      for counter= 0 to dbConn.errors.count
         SQLErrorCode = dbConn.errors(counter).number 
         sErrorMsg = sErrorMsg &  "Error desc: " & db.errors(counter).description 
      next
     ' CALL SurveyErrorLog_Insert( ProcNumber, "DBError", SQLErrorCode, sErrorMsg, ProcName,ProcSP ,ProcSQL)
     ' Response.Redirect "/360/_common/LDS_Stop.asp?errmsg=database error: [getsurveydata]"
     
     CheckDbConnError = True
   else
     CheckDBConnError = False  
   end if

end function

function RSLoopTemplate(dbConn)
 dim rsL
 dim sSQL

  '  set dbConn = GetDatabaseConnection()
  '  CloseRecordSet(rsL)

    sSQL = "EXEC [dcReportDefinitionList_SelectPortal] 'F64F1930-7E4E-4CFF-946D-4DDA01EC0CE0',2,20"


	CreateRecordSet(rsL)
	set rsL = dbConn.execute(sSql)	
	
	if CheckDbConnError(dbConn, "101", "RSLoopTemplate", "dcReportDefinitionList_SelectPortal", sSQL) then
       CloseRecordSet(rsL)
	else
				
    if NOT ((rsL.EOF) AND (rsL.BOF)) Then 
   rsL.MoveFirst

   Do Until rsL.EOF
    For each fld in rsL.Fields
     lsVar = fld.name 
     liValue = fld.value
     
     lsVar1 = lsVar
     if NOT IsNull(liValue) then
       if NOT IsEmpty(liValue)  then
          'Session("Data_"&lsVar1) = liValue 
          Response.write(lsvar&" "&cstr(liValue)&"<br>")
       end if
     end if  
    Next
         
  rsL.MoveNext
  Loop		
  
  else
    'NO DATA
  end if
   CloseRecordSet(rsL)
  end if
  
  
end function

function RSArrayTemplate(dbConn)
 dim rsL
 dim aData
 dim sSQL
 dim jsontree

    sSQL = "EXEC [dcReportDefinitionList_SelectPortal] 'F64F1930-7E4E-4CFF-946D-4DDA01EC0CE0',2,20"


	CreateRecordSet(rsL)
	set rsL = dbConn.execute(sSql)	
   
	if CheckDbConnError(dbConn, "101", "RSLoopTemplate", "dcReportDefinitionList_SelectPortal", sSQL) then
       CloseRecordSet(rsL)
	else
   
    if NOT ((rsL.EOF) AND (rsL.BOF)) Then 
   
     ' jsontree= "["
      RSArrayTemplate = rsL.GetRows
  
    '  For i = 0 to uBound(aData,2)
     '   lsVar = aData(0,i)
     '   lsValue = aData(1,i)
     '   jsontree = jsontree +"{ 'title':'"&aData(4,i)&"', 'key': 'k1"&adata(0,i)&", 'isLazy': true }"
     '   if i <> uBound(aData,2) then
     '     jsontree = jsontree +","
     '   end if

     
        'Session("Text_"&lsVar) = lsValue
      'Next
      'jsontree = jsontree+"]"
      'response.write(jsontree)
   else
    'NO DATA
   end if
   CloseRecordSet(rsL)
   
   end if 
  
end function

function ReportDefinitionList(dbConn)
 dim rsL
 dim aData
 dim sSQL
 dim jsontree


    sSQL = "EXEC [dcReportDefinitionList_SelectPortal] 'F64F1930-7E4E-4CFF-946D-4DDA01EC0CE0',2,20"
	CreateRecordSet(rsL)
	set rsL = dbConn.execute(sSql)	
   
	if CheckDbConnError(dbConn, "101", "RSLoopTemplate", "dcReportDefinitionList_SelectPortal", sSQL) then
       CloseRecordSet(rsL)
       ReportDefinitionList = -1
       Exit Function
	end if
   
    if NOT ((rsL.EOF) AND (rsL.BOF)) Then 
       ReportDefinitionList = rsL.GetRows
    else
      ReportDefinitionList = 0
    end if
    CloseRecordSet(rsL)
   
     
  
end function





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


sub CreateNewTranslationLanguage(language, editor)

	dim sSQL

	sSQL = "exec ml_CreateNewTranslationLanguage " & language & ",'" & editor & "'"

	dim db
	set db = GetDatabaseConnection()
	
	db.execute(sSql)

	call  CloseDatabaseConnection(db)


end sub

 %>