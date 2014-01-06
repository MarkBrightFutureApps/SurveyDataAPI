<%

Function isValidEmail(myEmail)
  dim isValidE
  dim regEx 
  isValidE = True
  set regEx = New RegExp 
  regEx.IgnoreCase = False 
  regEx.Pattern = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
  isValidE = regEx.Test(myEmail)
  isValidEmail = isValidE
End Function


 function isValidPassword(myString)
  myString = myString&""
  if (Len(myString)<6) then isValidPassword = false : exit function
  dim isValidE
  dim regEx
  isValidE = True
  set regEx = New RegExp
  regEx.IgnoreCase = False
  regEx.Pattern = "[0-9]"
  isValidE = regEx.Test(myString)
  isValidPassword = isValidE
End Function
%>
