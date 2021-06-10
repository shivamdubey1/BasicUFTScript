
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''Function ID - START

	''Purpose of the Function -	 If the file does not exist it will create the file and write the status in the file. If the file exist then it will update the status in the file. The File is located at the root of the QTPResults folder.
	''Input Parameters = Status,PropertyFileLocation
	''Return Value - None
	''Sample Function Call - WtiteExecutionTextFile "Passed","C:\QTPResults"
	''Created/Updated by: : Sunil
	
Public Function WriteExecutionTextFile(Status,PropertyFileLocation)
	Dim ExecutionStatus, myFSO, WriteStatus 'Variable Decleration
	WriteStatus = Status
	Set myFSO = CreateObject("Scripting.FileSystemObject")
	If (myFSO.FileExists(PropertyFileLocation&"\ExecutionStatus.txt")) Then
		myFSO.DeleteFile(PropertyFileLocation&"\ExecutionStatus.txt")
		Set ExecutionStatus = myFSO.CreateTextFile(PropertyFileLocation&"\ExecutionStatus.txt", True)
		ExecutionStatus.WriteLine(WriteStatus)
		ExecutionStatus.Close
	Else
		Set ExecutionStatus = myFSO.CreateTextFile(PropertyFileLocation&"\ExecutionStatus.txt", True)
		ExecutionStatus.WriteLine(WriteStatus)
		ExecutionStatus.Close
		
	End If
	Set myFSO = nothing
	Set ExecutionStatus =  nothing
End Function

'''Function ID - END
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Function GetRecordSet(DataSourcepath,SQL,RS)
	Dim Conn
	Set Conn = CreateObject("ADODB.Connection") ' create connection object
	'Create Connection 
		With Conn ' build connection string and open connection
			.Provider = "Microsoft.Jet.OLEDB.4.0"
			.ConnectionString = "Data Source=" & DataSourcepath & ";" & "Extended Properties=Excel 8.0;"
			.Open
		End With
	Set RS = CreateObject("ADODB.Recordset") ' create recordset object
	RS.open SQL, Conn ' open recordset using the connection object 
	Set Conn = Nothing
End Function



Function TIAOracleListSelect(TIAOracleList,SelectList)
   If Trim(SelectList) <> "" Then
	   TIAOracleList.Select Cstr(SelectList)
   End If
End Function

 Function TIAOracleTextFieldEnter(TIAOracleTextField,TextFieldValue)
   Dim sError
   If Trim(TextFieldValue) <> "" Then
	   'Calling Function To Check Object Enabled
		CheckObjEnabled TIAOracleTextField,sError
		If IsEmpty(sError) = True  Then
			'TIAOracleTextField.Click
			TIAOracleTextField.Enter TextFieldValue
		Else
			Reporter.ReportEvent micFail,"TIAOracleTextFieldEnter","Object Disabled. Cannot Enter Value"
		End If
   End If
End Function

Function TIAOracleCheckBoxSelect(TIAOracleCheckBox)
   Dim sError
   'Calling Function To Check Object Enabled
	CheckObjEnabled TIAOracleCheckBox,sError
	If IsEmpty(sError) = True  Then
		TIAOracleCheckBox.Select
	Else
		Reporter.ReportEvent micFail,"TIAOracleCheckBox","Object Disabled. Cannot Click."
	End If
End Function

Function TIAOracleRadioGroupSelect(TIAOracleRadioGroup,SelectValue)
   Dim sError
   If Trim(SelectValue) <> "" Then
	   'Calling Function To Check Object Enabled
		CheckObjEnabled TIAOracleRadioGroup,sError
		If IsEmpty(sError) = True Then
			TIAOracleRadioGroup.Select SelectValue
		Else
			Reporter.ReportEvent micFail,"TIAOracleRadioGroup","Object Disabled. Cannot Click."
		End If
   End If
End Function


Function TIAOracleButtonClick(TIAOracleButton)
   Dim sError
   'Calling Function To Check Object Enabled
	CheckObjEnabled TIAOracleButton,sError
	If IsEmpty(sError) = True  Then
		TIAOracleButton.Click
	Else
		Reporter.ReportEvent micFail,"TIAOracleButton","Object Disabled. Cannot Click."
	End If
End Function

Function CheckObjEnabled(OracleObject,sError)
   sError = "Object Not Enabled"
   For i=1 to 60
		If  OracleObject.GetROProperty("enabled") = True Then
			sError = Empty
			Exit For
		End If
		Wait(2)
   Next
End Function

Function CloseAllBrowser()
		Dim strSQL, oWMIService, ProcColl, oElem
		strSQL = "Select * From Win32_Process Where Name = 'iexplore.exe' OR Name = 'firefox.exe'"
		Set oWMIService = GetObject("winmgmts:\\.\root\cimv2")
		Set ProcColl = oWMIService.ExecQuery(strSQL)
		For Each oElem in ProcColl
		    oElem.Terminate
		Next
		Set oWMIService = Nothing
End function

Function WriteRecordSet (DataSourcePath,SQL)
	Set ConnToWrite = CreateObject("ADODB.Connection") ' create connection object
	'Create Connection 
	With ConnToWrite ' build connection string and open connection
	.Provider = "Microsoft.Jet.OLEDB.4.0"
	.ConnectionString = "Data Source=" & DataSourcepath & ";" & "Extended Properties=Excel 8.0;"
	.Open
	End With
	ConnToWrite.Execute Sql 'Execute the SQL recd as input param
	ConnToWrite.Close ' close connection
	Set ConnToWrite = Nothing ' release connection object
End Function

Function ConvertRecordsetToArray(rs,sData)
   Dim i
   For i = 0 to rs.fields.count-1
		If rs.Fields(i) <> "" Then
			sString = sString&"|"&rs.fields(i).Name &":=" &rs.fields(i)
		End If
   Next
   sData = Split(sString,"|",-1,1)
End Function

Sub Preference()
   Dim i
	Do 
		For i = 10 TO 0 Step -1
			If OracleFormWindow("index:="&i).Exist(1) Then
				If OracleFormWindow("index:="&i).GetROProperty("short title") <> "StartUp" Then
					OracleFormWindow("index:="&i).CloseWindow
				Else
					Exit Do
				End If
			End If
		Next
		If i = 0 Then
			Exit Do
		End If
	Loop
End Sub


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''Function ID - START

	''Purpose of the Function -	 If the file does not exist it will create the file and write the status in the file. If the file exist then it will update the status in the file. The File is located at the root of the QTPResults folder.
	''Input Parameters = Status,PropertyFileLocation
	''Return Value - None
	''Sample Function Call - WtiteExecutionTextFile "Passed","C:\QTPResults"
	''Created/Updated by: : Sunil
	
Public Function WriteExecutionTextFile(Status,PropertyFileLocation)
	Dim ExecutionStatus, myFSO, WriteStatus 'Variable Decleration
	WriteStatus = Status
	Set myFSO = CreateObject("Scripting.FileSystemObject")
	If (myFSO.FileExists(PropertyFileLocation&"\ExecutionStatus.txt")) Then
		myFSO.DeleteFile(PropertyFileLocation&"\ExecutionStatus.txt")
		Set ExecutionStatus = myFSO.CreateTextFile(PropertyFileLocation&"\ExecutionStatus.txt", True)
		ExecutionStatus.WriteLine(WriteStatus)
		ExecutionStatus.Close
	Else
		Set ExecutionStatus = myFSO.CreateTextFile(PropertyFileLocation&"\ExecutionStatus.txt", True)
		ExecutionStatus.WriteLine(WriteStatus)
		ExecutionStatus.Close
		
	End If
	Set myFSO = nothing
	Set ExecutionStatus =  nothing
End Function

'''Function ID - END
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Function GetRecordSet(DataSourcepath,SQL,RS)
	Dim Conn
	Set Conn = CreateObject("ADODB.Connection") ' create connection object
	'Create Connection 
		With Conn ' build connection string and open connection
			.Provider = "Microsoft.Jet.OLEDB.4.0"
			.ConnectionString = "Data Source=" & DataSourcepath & ";" & "Extended Properties=Excel 8.0;"
			.Open
		End With
	Set RS = CreateObject("ADODB.Recordset") ' create recordset object
	RS.open SQL, Conn ' open recordset using the connection object 
	Set Conn = Nothing
End Function



Function TIAOracleListSelect(TIAOracleList,SelectList)
   If Trim(SelectList) <> "" Then
	   TIAOracleList.Select Cstr(SelectList)
   End If
End Function

 Function TIAOracleTextFieldEnter(TIAOracleTextField,TextFieldValue)
   Dim sError
   If Trim(TextFieldValue) <> "" Then
	   'Calling Function To Check Object Enabled
		CheckObjEnabled TIAOracleTextField,sError
		If IsEmpty(sError) = True  Then
			'TIAOracleTextField.Click
			TIAOracleTextField.Enter TextFieldValue
		Else
			Reporter.ReportEvent micFail,"TIAOracleTextFieldEnter","Object Disabled. Cannot Enter Value"
		End If
   End If
End Function

Function TIAOracleCheckBoxSelect(TIAOracleCheckBox)
   Dim sError
   'Calling Function To Check Object Enabled
	CheckObjEnabled TIAOracleCheckBox,sError
	If IsEmpty(sError) = True  Then
		TIAOracleCheckBox.Select
	Else
		Reporter.ReportEvent micFail,"TIAOracleCheckBox","Object Disabled. Cannot Click."
	End If
End Function

Function TIAOracleRadioGroupSelect(TIAOracleRadioGroup,SelectValue)
   Dim sError
   If Trim(SelectValue) <> "" Then
	   'Calling Function To Check Object Enabled
		CheckObjEnabled TIAOracleRadioGroup,sError
		If IsEmpty(sError) = True Then
			TIAOracleRadioGroup.Select SelectValue
		Else
			Reporter.ReportEvent micFail,"TIAOracleRadioGroup","Object Disabled. Cannot Click."
		End If
   End If
End Function


'Function TIAOracleButtonClick(TIAOracleButton)
'   Dim sError
'   'Calling Function To Check Object Enabled
'	CheckObjEnabled TIAOracleButton,sError
'	If IsEmpty(sError) = True  Then
'		TIAOracleButton.Click
'	Else
'		Reporter.ReportEvent micFail,"TIAOracleButton","Object Disabled. Cannot Click."
'	End If
'End Function

Function TIAOracleButtonClick(TIAOracleButton)
   Dim sError
   'Calling Function To Check Object Enabled
                CheckObjEnabled TIAOracleButton,sError
                If IsEmpty(sError) = True  Then
'                                SectionName = OracleFormWindow("index:=0").GetROProperty("short title")
'                                sFileName = Environment("SnapDir")&"\Snapshots\" &"SNAPSHOT_"&Environment("PicNo")&".png"
'                                Environment("PicNo") = Environment("PicNo")+1
'                                OracleApplications("index:=0").CaptureBitmap sFileName,True
                                TIAOracleButton.Click
                Else
                                Reporter.ReportEvent micFail,"TIAOracleButton","Object Disabled. Cannot Click."
                End If
End Function


Function CheckObjEnabled(OracleObject,sError)
   sError = "Object Not Enabled"
   For i=1 to 60
		If  OracleObject.GetROProperty("enabled") = True Then
			sError = Empty
			Exit For
		End If
		Wait(2)
   Next
End Function

Function CloseAllBrowser()
		Dim strSQL, oWMIService, ProcColl, oElem
		strSQL = "Select * From Win32_Process Where Name = 'iexplore.exe' OR Name = 'firefox.exe'"
		Set oWMIService = GetObject("winmgmts:\\.\root\cimv2")
		Set ProcColl = oWMIService.ExecQuery(strSQL)
		For Each oElem in ProcColl
		    oElem.Terminate
		Next
		Set oWMIService = Nothing
End function

Function WriteRecordSet (DataSourcePath,SQL)
	Set ConnToWrite = CreateObject("ADODB.Connection") ' create connection object
	'Create Connection 
	With ConnToWrite ' build connection string and open connection
	.Provider = "Microsoft.Jet.OLEDB.4.0"
	.ConnectionString = "Data Source=" & DataSourcepath & ";" & "Extended Properties=Excel 8.0;"
	.Open
	End With
	ConnToWrite.Execute Sql 'Execute the SQL recd as input param
	ConnToWrite.Close ' close connection
	Set ConnToWrite = Nothing ' release connection object
End Function

Function ConvertRecordsetToArray(rs,sData)
   Dim i
   For i = 0 to rs.fields.count-1
		If rs.Fields(i) <> "" Then
			sString = sString&"|"&rs.fields(i).Name &":=" &rs.fields(i)
		End If
   Next
   sData = Split(sString,"|",-1,1)
End Function

Sub Preference()
   Dim i
	Do 
		For i = 10 TO 0 Step -1
			If OracleFormWindow("index:="&i).Exist(1) Then
				If OracleFormWindow("index:="&i).GetROProperty("short title") <> "StartUp" Then
					OracleFormWindow("index:="&i).CloseWindow
				Else
					Exit Do
				End If
			End If
		Next
		If i = 0 Then
			Exit Do
		End If
	Loop
End Sub


Function fnRandomNumberWithDateTimeStamp()
		'Find out the current date and time
		Dim sDate : sDate = Day(Now)&"-"
		'Dim sMonth : sMonth = Month(Now)
		Dim sMonth : sMonth = MonthName(Month(Now),True)&"-"
		Dim sYear : sYear = Year(Now)&"_"
		Dim sHour : sHour = Hour(Now)&":"
		Dim sMinute : sMinute = Minute(Now)&":"
		Dim sSecond : sSecond = Second(Now)
		'Create Random Number
		'fnRandomNumberWithDateTimeStamp = Int(sDate & sMonth & sYear & sHour & sMinute & sSecond)
		fnRandomNumberWithDateTimeStamp = sDate & sMonth & sYear & sHour & sMinute & sSecond
End Function

Function fnLongRandomNumber(LengthOfRandomNumber)

Dim sMaxVal : sMaxVal = ""
Dim iLength : iLength = LengthOfRandomNumber

'Find the maximum value for the given number of digits
For iL = 1 to iLength-1
sMaxVal = sMaxVal & "9"
Next
sMaxVal = Int(sMaxVal)

'Find Random Value
 iTmp = Int(((Second(Now)*sMaxVal) * Rnd) + 1)
'Add Trailing Zeros if required
iLen = Len(iTmp)
fnLongRandomNumber = "4"& iTmp 

End Function

Function fnRandomNumber()

		Dim sHour : sHour = Hour(Now)
		Dim sMinute : sMinute = Minute(Now)
		Dim sSecond : sSecond = Second(Now)
		Dim rNum : rNum = Int((Rnd * 10)+1)
		Dim Num : Num = sHour & sMinute & sSecond &rNum
		fnRandomNumber = Num

End Function

Function IDNumber(Dob)
Dim IDN1,IDN2,IDN3,IDN4,IDN5,GEN, IDNum
		
		IDN1=Trim(Mid(Dob,1,6))
		GEN=Trim(Mid(Dob,7,1))
		If UCase(GEN)="M" Then
			IDN2=5
		Else
			IDN2=0
		End If
		For i=1 To Second(Now)
			IDN3=Int((999 - 100 + 1) * Rnd + 100)
		Next
		IDN4="08"
		IDN=Int(IDN1&IDN2&IDN3&IDN4)
		OddSum=0
		For i=1 To 11 Step 2
			OddPos=Mid(IDN,i,1)
			OddSum=OddSum+OddPos
		Next
		EvenNum=""
		For i=2 To 12 Step 2
			EvenPos=Mid(IDN,i,1)
			EvenNum=EvenNum&EvenPos
		Next
		EvenVal=Int(EvenNum)*2
		EvenValSum=0
		For i=1 to Len(EvenVal)
			EvenValPos=Int(Mid(EvenVal,i,1))
			EvenValSum=EvenValSum+EvenValPos
		Next
		OddEvenSum=OddSum+EvenValSum
		IDN5=10-Mid(OddEvenSum,2,1)
		If IDN5=10 Then
			Do
				For i=1 To Second(Now)
					IDN3=Int((999 - 100 + 1) * Rnd + 100)
				Next
				IDN=Int(IDN1&IDN2&IDN3&IDN4)
				OddSum=0
				For i=1 To 11 Step 2
					OddPos=Mid(IDN,i,1)
					OddSum=OddSum+OddPos
				Next
				EvenNum=""
				For i=2 To 12 Step 2
					EvenPos=Mid(IDN,i,1)
					EvenNum=EvenNum&EvenPos
				Next
				EvenVal=Int(EvenNum)*2
				EvenValSum=0
				For i=1 to Len(EvenVal)
					EvenValPos=Int(Mid(EvenVal,i,1))
					EvenValSum=EvenValSum+EvenValPos
				Next
				OddEvenSum=OddSum+EvenValSum
				IDN5=10-Mid(OddEvenSum,2,1)
			Loop Until IDN5<>10
		End If
		IDNum=Int(IDN1&IDN2&IDN3&IDN4&IDN5)
		IDNumber=IDNum

End Function

Function MultipleWindowHandler(oScreen) 
	 winCounter = 0
	 MultipleWindowHandler = 1 ' Initially fail 
	If Not IsObject(oScreen) Then Exit Function
		oScreen.SetTOProperty "index", winCounter
	 Do
		  rc = oScreen.OracleButton("developer name:=TIA_BUTTON_BLOCK2_ACCEPT_BUTTON_0").GetROProperty("enabled")
		  If rc <> True Then
			   winCounter = winCounter +1
			   oScreen.SetTOProperty "index", winCounter
		  Else
			MultipleWindowHandler = 0 ' flag to determide pass or fail 
		   Exit Do
		  End If
		 Loop Until winCounter > 2 
End Function 'MultipleWindowHandler


Function CheckError()
		'For checking Link Address Form and Message errors
				If OracleFormWindow("short title:=Link Risk Address","index:=1").Exist(10) Then
						OracleFormWindow("short title:=Link Risk Address","index:=1").OracleButton("developer name:=TIA_BUTTON_BLOCK2_ACCEPT_BUTTON_0").Click
						If OracleNotification("title:=Message").Exist(3) Then
							PartyIDError = OracleNotification("title:=Message").GetROProperty("message")
							Reporter.ReportEvent micFail,"Step 17 CreateParty","Error:" &PartyIDError
							OracleNotification("title:=Message").OracleButton("label:=OK").Click
						End If	
				End If
		'For handling Forms messages
				If OracleNotification("title:=Forms").Exist(3) Then
						If Instr(OracleNotification("title:=Forms").GetRoProperty("message"),"Do you want to save the changes you have made?") <> 0 Then
							OracleNotification("title:=Forms").OracleButton("label:=Yes").Click
						End If
						If OracleNotification("title:=Forms").Exist(3) Then
							OracleNotification("title:=Forms").OracleButton("label:=OK").Click
						End If
				End If
			   If  OracleFormWindow("title:=.*f7m08.*").Exist(1) Then
				'OracleFormWindow("f7m08").Activate
				OracleFormWindow("title:=.*f7m08.*").OracleButton("label:=OK").Click
			   End If
		'For Checking Message Notifications
				If OracleNotification("title:=Message").Exist(3)  Then
					OracleNotification("title:=Message").OracleButton("label:=OK").Click
				End If
		'For handling Standard Error Form - f0x00
			If OracleFormWindow("short title:=Standard Error Form").Exist(1) Then
				OracleFormWindow("short title:=Standard Error Form").OracleButton("developer name:=HEADER_OK_0").Click
			End If
'		'For handling List of Messages on Save Policy/Quote
'			While OracleFormWindow("short title:=List of Messages").Exist(1) 'Then
'				Wait 1
'				OracleFormWindow("short title:=List of Messages").OracleButton("developer name:=HEADER_OK_0").Click	
'			Wend
'		'Security messages when adding a car
'			While OracleFormWindow("short title:=List of Messages").Exist(1) 'Then
'				Wait 1
'				OracleFormWindow("short title:=List of Messages").OracleButton("developer name:=HEADER_OK_0").Click
'			Wend
'		'For checking messages with Errors and Warnings
'			While OracleNotification("title:=Message").Exist(1)  'Then
'						msg_detail = OracleNotification("title:=Message").GetROProperty ("message")
'						msg_type = OracleNotification("title:=Message").GetROProperty ("type")
'						a_result = UCASE(LEFT(msg_detail,7))
'				  If a_result = "WARNING" or a_result ="Decision"Then
'					   Reporter.ReportEvent micWarning, "Warning", "Warning Message was displayed : "&a
'					   Wait 1
'					  OracleNotification("title:=Message").Approve
'				 Elseif  msg_type ="Info" and  a_result <>"ERROR: " and  a_result <>"ACCESS " Then 
'						Wait 1
'						OracleNotification("title:=Message").Approve
'				Else
'					Wait 1
'					OracleNotification("title:=Message").Approve 
'				End If
'			Wend
'		'For handling Stop notification
'			While OracleNotification("title:=Stop").Exist(1) 
'				rc = OracleNotification("title:=Stop").GetROProperty("message")
'					If rc = "Do you want to save the changes you have made?" Then
'						Wait 1
'						OracleNotification("title:=Stop").Approve
'					ElseIf rc = "Could this loss be as a result of a catastrophe?" Then
'							Wait 1
'							OracleNotification("title:=Stop").Approve
'					ElseIf rc = "Question is mandatory. Do you want to correct?" Then
'							Wait 1
'							OracleNotification("title:=Stop").Approve
'					ElseIf Instr(1, rc, "Estimate is zero")>0 Then
'							Wait 1
'							OracleNotification("title:=Stop").Approve
'					Else
'						Wait 1
'					End If
'			Wend
		'For checking and entering the effective date
			If OracleFormWindow("short title:=Name Information").Exist(5)  Then
				EffectiveDate = OracleFormWindow("short title:=Name Information").OracleTextField("developer name:=NAME_X_D05_0").GetROProperty("value")
				If EffectiveDate="" Then
				OracleFormWindow("short title:=Name Information").OracleTextField("developer name:=NAME_X_D05_0").Enter Day(Date)&"."&Month(Date)&"."&Year(Date)
				End If
			End If 
End Function