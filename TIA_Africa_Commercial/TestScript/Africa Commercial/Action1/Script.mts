



'Created by: Warren Carr
'Created on: 11/05/2016
'Script Name: TIA Africa Commercial


If Trim(DataTable("RUNSKIP"))="RUN" Then
	ExecuteAction=DataTable("Action")
	Execute ExecuteAction
End If


 
