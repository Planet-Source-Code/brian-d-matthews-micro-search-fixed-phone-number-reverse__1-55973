Attribute VB_Name = "Parsing"
'- [Micro_Search] Created By Brian Matthews & keith_escalade!
'-/ The "Parsing/Analyze.bas" was provided in part
'-/ by the team of http://yahpro.org, [keith_escalade]
'-/ Coded By Brian Matthews
'-/ Copyright 2004 Microsoft-Murder(C)
Function Get_Between(Start As Integer, Data As String, Before As String, After As String)
On Error GoTo ErrorHandler
Dim Before1, After1 As Integer
Before1 = InStr(Start, Data, Before)
After1 = InStr(Before1 + Len(Before), Data, After)
Get_Between = Mid(Data, Before1, After1 - Before1)
Get_Between = Right(Get_Between, Len(Get_Between) - Len(Before))
Exit Function
ErrorHandler:
Err.Raise 1, , "Cannot Find String"
End Function
Function Get_Item(ItemNumber As String, Data As String, Delimeter As String)
On Error GoTo ErrorHandler
Get_Item = Split(Data, Delimeter)(ItemNumber - 1)
Exit Function
ErrorHandler:
Err.Raise 2, , "Cannot Find Item"
End Function
Function Right_Of(Start As Integer, Data As String, Text As String)
Dim LeftText As String
On Error GoTo ErrorHandler
If InStr(Start, Data, Text) = 0 Then GoTo ErrorHandler
LeftText = InStr(Start, Data, Text)
LeftText = Left(Data, LeftText + Len(Text) - 1)
Right_Of = Right(Data, Len(Data) - Len(LeftText))
Exit Function
ErrorHandler:
Err.Raise 3, , "Cannot Get Right Of String"
End Function
Function Left_Of(Start As Integer, Data As String, Text As String)
Dim RightText As String
On Error GoTo ErrorHandler
If InStr(Start, Data, Text) = 0 Then GoTo ErrorHandler
RightText = InStr(Start, Data, Text)
RightText = Right(Data, Len(Data) - RightText + 1)
Left_Of = Left(Data, Len(Data) - Len(RightText))
Exit Function
ErrorHandler:
Err.Raise 4, , "Cannot Get Left Of String"
End Function

