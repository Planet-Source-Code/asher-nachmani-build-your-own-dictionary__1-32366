Attribute VB_Name = "ModDic"
Global Lang As String
Global Translate As String
Global NewTextWord As String
Global ReadTranslate As String
Global ReadNewTextWord(0 To 100) As String
Global Request As String
Global EDITWordToDic As String
Global Del As String
Global deltranslet As String
Public Sub AddNewWord(Lang As String, NewWord As String, FirstWord As String)
Dim tmp As String
Dim FileNumber
   FileNumber = FreeFile   ' Get unused file
On Error GoTo 10
 If Dir(App.Path & "\" & Lang & ".dat") = "" Then
   Open App.Path & "\" & Lang & ".dat" For Append As #FileNumber
   Close #FileNumber
 End If
FileNumber = FreeFile   ' Get unused file
Open App.Path & "\" & Lang & ".dat" For Input As #FileNumber
Do While Not EOF(1)
    Input #FileNumber, tmp, ReadTranslate
    If LCase(NewWord) = tmp And FirstWord = ReadTranslate Then
        Close #FileNumber
        MsgBox "this word is existing "
        Exit Sub
    End If
Loop
Close #FileNumber
FileNumber = FreeFile   ' Get unused file
Open App.Path & "\" & Lang & ".dat" For Append As #FileNumber
    Print #FileNumber, LCase(NewWord) & "," & FirstWord
    Close #FileNumber

Exit Sub
exit10:
Exit Sub
10:
MsgBox Err.Description
Resume exit10

End Sub
Public Sub TransletWord(Lang As String, NewWord As String)
Dim count As Integer, tmp As String, tmpb As String
Dim FileNumber
   FileNumber = FreeFile   ' Get unused file

On Error GoTo 10
ReadNewTextWord(0) = ""
count = 0
 If Dir(App.Path & "\" & Lang & ".dat") = "" Then
    MsgBox "this dictionary not find"
   Exit Sub
 End If

Open App.Path & "\" & Lang & ".dat" For Input As #FileNumber
Do While Not EOF(1)
    Input #FileNumber, tmp, tmpb
    If LCase(NewWord) = tmp Then
        ReadNewTextWord(count) = tmpb
        count = count + 1
    End If
Loop
If ReadNewTextWord(0) = "" Then
  Close #FileNumber
  Exit Sub
Else
  ReadNewTextWord(count) = ""
  Close #FileNumber
  Exit Sub
End If

Exit Sub
exit10:
Exit Sub
10:
MsgBox Err.Description
Resume exit10


End Sub

Public Function ChekDic(x As String) As Boolean
 If Dir(App.Path & "\" & x & ".dat") = "" Then
    ChekDic = False
   Exit Function
 Else
   ChekDic = True
   Exit Function
 End If
End Function

Public Sub FillListColumnHeader(LV As ListView)
   Dim colNew As ColumnHeader
   On Error GoTo 10
  
   LV.ColumnHeaders.Clear
    Set colNew = LV.ColumnHeaders.Add(, , "", , 0)
    Set colNew = LV.ColumnHeaders.Add(, , "source", , 1)
    Set colNew = LV.ColumnHeaders.Add(, , "goal", , 1)
    LV.View = 3    ' Set View property to 'Report'.
    LV.ColumnHeaders.Item(1).Width = 0
    LV.ColumnHeaders.Item(2).Width = 2500
   LV.ColumnHeaders.Item(3).Width = 2500

Exit Sub
exit10:
Exit Sub
10:
MsgBox Err.Description
Resume exit10

End Sub
Public Sub FillParitim(LV As ListView, sourse As String)
 Dim NewLine As ListItem, count As Integer
 On Error GoTo 10
 
 LV.ListItems.Clear
 count = 0
     While ReadNewTextWord(count) <> ""
        Set NewLine = LV.ListItems.Add(, , "")
        NewLine.SubItems(1) = sourse
        NewLine.SubItems(2) = ReadNewTextWord(count)
     count = count + 1
    Wend

exit10:
Exit Sub
10:
 MsgBox Err.Description
 Resume exit10
End Sub
Public Sub DelWord(Lang As String, Del As String, FirstWord As String)
Dim tmp As String, tmpb As String
Dim FileNumber, FileNumber1
   FileNumber = FreeFile   ' Get unused file

On Error GoTo 10
Open App.Path & "\" & Lang & ".dat" For Input As #FileNumber
FileNumber1 = FreeFile   ' Get unused file

Open App.Path & "\" & Lang & "1.dat" For Append As #FileNumber1
Do While Not EOF(1)
    Input #FileNumber, tmp, tmpb
    If LCase(Del) <> tmp Or LCase(FirstWord) <> tmpb Then
        Print #FileNumber1, LCase(tmp) & "," & tmpb
    End If
Loop

Close #FileNumber
    Close #FileNumber1
    Kill App.Path & "\" & Lang & ".dat"
    Name App.Path & "\" & Lang & "1.dat" As App.Path & "\" & Lang & ".dat"

Exit Sub
exit10:
Exit Sub
10:
MsgBox Err.Description
Resume exit10

End Sub
