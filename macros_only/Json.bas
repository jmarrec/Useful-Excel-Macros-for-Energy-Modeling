Attribute VB_Name = "Json"
Sub Export_To_JSON()
' Author: Julien Marrec
' Date: 2016-04-28
' Modified: 2016-07-21
' Version: 1.1
' References: Microsoft Forms Object 2.0, needed for using the Clipboard

' Export current region to a JSON array of hash you can paste in ruby for example

Dim wS As Worksheet
Set wS = ActiveSheet

Dim rS%, rE%, cS%, cE%
Dim i%, j%
Dim Class As String
Dim s As String
Dim Answer

Dim DefaultRange As Range
Set DefaultRange = Selection

' If the selected range is only one cell, then the default range is set to the current region
If DefaultRange.Cells.Count = 1 Then
    Set DefaultRange = DefaultRange.CurrentRegion
End If


' Define active region (currentregion = same as CTRL +A) by start and end lines and rows
With DefaultRange

    rS = .Rows(1).row
    rE = rS + .Rows.Count - 1
    cS = .Columns(1).Column
    cE = cS + .Columns.Count - 1
    
End With


' Initialize s
s = "{ ""data"": ["

' Will set the Hash keys as the first line, so we start iterating on the last
For i = rS + 1 To rE Step 1
    
    s = s & "{"
    
    For j = cS To cE Step 1
    
        hash_key = s & """" & wS.Cells(rS, j).Value & """"
        
        If IsEmpty(wS.Cells(i, j).Value) Then
            hash_value = 0
        ElseIf IsNumeric(wS.Cells(i, j).Value) Then
            hash_value = wS.Cells(i, j).Value
        Else
            hash_value = """" & wS.Cells(i, j).Value & """"
        End If
        
    
        If j < cE Then
            s = hash_key & ":" & hash_value & ","
        ElseIf j = cE Then
            s = hash_key & ":" & hash_value
        End If
        
        
    Next j
    
    If i < rE Then
        s = s & "}," & vbCrLf
    ElseIf i = rE Then
        s = s & "}" & vbCrLf
    End If

Next i

s = s & "] }"

Debug.Print s



' Ask for saving method: write as a file or copy to clipboard
Answer = MsgBox(Prompt:="Click Yes to save it as a json or txt file, and click No to copy it in the clipboard", Buttons:=vbYesNo, Title:="Saving Method")

If Answer = vbYes Then

    strPath = Application.GetSaveAsFilename(InitialFileName:=Class, FileFilter:="Json (*.json), *.json, Text Files (*.txt), *.txt", Title:="Save output string")
    
    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = oFSO.CreateTextFile(strPath)
    oFile.WriteLine s
    oFile.Close
    
    ' Clean-up objects
    Set oFSO = Nothing
    Set oFile = Nothing


ElseIf Answer = vbNo Then

    ' Put generated string to Clipboard
    Dim MyDataObj As New DataObject
    MyDataObj.SetText s
    MyDataObj.PutInClipboard

    MsgBox Prompt:="Copied to clipboard", Buttons:=vbInformation, Title:="Success"
    
End If

End Sub
