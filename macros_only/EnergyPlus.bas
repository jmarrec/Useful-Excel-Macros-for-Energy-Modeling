Attribute VB_Name = "EnergyPlus"
Sub Export_To_IDF()
' Author: Julien Marrec
' Date: 2014-05-13
' Version: 1.0
' References: Microsoft Forms Object 2.0, needed for using the Clipboard

' Export current region to an IDF readable format. Paste in text format

Dim wS As Worksheet
Set wS = ActiveSheet

Dim rS%, rE%, cS%, cE%
Dim i%, j%
Dim Class As String
Dim s As String


' Define active region (currentregion = same as CTRL +A) by start and end lines and rows
With ActiveCell.CurrentRegion

    rS = .Rows(1).row
    rE = rS + .Rows.Count - 1
    cS = .Columns(1).Column + 1
    cE = cS + .Columns.Count - 2
    
End With


' Try setting the IDF class from the cell two lines above the start of the table
Class = wS.Cells(rS - 2, cS - 1).Value


' Ask for the IDF Class of the object
Class = InputBox(Prompt:="Input Class of object (eg: Zone, Building, BuildingSurface:Detailed", Title:="Object Class", Default:=Class)

' If nothing was entered or "Cancel" pressed, exit the sub
If Class = "" Then Exit Sub
    


' Loop through each column
For j = cS To cE Step 1

    ' Each object (column) starts with the name of the class and a coma
    s = s & Class & ","

    ' Loop on each line
    For i = rS To rE - 1 Step 1
    
        ' Each attribute of the object (line) is separated by a coma
        s = s & vbCrLf & Chr(9) & wS.Cells(i, j).Value & ","
    
    Next i

    ' Finish with a semi-colon to close the object
    s = s & vbCrLf & Chr(9) & wS.Cells(rE, j).Value & ";" & vbCrLf & vbCrLf

Next j


' Ask for saving method: write as a file or copy to clipboard
Answer = MsgBox(Prompt:="Click Yes to save it as an idf or txt file, and click No to copy it in the clipboard", Buttons:=vbYesNo, Title:="Saving Method")

If Answer = vbYes Then

    strPath = Application.GetSaveAsFilename(InitialFileName:=Class, FileFilter:="EnergyPlus IDF Files (*.idf), *.idf, Text Files (*.txt), *.txt", Title:="Save output string")
    
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


Function RH2HumidityRatio_SI(RH_Percent, Temp_C)
' Author: Julien Marrec
' Date: 2014-06-20
' Convers (RH, Temperature) to a humidity ratio.
' SI Coefficients from ASHRAE 2009 - Chapter 1 Psychrometrics

Dim RH#, p_atm#

p_atm = 101325 ' atmospheric pressure Pa

If RH_Percent < 0 Or RH_Percent > 1 Then
    MsgBox Prompt:="RH must be between 0 and 100%, and entered as a value between 0 and 1", Buttons:=vbOKCancel, Title:="Possible Bad RH"
Else
    RH = RH_Percent * 100
End If


p_ws = Temp2pws_SI(Temp_C)

p_w = RH * p_ws / 100

RH2HumidityRatio_SI = 0.621945 * p_w / (p_atm - p_w)

End Function

Function Temp2pws_SI(Temp_C)
' Author: Julien Marrec
' Date: 2014-06-20
' Calculates Saturation pressure at given temperature
' SI Coefficients from ASHRAE 2009 - Chapter 1 Psychrometrics

Dim Temp_K#, C1#, C2#, C3#, C4#, C5#, C6#, C7#, C9#, C10#, C11#, C12#, C13#, c2k

c2k = 273.15
' In both equations:
' pws = saturation pressure, Pa
' T = absolute temperature, K = °C + 273.15

Temp_K = Temp_C + c2k

' Saturation pressure over ICE for the temperature range of -100 to 0°C
' ln pws = C1/T + C2 + C3*T + C4*T^2 + C5*T^3 + C6*T^4 + C7*ln T
C1 = -5674.5359
C2 = 6.3925247
C3 = -0.009677843
C4 = 0.00000062215701
C5 = 2.0747825E-09
C6 = -9.484024E-13
C7 = 4.1635019



' Saturation pressure over LIQUID WATER for the temperature range of 0 to 200°C
' ln pws = C8/T + C9 + C10*T + C11*T^2 + C12*T^3 + C13*ln T
C8 = -5800.2206
C9 = 1.3914993
C10 = -0.048640239
C11 = 0.000041764768
C12 = -0.000000014452093
C13 = 6.5459673


' If bad temperature
If Temp_C < -100 Or Temp_C > 200 Then
    MsgBox Prompt:="Temperature in celsius must be between -100°C and 200°C", Buttons:=vbCritical, Title:="Bad temperature"
   
' If ice: 0-100°C
ElseIf Temp_K < c2k Then
    p_ws = Exp(C1 / Temp_K + C2 + C3 * Temp_K + C4 * Temp_K ^ 2 + C5 * Temp_K ^ 3 + C6 * Temp_K ^ 4 + C7 * Log(Temp_K))

' If liquid water: 0-200°C
ElseIf Temp_K < c2k + 200 Then
    p_ws = Exp(C8 / Temp_K + C9 + C10 * Temp_K + C11 * Temp_K ^ 2 + C12 * Temp_K ^ 3 + C13 * Log(Temp_K))
    
End If

Temp2pws_SI = p_ws

End Function

Function DB_WB2HumidityRatio_SI(RH_Percent, Tdb_C, T_wb_C)
' Author: Julien Marrec
' Date: 2014-06-20
' Convers (RH, Temperature) to a humidity ratio.
' SI Coefficients from ASHRAE 2009 - Chapter 1 Psychrometrics

Dim Tdb_K#, Tdb_K#, C1#, C2#, C3#, C4#, C5#, C6#, C7#, C9#, C10#, C11#, C12#, C13#, c2k#
Dim RH#


If RH_Percent < 0 Or RH_Percent > 1 Then
    MsgBox Prompt:="RH must be between 0 and 100%, and entered as a value between 0 and 1", Buttons:=vbOKCancel, Title:="Possible Bad RH"
Else
    RH = RH_Percent * 100
End If


c2k = 273.15
' In both equations:
' pws = saturation pressure, Pa
' T = absolute temperature, K = °C + 273.15

Temp_K = Temp_C + c2k

' Saturation pressure over ICE for the temperature range of -100 to 0°C
' ln pws = C1/T + C2 + C3*T + C4*T^2 + C5*T^3 + C6*T^4 + C7*ln T
C1 = -5674.5359
C2 = 6.3925247
C3 = -0.009677843
C4 = 0.00000062215701
C5 = 2.0747825E-09
C6 = -9.484024E-13
C7 = 4.1635019



' Saturation pressure over LIQUID WATER for the temperature range of 0 to 200°C
' ln pws = C8/T + C9 + C10*T + C11*T^2 + C12*T^3 + C13*ln T
C8 = -5800.2206
C9 = 1.3914993
C10 = -0.048640239
C11 = 0.000041764768
C12 = -0.000000014452093
C13 = 6.5459673



p_atm = 101325 ' atmospheric pressure Pa


' If bad temperature
If Temp_C < -100 Or Temp_C > 200 Then
    MsgBox Prompt:="Temperature in celsius must be between -100°C and 200°C", Buttons:=vbCritical, Title:="Bad temperature"
   
' If ice: 0-100°C
ElseIf Temp_K < c2k Then
    p_ws = Exp(C1 / Temp_K + C2 + C3 * Temp_K + C4 * Temp_K ^ 2 + C5 * Temp_K ^ 3 + C6 * Temp_K ^ 4 + C7 * Log(Temp_K))

' If liquid water: 0-200°C
ElseIf Temp_K < c2k + 200 Then
    p_ws = Exp(C8 / Temp_K + C9 + C10 * Temp_K + C11 * Temp_K ^ 2 + C12 * Temp_K ^ 3 + C13 * Log(Temp_K))
    
End If

p_w = RH * p_ws / 100

RH2HumidityRatio_SI = 0.621945 * p_w / (p_atm - p_w)

End Function
