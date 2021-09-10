Attribute VB_Name = "modCalc"
Public boolApplicationStart As Boolean
Public objQ As TextBox, objA As TextBox
Public intCalcCount As Long
Public intCalcHeight As Long
Public intCurrentCalc As Integer

Public boolSaveOnExit As Boolean
Public arrSavedCalcs() As String
Dim I As Integer

Public Sub initCalc()
intCalcCount = 0
intCalcHeight = 510
intCurrentCalc = 1
boolApplicationStart = True
Call readSettings
End Sub

Public Function initSettings()
'MsgBox arrSavedCalcs(1)
If intCalcCount = 0 And arrSavedCalcs(1) = "" Then
    Exit Function
End If
For I = 1 To UBound(arrSavedCalcs)
    Call newCalc
    frmCalc.txtQ(intCalcCount - 1).Text = arrSavedCalcs(I)
Next I
Call newCalc
End Function

Public Function readSettings()
Dim intFileIn As Integer
Dim strTemp As String
Dim strReadOption As String
strReadOption = "[OPTIONS]"

intFileIn = FreeFile
On Error Resume Next
Open App.Path & "\Calc.ini" For Input As intFileIn
'Debug.Print Err.Number & ", " & Err.Description
If Err.Number = 53 Then 'File not found
    Call writeInitialSettings
    Open App.Path & "\Calc.ini" For Input As intFileIn
End If
On Error GoTo 0

I = 1
While Not EOF(intFileIn)
    Line Input #intFileIn, strTemp
    If strTemp <> "[OPTIONS]" And strTemp <> "[CALCS]" Then
        If strReadOption = "[OPTIONS]" Then
            If strTemp = "Save:0" Then
                boolSaveOnExit = False
            Else
                boolSaveOnExit = True
            End If
        ElseIf strReadOption = "[CALCS]" Then
            ReDim Preserve arrSavedCalcs(I)
            arrSavedCalcs(I) = strTemp
            I = I + 1
        End If
    Else
        ' Change the Mode
        strReadOption = strTemp
        If strTemp = "[CALCS]" Then ReDim arrSavedCalcs(1)
    End If
Wend
End Function
Public Function writeSettings()

End Function
Public Function writeInitialSettings()
Dim intFileOut As Integer
intFileOut = FreeFile
On Error Resume Next
Open App.Path & "\Calc.ini" For Output As intFileOut
Print #intFileOut, "[OPTIONS]"
Print #intFileOut, "Save:0"
Print #intFileOut, "[CALCS]"
Close intFileOut
End Function

Public Function newCalc()
intCalcCount = intCalcCount + 1

Load frmCalc.txtQ(intCalcCount)
Load frmCalc.txtA(intCalcCount)
Load frmCalc.lblCounter(intCalcCount)
Load frmCalc.linBlack(intCalcCount)
Load frmCalc.linWhite(intCalcCount)

With frmCalc.txtQ(intCalcCount)
    .Top = frmCalc.txtQ(intCalcCount - 1).Top + intCalcHeight
    .Left = frmCalc.txtQ(intCalcCount - 1).Left
    .Text = ""
    .Visible = True
    '.SetFocus
End With
With frmCalc.txtA(intCalcCount)
    .Top = frmCalc.txtA(intCalcCount - 1).Top + intCalcHeight
    .Left = frmCalc.txtA(intCalcCount - 1).Left
    .Text = ""
    .Visible = True
End With
With frmCalc.lblCounter(intCalcCount)
    .Top = frmCalc.lblCounter(intCalcCount - 1).Top + intCalcHeight
    .Left = frmCalc.lblCounter(intCalcCount - 1).Left
    .Caption = (intCalcCount + 1)
    .Visible = True
End With
With frmCalc.linBlack(intCalcCount)
    .Y1 = frmCalc.linBlack(intCalcCount - 1).Y1 + intCalcHeight
    .Y2 = frmCalc.linBlack(intCalcCount - 1).Y2 + intCalcHeight
    .X1 = frmCalc.linBlack(intCalcCount - 1).X1
    .Visible = True
End With
With frmCalc.linWhite(intCalcCount)
    .Y1 = frmCalc.linWhite(intCalcCount - 1).Y1 + intCalcHeight
    .Y2 = frmCalc.linWhite(intCalcCount - 1).Y2 + intCalcHeight
    .X1 = frmCalc.linWhite(intCalcCount - 1).X1
    .Visible = True
End With

frmCalc.txtQ(intCalcCount - 1).SetFocus

frmCalc.picInner.Height = (intCalcCount * intCalcHeight)
frmCalc.vscrCalc.Max = intCalcCount
If intCalcCount > 6 Then
    frmCalc.vscrCalc.Value = frmCalc.vscrCalc.Value + 1
End If
End Function
