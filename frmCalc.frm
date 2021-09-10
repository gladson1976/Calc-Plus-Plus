VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calc ++"
   ClientHeight    =   3345
   ClientLeft      =   405
   ClientTop       =   390
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCalc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTotal 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3135
      ScaleWidth      =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   4800
      Begin VB.PictureBox picOuter 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3090
         Left            =   0
         ScaleHeight     =   3090
         ScaleWidth      =   4800
         TabIndex        =   4
         Top             =   0
         Width           =   4800
         Begin VB.PictureBox picInner 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            FillColor       =   &H80000005&
            FillStyle       =   0  'Solid
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   90
            ScaleHeight     =   570
            ScaleWidth      =   4725
            TabIndex        =   5
            Top             =   0
            Width           =   4725
            Begin VB.TextBox txtGoto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   225
               Left            =   120
               TabIndex        =   7
               Top             =   240
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.TextBox txtA 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               ForeColor       =   &H00000000&
               Height          =   225
               Index           =   0
               Left            =   15
               Locked          =   -1  'True
               TabIndex        =   1
               Top             =   255
               Width           =   4665
            End
            Begin VB.TextBox txtQ 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               ForeColor       =   &H00000000&
               Height          =   225
               Index           =   0
               Left            =   375
               TabIndex        =   0
               Top             =   0
               Width           =   4305
            End
            Begin VB.Label lblCounter 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Caption         =   "1"
               ForeColor       =   &H00FF0000&
               Height          =   225
               Index           =   0
               Left            =   0
               MouseIcon       =   "frmCalc.frx":0ABA
               MousePointer    =   99  'Custom
               TabIndex        =   6
               Top             =   0
               Width           =   300
            End
            Begin VB.Line linWhite 
               BorderColor     =   &H00FFFFFF&
               Index           =   0
               X1              =   0
               X2              =   4665
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Line linBlack 
               BorderColor     =   &H00808080&
               Index           =   0
               X1              =   0
               X2              =   4665
               Y1              =   225
               Y2              =   225
            End
         End
      End
   End
   Begin VB.VScrollBar vscrCalc 
      Height          =   2295
      LargeChange     =   5
      Left            =   4920
      Max             =   1
      Min             =   1
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Value           =   1
      Width           =   255
   End
   Begin MSScriptControlCtl.ScriptControl scrCalc 
      Left            =   4080
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   0   'False
   End
   Begin VB.Image imgOptions 
      Height          =   240
      Left            =   4920
      MouseIcon       =   "frmCalc.frx":0DC4
      MousePointer    =   99  'Custom
      Picture         =   "frmCalc.frx":10CE
      ToolTipText     =   "Options"
      Top             =   2580
      Width           =   240
   End
   Begin VB.Image imgQuit 
      Height          =   240
      Left            =   4920
      MouseIcon       =   "frmCalc.frx":1658
      MousePointer    =   99  'Custom
      Picture         =   "frmCalc.frx":1962
      ToolTipText     =   "Press [ALT]+X to Exit"
      Top             =   3000
      Width           =   240
   End
   Begin VB.Line linV 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   4810
      X2              =   4810
      Y1              =   120
      Y2              =   3240
   End
   Begin VB.Line linV 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   4800
      X2              =   4800
      Y1              =   120
      Y2              =   3240
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
If boolApplicationStart = True Then
    Call initSettings
    boolApplicationStart = False
End If
'frmCalc.txtQ(intCalcCount - 1).SetFocus
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Debug.Print Shift & ", " & KeyCode
If Shift = 4 And KeyCode = 88 Then
    ' [ALT]+X
    Call imgQuit_Click
End If
If Shift = 2 And KeyCode = 71 Then
    ' [CTRL]+G
    'vscrCalc.Value = intCurrentCalc + 1
    'txtQ(intCurrentCalc).SetFocus
    Call lblCounter_Click(intCurrentCalc)
End If
End Sub
Private Sub Form_Load()
Call initCalc
End Sub
Private Sub imgOptions_Click()
Load frmOptions
frmOptions.Show vbModal, Me
End Sub
Private Sub imgQuit_Click()
Unload Me
End
End Sub
Private Sub lblCounter_Click(Index As Integer)
txtGoto.Top = lblCounter(Index).Top
txtGoto.Left = lblCounter(Index).Left
txtGoto.Height = lblCounter(Index).Height
txtGoto.Width = lblCounter(Index).Width + 60
txtGoto.Text = lblCounter(Index).Caption
txtGoto.Visible = True
txtGoto.SetFocus
SendKeys "{HOME}+{END}"
End Sub
Private Sub txtA_GotFocus(Index As Integer)
intCurrentCalc = Index
End Sub
Private Sub txtGoto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsNumeric(txtGoto.Text) And Val(txtGoto.Text) <= intCalcCount Then
        vscrCalc.Value = txtGoto.Text
        txtQ(txtGoto.Text - 1).SetFocus
        txtGoto.Visible = False
    End If
End If
If KeyAscii = 27 Then
    txtGoto.Visible = False
End If
End Sub
Private Sub txtQ_Change(Index As Integer)
On Error Resume Next
frmCalc.txtA(Index).Text = frmCalc.scrCalc.Eval(txtQ(Index).Text)
' 6 - Overflow
' 1006 - Missing ) paranthesis

If Err.Number <> 0 And Err.Number <> 6 And Err.Number <> 1006 Then
    frmCalc.txtA(Index).Text = Err.Number & Err.Description
ElseIf Err.Number = 6 Then
    frmCalc.txtA(Index).Text = "Overflow"
ElseIf Err.Number = 1006 Then
    frmCalc.txtA(Index).Text = ""
End If
End Sub
Private Sub txtQ_GotFocus(Index As Integer)
intCurrentCalc = Index
End Sub
Private Sub txtQ_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 And intCalcCount < 480 Then
'    Load frmCalc.txtQ(intCalcCount)
'    Load frmCalc.txtA(intCalcCount)
'    Load frmCalc.lblCounter(intCalcCount)
'    Load frmCalc.linBlack(intCalcCount)
'    Load frmCalc.linWhite(intCalcCount)
'
'    With frmCalc.txtQ(intCalcCount)
'        .Top = frmCalc.txtQ(intCalcCount - 1).Top + intCalcHeight
'        .Left = frmCalc.txtQ(intCalcCount - 1).Left
'        .Text = ""
'        .Visible = True
'        .SetFocus
'    End With
'    With frmCalc.txtA(intCalcCount)
'        .Top = frmCalc.txtA(intCalcCount - 1).Top + intCalcHeight
'        .Left = frmCalc.txtA(intCalcCount - 1).Left
'        .Text = ""
'        .Visible = True
'    End With
'    With frmCalc.lblCounter(intCalcCount)
'        .Top = frmCalc.lblCounter(intCalcCount - 1).Top + intCalcHeight
'        .Left = frmCalc.lblCounter(intCalcCount - 1).Left
'        .Caption = (intCalcCount + 1)
'        .Visible = True
'    End With
'    With frmCalc.linBlack(intCalcCount)
'        .Y1 = frmCalc.linBlack(intCalcCount - 1).Y1 + intCalcHeight
'        .Y2 = frmCalc.linBlack(intCalcCount - 1).Y2 + intCalcHeight
'        .X1 = frmCalc.linBlack(intCalcCount - 1).X1
'        .Visible = True
'    End With
'    With frmCalc.linWhite(intCalcCount)
'        .Y1 = frmCalc.linWhite(intCalcCount - 1).Y1 + intCalcHeight
'        .Y2 = frmCalc.linWhite(intCalcCount - 1).Y2 + intCalcHeight
'        .X1 = frmCalc.linWhite(intCalcCount - 1).X1
'        .Visible = True
'    End With
'    intCalcCount = intCalcCount + 1
'
'    frmCalc.picInner.Height = (intCalcCount * intCalcHeight)
'    frmCalc.vscrCalc.Max = intCalcCount
'    If intCalcCount > 6 Then
'        frmCalc.vscrCalc.Value = frmCalc.vscrCalc.Value + 1
'    End If
    Call newCalc
End If
End Sub
Private Sub vscrCalc_Change()
frmCalc.picInner.Top = -((frmCalc.vscrCalc.Value - 1) * intCalcHeight)
End Sub
