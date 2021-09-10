VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calc ++ : Options"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4455
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSave 
      Caption         =   "&Save data on exit"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   4275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2400
      Width           =   735
   End
   Begin VB.Line linBlack 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   4440
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line linWhite 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   4440
      Y1              =   2295
      Y2              =   2295
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

