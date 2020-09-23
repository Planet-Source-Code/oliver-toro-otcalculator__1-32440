VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Begin VB.Form FrmCalculator 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OTCalculator - Using Microsot Script Control"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   Icon            =   "FrmCalculator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Vars List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   135
      TabIndex        =   5
      Top             =   450
      Width           =   5460
      Begin VB.CommandButton CmdHelp 
         Caption         =   "Help"
         Height          =   510
         Left            =   4140
         Picture         =   "FrmCalculator.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   600
      End
      Begin VB.CommandButton CmdReset 
         Caption         =   "Reset"
         Height          =   510
         Left            =   3555
         Picture         =   "FrmCalculator.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   600
      End
      Begin VB.TextBox TxtVarName 
         Height          =   285
         Left            =   900
         TabIndex        =   11
         Top             =   225
         Width           =   1950
      End
      Begin VB.CommandButton CmdDim 
         Caption         =   "Dim"
         Height          =   510
         Left            =   2970
         Picture         =   "FrmCalculator.frx":0B16
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   600
      End
      Begin VB.ListBox VarList 
         Height          =   1620
         Left            =   180
         TabIndex        =   9
         Top             =   1035
         Width           =   5190
      End
      Begin VB.TextBox TxtValue 
         Height          =   285
         Left            =   900
         TabIndex        =   8
         Top             =   540
         Width           =   1950
      End
      Begin VB.Label Label2 
         Caption         =   "Var Name"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   225
         Width           =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         X1              =   5355
         X2              =   135
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label4 
         Caption         =   "Value"
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   540
         Width           =   870
      End
   End
   Begin VB.TextBox TxtResult 
      BackColor       =   &H00FFFED0&
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   765
      TabIndex        =   3
      Top             =   4905
      Width           =   4335
   End
   Begin VB.CommandButton CmdSolve 
      Caption         =   "Eval"
      Height          =   510
      Left            =   5130
      Picture         =   "FrmCalculator.frx":10A0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   555
   End
   Begin VB.TextBox TxtFormula 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1275
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FrmCalculator.frx":162A
      Top             =   3600
      Width           =   4965
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   5130
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      Height          =   375
      Left            =   180
      Top             =   45
      Width           =   5325
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "You MUST to add reference: Microsoft Script Control(MSSCRIPT.OCX)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   180
      TabIndex        =   14
      Top             =   135
      Width           =   5370
   End
   Begin VB.Label Label3 
      Caption         =   "Result:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   90
      TabIndex        =   4
      Top             =   4905
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "Expression:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   135
      TabIndex        =   1
      Top             =   3330
      Width           =   1635
   End
End
Attribute VB_Name = "FrmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ContVar As Integer

Private Sub CmdDim_Click()
    Dim StrLine As String
    Dim StrSave As String
    
    If Trim(TxtVarName.Text) = "" Then
        MsgBox "You must fill the VarName...", vbCritical
        TxtVarName.SetFocus
        Exit Sub
    End If
    If Trim(TxtValue.Text) = "" Then
        MsgBox "You must fill the value...", vbCritical
        TxtValue.SetFocus
        Exit Sub
    End If
    StrLine = "DIM " & Trim(TxtVarName.Text)
    StrSave = Trim(TxtVarName.Text) & " = " & Trim(TxtValue.Text)
    
    VarList.AddItem StrLine
    VarList.AddItem StrSave
    
    ScriptControl1.AddCode StrLine
    ScriptControl1.AddCode StrSave
End Sub

Private Sub CmdHelp_Click()
frmAbout.Show vbModal
End Sub

Private Sub CmdReset_Click()
    VarList.Clear
    ScriptControl1.Reset
    TxtResult.Text = ""
    TxtValue.Text = ""
    TxtVarName.Text = ""
End Sub

Private Sub CmdSolve_Click()
    TxtResult.Text = ScriptControl1.Eval(TxtFormula.Text)
End Sub

Private Sub Form_Load()
    ContVar = 1
    ScriptControl1.AddCode "Dim a"
    VarList.AddItem "Dim a"
    ScriptControl1.AddCode "a = 1"
    VarList.AddItem "a = 1"
    ScriptControl1.AddCode "Dim b"
    VarList.AddItem "Dim b"
    ScriptControl1.AddCode "b = 2"
    VarList.AddItem "b = 2"
End Sub

