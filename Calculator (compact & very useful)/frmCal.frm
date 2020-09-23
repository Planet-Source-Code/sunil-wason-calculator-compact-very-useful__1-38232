VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCal 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "frmCal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   21.688
   ScaleMode       =   4  'Character
   ScaleWidth      =   55
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2760
      Top             =   1560
   End
   Begin VB.OptionButton optRad 
      Caption         =   "&Rad"
      Height          =   255
      Left            =   960
      TabIndex        =   31
      Top             =   840
      Width           =   615
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   66
      Top             =   4950
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRnd 
      Caption         =   "Int"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      ToolTipText     =   "Rounds Off to the Neareast Value"
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmdPi 
      Caption         =   "Pi"
      Height          =   375
      Left            =   960
      TabIndex        =   45
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdCos 
      Caption         =   "Cosec"
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   35
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdCos 
      Caption         =   "Cot"
      Height          =   375
      Index           =   5
      Left            =   960
      TabIndex        =   37
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdCos 
      Caption         =   "Sec"
      Height          =   375
      Index           =   4
      Left            =   960
      TabIndex        =   36
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdCos 
      Caption         =   "Sin"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   32
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdCos 
      Caption         =   "Tan"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   34
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdCos 
      Caption         =   "Cos"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   33
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "log"
      Height          =   375
      Left            =   960
      TabIndex        =   43
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdEx 
      BackColor       =   &H8000000A&
      DisabledPicture =   "frmCal.frx":0442
      DownPicture     =   "frmCal.frx":3D084
      Height          =   375
      Left            =   960
      Picture         =   "frmCal.frx":79CC6
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdLn 
      Caption         =   "ln"
      Height          =   375
      Left            =   240
      TabIndex        =   39
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdCarat 
      Caption         =   "x^y"
      Height          =   375
      Left            =   240
      TabIndex        =   38
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdMMinus 
      Caption         =   "M-"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      ToolTipText     =   " Subtract Displayed value from Memory (MS) "
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmdMemPlus 
      Caption         =   "M+"
      Height          =   375
      Left            =   4920
      MaskColor       =   &H80000000&
      TabIndex        =   7
      ToolTipText     =   " Add Displayed value to Memory (MS) "
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<--"
      Height          =   375
      Left            =   1920
      TabIndex        =   27
      ToolTipText     =   "Backspace"
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmdKeyBoard 
      Caption         =   "Activate &Keyboard"
      Height          =   336
      Left            =   1800
      TabIndex        =   60
      ToolTipText     =   " Allows the user to enter the Numbers from the Keyboard "
      Top             =   4608
      Width           =   2175
   End
   Begin VB.CommandButton cmdMC 
      Caption         =   "MC"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      ToolTipText     =   " Memory Clear "
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmdMR 
      Caption         =   "MR"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      ToolTipText     =   " Memory Recall "
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton cmdMplus 
      Caption         =   "MS"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      ToolTipText     =   " Memory Save "
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmdOff 
      Caption         =   "O&ff"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      ToolTipText     =   " Switches Off the Calculator "
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   336
      Left            =   4080
      TabIndex        =   61
      ToolTipText     =   " Quit "
      Top             =   4608
      Width           =   2415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      ToolTipText     =   " Clear Contents "
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdOn 
      Caption         =   "&On"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      ToolTipText     =   " Switches On the Calculator "
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdSqrt 
      Caption         =   "sqrt"
      Height          =   375
      Left            =   240
      TabIndex        =   41
      ToolTipText     =   " Square Root "
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdFactorial 
      Caption         =   "!"
      Height          =   375
      Left            =   240
      TabIndex        =   40
      ToolTipText     =   " Factorial "
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdInvert 
      Caption         =   "1/x"
      Height          =   375
      Left            =   960
      TabIndex        =   44
      ToolTipText     =   " Invert "
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdEquals 
      Caption         =   "="
      Height          =   375
      Left            =   4020
      TabIndex        =   14
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmdPlusMinus 
      Caption         =   "+/-"
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmdDecimal 
      Caption         =   "."
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmdDiv 
      Caption         =   "/"
      Height          =   375
      Left            =   4020
      TabIndex        =   10
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmdMul 
      Caption         =   "*"
      Height          =   375
      Left            =   4020
      TabIndex        =   11
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "-"
      Height          =   375
      Left            =   4020
      TabIndex        =   13
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "8"
      Height          =   375
      Left            =   2520
      TabIndex        =   25
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "9"
      Height          =   375
      Left            =   3120
      TabIndex        =   26
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "7"
      Height          =   375
      Left            =   1920
      TabIndex        =   24
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "5"
      Height          =   375
      Left            =   2520
      TabIndex        =   22
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "6"
      Height          =   375
      Left            =   3120
      TabIndex        =   23
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "0"
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "3"
      Height          =   375
      Left            =   3120
      TabIndex        =   20
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "4"
      Height          =   375
      Left            =   1920
      TabIndex        =   21
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      Height          =   375
      Left            =   4020
      TabIndex        =   12
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2"
      Height          =   375
      Left            =   2520
      TabIndex        =   19
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtNum 
      Height          =   195
      Left            =   3240
      TabIndex        =   64
      Top             =   960
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Trignometric Fns"
      ForeColor       =   &H00404040&
      Height          =   2055
      Left            =   120
      TabIndex        =   67
      Top             =   600
      Width           =   1575
      Begin VB.OptionButton optDeg 
         Caption         =   "&Deg"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operand Keys"
      Height          =   2655
      Left            =   1800
      TabIndex        =   68
      Top             =   600
      Width           =   1935
      Begin VB.CommandButton cmdEMinus 
         Caption         =   "E-"
         Height          =   375
         Left            =   1320
         TabIndex        =   29
         ToolTipText     =   " Exponential Powers (Minus) "
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton cmdEPlus 
         Caption         =   "E+"
         Height          =   375
         Left            =   720
         TabIndex        =   28
         ToolTipText     =   " Exponential Powers (Plus) "
         Top             =   2160
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Memory"
      Height          =   2655
      Left            =   4800
      TabIndex        =   69
      Top             =   600
      Width           =   735
   End
   Begin VB.Frame Frame4 
      Caption         =   "Special Fns"
      Height          =   2175
      Left            =   120
      TabIndex        =   70
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      Caption         =   "General"
      Height          =   2655
      Left            =   5640
      TabIndex        =   71
      Top             =   600
      Width           =   855
      Begin VB.CommandButton cmdRandom 
         Caption         =   "Rand"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Generates Random Numbers"
         Top             =   1680
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Opeartors"
      Height          =   2655
      Left            =   3840
      TabIndex        =   72
      Top             =   600
      Width           =   855
   End
   Begin VB.Frame Frame7 
      Caption         =   "Addl Memory"
      Height          =   1215
      Left            =   1800
      TabIndex        =   73
      Top             =   3360
      Width           =   4695
      Begin VB.CommandButton cmdClearAddMem 
         Caption         =   "C"
         Height          =   855
         Left            =   4320
         TabIndex        =   81
         ToolTipText     =   "Clears all Addl Memory Locations"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdMem1 
         Caption         =   "MR7"
         Height          =   375
         Index           =   13
         Left            =   3720
         TabIndex        =   59
         ToolTipText     =   "Memory Recall 7"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdMem1 
         Caption         =   "MR6"
         Height          =   375
         Index           =   12
         Left            =   3120
         TabIndex        =   57
         ToolTipText     =   "Memory Recall 6"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdMem1 
         Caption         =   "MR5"
         Height          =   375
         Index           =   11
         Left            =   2520
         TabIndex        =   55
         ToolTipText     =   "Memory Recall 5"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdMem1 
         Caption         =   "MR4"
         Height          =   375
         Index           =   10
         Left            =   1920
         TabIndex        =   53
         ToolTipText     =   "Memory Recall 4"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdMem1 
         Caption         =   "MR3"
         Height          =   375
         Index           =   9
         Left            =   1320
         TabIndex        =   51
         ToolTipText     =   "Memory Recall 3"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdMem1 
         Caption         =   "MR2"
         Height          =   375
         Index           =   8
         Left            =   720
         TabIndex        =   49
         ToolTipText     =   "Memory Recall 2"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdMem1 
         Caption         =   "MR1"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   47
         ToolTipText     =   "Memory Recall 1"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdMem1 
         Caption         =   "M7"
         Height          =   375
         Index           =   6
         Left            =   3720
         TabIndex        =   58
         ToolTipText     =   "Store Memory 7"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdMem1 
         Caption         =   "M6"
         Height          =   375
         Index           =   5
         Left            =   3120
         TabIndex        =   56
         ToolTipText     =   "Store Memory 6"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdMem1 
         Caption         =   "M5"
         Height          =   375
         Index           =   4
         Left            =   2520
         TabIndex        =   54
         ToolTipText     =   "Store Memory 5"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdMem1 
         Caption         =   "M4"
         Height          =   375
         Index           =   3
         Left            =   1920
         TabIndex        =   52
         ToolTipText     =   "Store Memory 4"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdMem1 
         Caption         =   "M3"
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   50
         ToolTipText     =   "Store Memory 3"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdMem1 
         Caption         =   "M2"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   48
         ToolTipText     =   "Store Memory 2"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdMem1 
         Caption         =   "M1"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   46
         ToolTipText     =   "Store Memory 1"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblLight 
         BackColor       =   &H0080FF80&
         Height          =   135
         Index           =   6
         Left            =   3720
         TabIndex        =   80
         Top             =   525
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblLight 
         BackColor       =   &H0080FF80&
         Height          =   135
         Index           =   5
         Left            =   3120
         TabIndex        =   79
         Top             =   525
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblLight 
         BackColor       =   &H0080FF80&
         Height          =   135
         Index           =   4
         Left            =   2520
         TabIndex        =   78
         Top             =   525
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblLight 
         BackColor       =   &H0080FF80&
         Height          =   135
         Index           =   3
         Left            =   1920
         TabIndex        =   77
         Top             =   525
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblLight 
         BackColor       =   &H0080FF80&
         Height          =   135
         Index           =   2
         Left            =   1320
         TabIndex        =   76
         Top             =   525
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblLight 
         BackColor       =   &H0080FF80&
         Height          =   135
         Index           =   1
         Left            =   720
         TabIndex        =   75
         Top             =   525
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblLight 
         BackColor       =   &H0080FF80&
         Height          =   135
         Index           =   0
         Left            =   120
         TabIndex        =   74
         Top             =   525
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Label lblMem 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   65
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblResult 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   63
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label lblOp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   62
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Name : Calculator (compact & very useful)
' Author : Sunil Wason (sunilwason@yahoo.com)
' Purpose : 'Made it during one of my first school submissions
' Its compact & definately useful
Const Pi = 3.14159265358979
Dim booPlus As Boolean
Dim booSub As Boolean
Dim booMul As Boolean
Dim booDiv As Boolean
Dim booOn As Boolean
Dim booEquals As Boolean
Dim booSqrt As Boolean
Dim booFactorial As Boolean
Dim booActiKeybd As Boolean
Dim booMem As Boolean
Dim booDot As Boolean
Dim booInvert As Boolean
Dim booCarat As Boolean
Dim booLn As Boolean
Dim booLog As Boolean
Dim booExpo As Boolean
Dim booMemPlus As Boolean
Dim booMemMinus As Boolean
Dim booAngle As Boolean
Dim sinAngle As Single
Dim dblMem As Double
Dim dblCaratVal As Double
Dim Result As Double
Dim dblMem1 As Double
Dim dblMem2 As Double
Dim dblMem3 As Double
Dim dblMem4 As Double
Dim dblMem5 As Double
Dim dblMem6 As Double
Dim dblMem7 As Double


Private Sub cmd0_Click()

If booOn Then
    InitializeStatus
    lblOp.Caption = (lblOp.Caption & "0")
End If
CheckKeyBoardStatus

End Sub

Private Sub cmd1_Click()

If booOn Then
    InitializeStatus
    If FirstCharac = 0 Then
        lblOp.Caption = Val(lblOp.Caption & "1")
    Else
        lblOp.Caption = (lblOp.Caption & "1")
    End If
End If
CheckKeyBoardStatus

End Sub

Private Sub cmd2_Click()

If booOn Then
    InitializeStatus
    If FirstCharac = 0 Then
        lblOp.Caption = Val(lblOp.Caption & "2")
    Else
        lblOp.Caption = (lblOp.Caption & "2")
    End If
End If
CheckKeyBoardStatus

End Sub

Private Sub cmd3_Click()

If booOn Then
    InitializeStatus
    If FirstCharac = 0 Then
        lblOp.Caption = Val(lblOp.Caption & "3")
    Else
        lblOp.Caption = (lblOp.Caption & "3")
    End If
End If
CheckKeyBoardStatus

End Sub

Private Sub cmd4_Click()

If booOn Then
    InitializeStatus
    If FirstCharac = 0 Then
        lblOp.Caption = Val(lblOp.Caption & "4")
    Else
        lblOp.Caption = (lblOp.Caption & "4")
    End If
End If
CheckKeyBoardStatus

End Sub

Private Sub cmd5_Click()

If booOn Then
    InitializeStatus
    If FirstCharac = 0 Then
        lblOp.Caption = Val(lblOp.Caption & "5")
    Else
        lblOp.Caption = (lblOp.Caption & "5")
    End If
End If
CheckKeyBoardStatus

End Sub

Private Sub cmd6_Click()

If booOn Then
    InitializeStatus
    If FirstCharac = 0 Then
        lblOp.Caption = Val(lblOp.Caption & "6")
    Else
        lblOp.Caption = (lblOp.Caption & "6")
    End If
End If
CheckKeyBoardStatus

End Sub

Private Sub cmd7_Click()

If booOn Then
    InitializeStatus
    If FirstCharac = 0 Then
        lblOp.Caption = Val(lblOp.Caption & "7")
    Else
        lblOp.Caption = (lblOp.Caption & "7")
    End If
End If
CheckKeyBoardStatus

End Sub

Private Sub cmd8_Click()

If booOn Then
    InitializeStatus
    If FirstCharac = 0 Then
        lblOp.Caption = Val(lblOp.Caption & "8")
    Else
        lblOp.Caption = (lblOp.Caption & "8")
    End If
End If
CheckKeyBoardStatus

End Sub

Public Sub cmd9_Click()

If booOn Then
    InitializeStatus
    If FirstCharac = 0 Then
        lblOp.Caption = Val(lblOp.Caption & "9")
    Else
        lblOp.Caption = (lblOp.Caption & "9")
    End If
End If
CheckKeyBoardStatus

End Sub

Private Function FirstCharac()

Dim LengthOfOp As Byte
LengthOfOp = Len(lblOp.Caption)
If LengthOfOp > 1 Then
    FirstCharac = Left(lblOp.Caption, 1)
End If

End Function

Private Sub InitializeStatus()

If booEquals = True Then
  lblOp.Caption = ""
  lblResult.Caption = ""
  booEquals = False
End If
If booSqrt = True Then
  lblOp.Caption = ""
  booSqrt = False
End If
If booFactorial = True Then
  lblOp.Caption = ""
  booFactorial = False
End If
If booMem = True Then
  lblOp.Caption = ""
  booMem = False
End If
If booInvert = True Then
  lblOp.Caption = ""
  booInvert = False
End If
If booLn = True Then
  lblOp.Caption = ""
  booLn = False
End If
If booExpo = True Then
  lblOp.Caption = ""
  booExpo = False
End If
If booAngle = True Then
  lblOp.Caption = ""
  booAngle = False
End If
If booMemPlus = True Then
  lblOp.Caption = ""
  booMemPlus = False
End If
If booMemMinus = True Then
  lblOp.Caption = ""
  booMemMinus = False
End If
lblResult.Visible = False
lblOp.Visible = True

End Sub



Private Sub cmdBack_Click()

Dim LengthOfOperand As Byte
Dim EPresent As Integer
LengthOfOperand = Len(lblOp.Caption)
If lblOp.Visible = True And LengthOfOperand > 0 Then
    EPresent = InStr(1, lblOp.Caption, "E")
    If EPresent > 0 Then
        DecrementPowersOnly EPresent
        CheckKeyBoardStatus
        Exit Sub
    End If
    lblOp.Caption = Left(lblOp.Caption, (LengthOfOperand - 1))
    lblOp.Caption = Val(lblOp.Caption)
End If
CheckKeyBoardStatus

End Sub

Private Sub DecrementPowersOnly(LocnOfE As Integer)

Dim strBeforeE As String
Dim strAfterE As String
Dim ValAfterE As Integer
Dim LenOfStr As Integer
Dim RightLenSought As Integer
LenOfStr = Len(lblOp.Caption)
strBeforeE = Left(lblOp.Caption, LocnOfE)
RightLenSought = LenOfStr - LocnOfE
strAfterE = Right(lblOp.Caption, RightLenSought)
On Error GoTo InvalidInput
ValAfterE = Val(strAfterE)
If ValAfterE = 1 Or ValAfterE = -1 Then
    LenOfStr = Len(strBeforeE)
    lblOp.Caption = Left(strBeforeE, (LenOfStr - 1))
    Exit Sub
End If
If ValAfterE > 0 Then
    ValAfterE = ValAfterE - 1
    lblOp.Caption = strBeforeE & "+" & ValAfterE
Else
    ValAfterE = ValAfterE + 1
    lblOp.Caption = strBeforeE & ValAfterE
End If
CheckKeyBoardStatus
Exit Sub
InvalidInput:
    MsgBox "Invalid Input for function", vbCritical + vbDefaultButton1 + vbOKOnly, "Invalid Input"
cmdClear_Click
CheckKeyBoardStatus

End Sub

Private Sub cmdBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Deletes last entry of operand"
End If

End Sub

Private Sub cmdBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Deletes last entry of operand"

End Sub

Private Sub cmdBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdCarat_Click()

If booOn Then
    If lblOp.Visible = True Then
        dblCaratVal = Val(lblOp.Caption)
        lblResult.Caption = lblOp.Caption
        lblOp.Caption = ""
        lblOp.Visible = False
        lblResult.Visible = True
    ElseIf lblResult.Visible = True Then
        dblCaratVal = Val(lblResult.Caption)
    End If
    cmdCarat.Enabled = False
    booCarat = True
End If
CheckKeyBoardStatus

End Sub

Private Sub cmdCarat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "X raised to the power of Y"

End Sub

Private Sub cmdClear_Click()

If booOn Then
    lblResult.Caption = ""
    lblOp.Caption = ""
End If
CheckKeyBoardStatus

End Sub

Private Sub cmdClear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Clears Contents"
End If

End Sub

Private Sub cmdClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Clears Contents"

End Sub

Private Sub cmdClear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub


Private Sub cmdClearAddMem_Click()

Dim i As Integer
Dim YesNo As Integer
If dblMem1 = 0 And dblMem2 = 0 And dblMem2 = 0 And dblMem3 = 0 And dblMem4 = 0 _
And dblMem5 = 0 And dblMem6 = 0 And dblMem7 = 0 Then
    MsgBox "There are no additional memory locations to be cleared." & vbCrLf & vbCrLf _
    & "If you want to clear the main memory , please click the MC button.", vbExclamation + vbDefaultButton1 + vbOKOnly, "Clear mem Locations"
    Exit Sub
End If
YesNo = MsgBox("This will clear the contents of all the " & vbCrLf _
& "additional memory locations." & vbCrLf & vbCrLf _
& "Do you want to continue ?", vbDefaultButton1 + vbQuestion + vbYesNo, "Clear all Addl memory Locations")
If YesNo = 6 Then
    dblMem1 = 0
    dblMem2 = 0
    dblMem3 = 0
    dblMem4 = 0
    dblMem5 = 0
    dblMem6 = 0
    dblMem7 = 0
    For i = 0 To 6
        lblLight(i).Visible = False
    Next i
End If

End Sub

Private Sub cmdClearAddMem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Clears All Additional Memory Locations"

End Sub

Private Sub cmdCos_Click(Index As Integer)

If booOn Then
If lblOp.Caption = "" And lblResult.Caption = "" Then
        Exit Sub
End If
Select Case Index
Case 0
    If optDeg.Value = True Then
        If lblOp.Visible = True Then
            On Error GoTo InvalidInput
            lblOp.Caption = Cos(Pi * (Val(lblOp.Caption)) / 180)
            If FindE > 0 Then
                lblOp.Caption = 0
            End If
        ElseIf lblResult.Visible = True Then
            If lblResult.Caption = "" Then
                lblResult.Caption = lblOp.Caption
            End If
            On Error GoTo InvalidInput
            lblResult.Caption = Cos(Pi * (Val(lblResult.Caption)) / 180)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        End If
    ElseIf optRad.Value = True Then
        If lblOp.Visible = True Then
            On Error GoTo InvalidInput
            lblOp.Caption = Cos(Val(lblOp.Caption))
            If FindE > 0 Then
                lblOp.Caption = 0
            End If
        ElseIf lblResult.Visible = True Then
            If lblResult.Caption = "" Then
                lblResult.Caption = lblOp.Caption
            End If
            On Error GoTo InvalidInput
            lblResult.Caption = Cos(Val(lblResult.Caption))
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        End If
    End If
Case 1
    If optDeg.Value = True Then
        If lblOp.Visible = True Then
            On Error GoTo InvalidInput
            lblOp.Caption = Sin(Pi * (Val(lblOp.Caption)) / 180)
            If FindE > 0 Then
                lblOp.Caption = 0
            End If
        ElseIf lblResult.Visible = True Then
            If lblResult.Caption = "" Then
                lblResult.Caption = lblOp.Caption
            End If
            On Error GoTo InvalidInput
            lblResult.Caption = Sin(Pi * (Val(lblResult.Caption)) / 180)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        End If
    ElseIf optRad.Value = True Then
        If lblOp.Visible = True Then
            On Error GoTo InvalidInput
            lblOp.Caption = Sin(Val(lblOp.Caption))
            If FindE > 0 Then
                lblOp.Caption = 0
            End If
        ElseIf lblResult.Visible = True Then
            If lblResult.Caption = "" Then
                lblResult.Caption = lblOp.Caption
            End If
            On Error GoTo InvalidInput
            lblResult.Caption = Sin(Val(lblResult.Caption))
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        End If
    End If
Case 2
    If optDeg.Value = True Then
        If lblOp.Visible = True Then
            On Error GoTo InvalidInput
            If Val(lblOp.Caption) = 90 Then GoTo InvalidInput
            lblOp.Caption = Tan(Pi * (Val(lblOp.Caption)) / 180)
            If FindE > 0 Then
                lblOp.Caption = 0
            End If
        ElseIf lblResult.Visible = True Then
            If lblResult.Caption = "" Then
                lblResult.Caption = lblOp.Caption
            End If
            On Error GoTo InvalidInput
            If Val(lblResult.Caption) = 90 Then GoTo InvalidInput
            lblResult.Caption = Tan(Pi * (Val(lblResult.Caption)) / 180)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        End If
    ElseIf optRad.Value = True Then
        If lblOp.Visible = True Then
            On Error GoTo InvalidInput
            If Val(lblOp.Caption) >= 1.57 And Val(lblOp.Caption) <= 1.58 Then GoTo InvalidInput
            lblOp.Caption = Tan(Val(lblOp.Caption))
            If FindE > 0 Then
                lblOp.Caption = 0
            End If
        ElseIf lblResult.Visible = True Then
            If lblResult.Caption = "" Then
                lblResult.Caption = lblOp.Caption
            End If
            On Error GoTo InvalidInput
            If Val(lblResult.Caption) >= 1.57 And Val(lblResult.Caption) <= 1.58 Then GoTo InvalidInput
            lblResult.Caption = Tan(Val(lblResult.Caption))
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        End If
    End If
Case 3
    If optDeg.Value = True Then
        If lblOp.Visible = True Then
            On Error GoTo InvalidInput
            lblOp.Caption = Sin(Pi * (Val(lblOp.Caption)) / 180)
            If FindE > 0 Then
                lblOp.Caption = 0
            End If
            lblOp.Caption = 1 / Val(lblOp.Caption)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        ElseIf lblResult.Visible = True Then
            If lblResult.Caption = "" Then
                lblResult.Caption = lblOp.Caption
            End If
            On Error GoTo InvalidInput
            lblResult.Caption = Sin(Pi * (Val(lblResult.Caption)) / 180)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
            lblResult.Caption = 1 / Val(lblResult.Caption)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        End If
    ElseIf optRad.Value = True Then
        If lblOp.Visible = True Then
            On Error GoTo InvalidInput
            lblOp.Caption = Sin(Val(lblOp.Caption))
            If FindE > 0 Then
                lblOp.Caption = 0
            End If
            lblOp.Caption = 1 / Val(lblOp.Caption)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        ElseIf lblResult.Visible = True Then
            If lblResult.Caption = "" Then
                lblResult.Caption = lblOp.Caption
            End If
            On Error GoTo InvalidInput
            lblResult.Caption = Sin(Val(lblResult.Caption))
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
            lblResult.Caption = 1 / Val(lblResult.Caption)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        End If
    End If
Case 4
    If optDeg.Value = True Then
        If lblOp.Visible = True Then
            On Error GoTo InvalidInput
            lblOp.Caption = Cos(Pi * (Val(lblOp.Caption)) / 180)
            If FindE > 0 Then
                lblOp.Caption = 0
            End If
            lblOp.Caption = 1 / Val(lblOp.Caption)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        ElseIf lblResult.Visible = True Then
            If lblResult.Caption = "" Then
                lblResult.Caption = lblOp.Caption
            End If
            On Error GoTo InvalidInput
            lblResult.Caption = Cos(Pi * (Val(lblResult.Caption)) / 180)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
            lblResult.Caption = 1 / Val(lblResult.Caption)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        End If
    ElseIf optRad.Value = True Then
        If lblOp.Visible = True Then
            On Error GoTo InvalidInput
            lblOp.Caption = Cos(Val(lblOp.Caption))
            If FindE > 0 Then
                lblOp.Caption = 0
            End If
            lblOp.Caption = 1 / Val(lblOp.Caption)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        ElseIf lblResult.Visible = True Then
            If lblResult.Caption = "" Then
                lblResult.Caption = lblOp.Caption
            End If
            On Error GoTo InvalidInput
            lblResult.Caption = Cos(Val(lblResult.Caption))
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
            lblResult.Caption = 1 / Val(lblResult.Caption)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        End If
    End If
Case 5
    If optDeg.Value = True Then
        If lblOp.Visible = True Then
            On Error GoTo InvalidInput
            lblOp.Caption = Tan(Pi * (Val(lblOp.Caption)) / 180)
            If FindE > 0 Then
                lblOp.Caption = 0
            End If
            lblOp.Caption = 1 / Val(lblOp.Caption)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        ElseIf lblResult.Visible = True Then
            If lblResult.Caption = "" Then
                lblResult.Caption = lblOp.Caption
            End If
            On Error GoTo InvalidInput
            lblResult.Caption = Tan(Pi * (Val(lblResult.Caption)) / 180)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
            lblResult.Caption = 1 / Val(lblResult.Caption)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        End If
    ElseIf optRad.Value = True Then
        If lblOp.Visible = True Then
            On Error GoTo InvalidInput
            lblOp.Caption = Tan(Val(lblOp.Caption))
            If FindE > 0 Then
                lblOp.Caption = 0
            End If
            lblOp.Caption = 1 / Val(lblOp.Caption)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        ElseIf lblResult.Visible = True Then
            If lblResult.Caption = "" Then
                lblResult.Caption = lblOp.Caption
            End If
            On Error GoTo InvalidInput
            lblResult.Caption = Tan(Val(lblResult.Caption))
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
            lblResult.Caption = 1 / Val(lblResult.Caption)
            If FindE > 0 Then
                lblResult.Caption = 0
            End If
        End If
    End If
End Select
End If
CheckKeyBoardStatus
booAngle = True
Exit Sub
InvalidInput:
    MsgBox "Invalid Input for function", vbCritical + vbDefaultButton1 + vbOKOnly, "Invalid Input"
cmdClear_Click
CheckKeyBoardStatus

End Sub

Private Function FindE() As Integer

Dim NegSign As Integer
Dim LenOfStr As Integer
Dim PosnOfSign As Integer
Dim PosnOfSignFromRight As Integer
Dim ValOfNegSign As Integer
If lblOp.Visible = True Then
    If Val(lblOp.Caption) < 0 Then
        FindE = InStr(1, lblOp.Caption, "E")
    Else
        FindE = InStr(1, lblOp.Caption, "E")
        NegSign = InStr(1, lblOp.Caption, "-")
        If NegSign = 0 Then
            FindE = 0
        End If
        If NegSign > 0 Then
            LenOfStr = Len(lblOp.Caption)
            PosnOfSign = InStr(1, lblOp.Caption, "-")
            PosnOfSignFromRight = LenOfStr - PosnOfSign
            ValOfNegSign = Val(Right(lblOp.Caption, (PosnOfSignFromRight)))
            If ValOfNegSign <= 10 Then
                FindE = 0
            End If
         End If
    End If
ElseIf lblResult.Visible = True Then
    If Val(lblResult.Caption) < 0 Then
        FindE = InStr(1, lblResult.Caption, "E")
    Else
        FindE = InStr(1, lblResult.Caption, "E")
        NegSign = InStr(1, lblResult.Caption, "-")
        If NegSign > 0 Then
            LenOfStr = Len(lblResult.Caption)
            PosnOfSign = InStr(1, lblResult.Caption, "-")
            PosnOfSignFromRight = LenOfStr - PosnOfSign
            ValOfNegSign = Val(Right(lblResult.Caption, (PosnOfSignFromRight)))
            If ValOfNegSign <= 10 Then
                FindE = 0
            End If
         End If
    End If
End If

End Function

Private Sub cmdCos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Choose the Deg or Rad option to enter the Number in Degrees or Radians"

End Sub

Private Sub cmdDecimal_Click()

Dim dblConvertResult As Double
Dim LocnOfDot As Integer
InitializeStatus
LocnOfDot = InStr(1, lblOp.Caption, ".")
If LocnOfDot > 0 Then
    booDot = True
Else
    booDot = False
End If
If booDot = False Then
    If lblOp.Visible = True Then
        lblOp.Caption = lblOp.Caption & "."
        booDot = True
    End If
End If
CheckKeyBoardStatus

End Sub

Private Sub cmdDiv_Click()

If booOn Then
    CheckEqualsStatus
    CheckFactorialStatusMultDiv
    If booSub = True Then
        cmdSub_Click
        booSub = False
        lblOp.Caption = "1"
    End If
    If booMul = True Then
        cmdMul_Click
        booMul = False
        lblOp.Caption = "1"
    End If
    If booPlus = True Then
        cmdPlus_Click
        booPlus = False
        lblOp.Caption = "1"
    End If
    If lblResult.Caption = "" Then
        lblResult.Caption = Val(lblOp.Caption)
        lblOp.Caption = "1"
    End If
    If lblOp.Caption = "" Then
        lblOp.Caption = "1"
    End If
    If Val(lblOp.Caption) = 0 Then
        MsgBox "Division by Zero !", vbExclamation + vbDefaultButton1 + vbOKOnly, "Error during division"
        Exit Sub
    End If
    On Error GoTo InvalidInput
    lblResult.Caption = Val(lblResult.Caption) / Val(lblOp.Caption)
    lblOp.Caption = ""
    lblOp.Visible = False
    lblResult.Visible = True
End If
IniAllBooVar
booDot = False
booDiv = True
CheckKeyBoardStatus
Exit Sub
InvalidInput:
    MsgBox "Invalid Input for function", vbCritical + vbDefaultButton1 + vbOKOnly, "Invalid Input"
cmdClear_Click
CheckKeyBoardStatus

End Sub

Private Sub cmdDiv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Division"
End If

End Sub

Private Sub cmdDiv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Division"

End Sub

Private Sub cmdDiv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdEMinus_Click()

If lblOp.Visible = True Then
    lblOp.Caption = lblOp.Caption & "E-"
End If
CheckKeyBoardStatus

End Sub

Private Sub cmdEMinus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Exponential Powers"

End Sub

Private Sub cmdEPlus_Click()

If lblOp.Visible = True Then
    lblOp.Caption = lblOp.Caption & "E+"
End If
CheckKeyBoardStatus

End Sub

Private Sub cmdEPlus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Exponential Powers"

End Sub

Private Sub cmdEquals_Click()

If booPlus = True Then
    lblOp.Visible = False
    lblResult.Visible = True
    lblResult.Caption = Val(lblResult.Caption) + Val(lblOp.Caption)
ElseIf booSub = True Then
    If lblResult.Caption = "" Then
        lblResult.Caption = Val(lblOp.Caption)
        lblOp.Caption = ""
    End If
    lblResult.Caption = Val(lblResult.Caption) - Val(lblOp.Caption)
    lblOp.Caption = ""
    lblOp.Visible = False
    lblResult.Visible = True
ElseIf booMul = True Then
    If lblResult.Caption = "" Then
        lblResult.Caption = "1"
    End If
    lblResult.Caption = Val(lblResult.Caption) * Val(lblOp.Caption)
    lblOp.Caption = ""
    lblOp.Visible = False
    lblResult.Visible = True
ElseIf booDiv = True Then
    If lblResult.Caption = "" Then
        lblResult.Caption = Val(lblOp.Caption)
        lblOp.Caption = "1"
    End If
    If Val(lblOp.Caption) = 0 Then
        MsgBox "Division by Zero !", vbExclamation + vbDefaultButton1 + vbOKOnly, "Error during division"
        cmdClear_Click
        Exit Sub
    End If
    lblResult.Caption = Val(lblResult.Caption) / Val(lblOp.Caption)
    lblOp.Caption = ""
    lblOp.Visible = False
    lblResult.Visible = True
ElseIf booCarat = True Then
    cmdCarat.Enabled = True
    lblOp.Visible = False
    lblResult.Visible = True
    On Error GoTo Message
    lblResult.Caption = dblCaratVal ^ Val(lblOp.Caption)
    booCarat = False
End If
IniAllBooVar
booEquals = True
CheckKeyBoardStatus
Exit Sub
Message:
    MsgBox "The Number is too large to be handled even by a Double Variable." _
    & vbCrLf & vbCrLf & "The Calculator will now be reset to its default ", vbCritical + vbDefaultButton1 + vbOKOnly, "Overflow"
    cmdClear_Click
    On Error GoTo 0
End Sub

Private Sub cmdEx_Click()
Const expo = 2.718282
If booOn Then
    If lblOp.Caption = "" And lblResult.Caption = "" Then
        Exit Sub
    End If
    If lblOp.Visible = True Then
        On Error GoTo InvalidInput
        lblOp.Caption = expo ^ (Val(lblOp.Caption))
    ElseIf lblResult.Visible = True Then
        If lblResult.Caption = "" Then
            lblResult.Caption = lblOp.Caption
        End If
        On Error GoTo InvalidInput
        lblResult.Caption = expo ^ (Val(lblResult.Caption))
    End If
End If
booExpo = True
CheckKeyBoardStatus
Exit Sub
InvalidInput:
    MsgBox "Invalid Input for function", vbCritical + vbDefaultButton1 + vbOKOnly, "Invalid Input"
cmdClear_Click
CheckKeyBoardStatus

End Sub

Private Sub cmdExit_Click()

End

End Sub

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Close Calculator"
End If

End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Close Calculator"

End Sub

Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdFactorial_Click()

Dim dblFactorial As Double
Dim dblCounter As Double
If booOn Then
    If lblOp.Visible = True Then
        dblFactorial = Round(Val(lblOp.Caption))
        lblOp.Caption = dblFactorial
    ElseIf lblResult.Visible = True Then
        dblFactorial = Round(Val(lblResult.Caption))
        lblResult.Caption = dblFactorial
    End If
    For dblCounter = (dblFactorial - 1) To 1 Step -1
        On Error GoTo Message
        dblFactorial = dblFactorial * dblCounter
    Next dblCounter
    lblOp.Visible = False
    lblResult.Visible = True
    lblResult.Caption = dblFactorial
    booFactorial = True
    CheckKeyBoardStatus
    Exit Sub
Message:
    MsgBox "The Number is too large to be handled even by a Double Variable." _
    & vbCrLf & vbCrLf & "The Calculator will now be reset to its default ", vbCritical + vbDefaultButton1 + vbOKOnly, "Overflow"
    cmdClear_Click
    On Error GoTo 0
End If
CheckKeyBoardStatus

End Sub

Private Sub cmdFactorial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Factorial"
End If

End Sub

Private Sub cmdFactorial_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Factorial"

End Sub

Private Sub cmdFactorial_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdInvert_Click()

If booOn Then
    lblOp.Visible = False
    lblResult.Visible = True
    If lblResult.Caption = "" Then
        lblResult.Caption = lblOp.Caption
    End If
    If Val(lblResult.Caption) = 0 Then
        MsgBox "Division by Zero !", vbExclamation + vbDefaultButton1 + vbOKOnly, "Error during division"
        lblResult.Caption = ""
        Exit Sub
    End If
    lblResult.Caption = 1 / Val(lblResult.Caption)
End If
booInvert = True
CheckKeyBoardStatus

End Sub

Private Sub cmdInvert_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Inversion"
End If

End Sub

Private Sub cmdInvert_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Inversion"

End Sub

Private Sub cmdInvert_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdKeyBoard_Click()

If booActiKeybd = False Then
    txtNum.SetFocus
    cmdKeyBoard.Caption = "De-Activate &Keyboard"
    cmdKeyBoard.ToolTipText = " Deactivates the Keyboard "
    booActiKeybd = True
Else
    cmdKeyBoard.Caption = "Activate &Keyboard"
    cmdKeyBoard.ToolTipText = "Allows the user to enter the Numbers from the Keyboard "
    booActiKeybd = False
End If

End Sub

Private Sub cmdKeyBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 And booActiKeybd = True Then
    SB.SimpleText = "De-Activates the Keyboard"
ElseIf Button = 2 And booActiKeybd = False Then
    SB.SimpleText = "Activates the Keyboard for entering Operands"
End If

End Sub

Private Sub cmdKeyBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If booActiKeybd = True Then
    SB.SimpleText = "De-Activates the Keyboard"
ElseIf booActiKeybd = False Then
    SB.SimpleText = "Activates the Keyboard for entering Operands"
End If

End Sub

Private Sub cmdKeyBoard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdLn_Click()

If booOn Then
    If lblOp.Caption = "" And lblResult.Caption = "" Then
        Exit Sub
    End If
    If lblOp.Visible = True Then
        On Error GoTo InvalidInput
        lblOp.Caption = Log(Val(lblOp.Caption))
    ElseIf lblResult.Visible = True Then
        If lblResult.Caption = "" Then
            lblResult.Caption = lblOp.Caption
        End If
        On Error GoTo InvalidInput
        lblResult.Caption = Log(Val(lblResult.Caption))
    End If
End If
booLn = True
CheckKeyBoardStatus
Exit Sub
InvalidInput:
    MsgBox "Invalid Input for function", vbCritical + vbDefaultButton1 + vbOKOnly, "Invalid Input"
cmdClear_Click
CheckKeyBoardStatus

End Sub

Private Sub cmdLn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Natural Logarithm"
End If

End Sub

Private Sub cmdLn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Natural Logarithm"

End Sub

Private Sub cmdLn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdLog_Click()

If booOn Then
    If lblOp.Caption = "" And lblResult.Caption = "" Then
        Exit Sub
    End If
    If lblOp.Visible = True Then
        On Error GoTo InvalidInput
        lblOp.Caption = Log10(Val(lblOp.Caption))
    ElseIf lblResult.Visible = True Then
        If lblResult.Caption = "" Then
            lblResult.Caption = lblOp.Caption
        End If
        On Error GoTo InvalidInput
        lblResult.Caption = Log10(Val(lblResult.Caption))
    End If
End If
booLn = True
CheckKeyBoardStatus
Exit Sub
InvalidInput:
    MsgBox "Invalid Input for function", vbCritical + vbDefaultButton1 + vbOKOnly, "Invalid Input"
cmdClear_Click
CheckKeyBoardStatus

End Sub

Function Log10(X As Double) As Double

Log10 = Log(X) / Log(10)

End Function

Private Sub cmdLog_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Logarithm to the Base 10"
End If

End Sub

Private Sub cmdLog_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Logarithm to the Base 10"

End Sub

Private Sub cmdLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdMC_Click()

dblMem = 0
lblMem.Caption = ""
CheckKeyBoardStatus

End Sub

Private Sub cmdMC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Memory Clear"
End If

End Sub

Private Sub cmdMC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Memory Clear"

End Sub

Private Sub cmdMC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdMem1_Click(Index As Integer)

If booOn Then
Select Case Index
    Case 0
        If lblOp.Visible = True Then
            dblMem1 = Val(lblOp.Caption)
        ElseIf lblResult.Visible = True Then
            dblMem1 = Val(lblResult.Caption)
        End If
        lblLight(Index).Visible = True
    Case 1
        If lblOp.Visible = True Then
            dblMem2 = Val(lblOp.Caption)
        ElseIf lblResult.Visible = True Then
            dblMem2 = Val(lblResult.Caption)
        End If
        lblLight(Index).Visible = True
    Case 2
        If lblOp.Visible = True Then
            dblMem3 = Val(lblOp.Caption)
        ElseIf lblResult.Visible = True Then
            dblMem3 = Val(lblResult.Caption)
        End If
        lblLight(Index).Visible = True
    Case 3
        If lblOp.Visible = True Then
            dblMem4 = Val(lblOp.Caption)
        ElseIf lblResult.Visible = True Then
            dblMem4 = Val(lblResult.Caption)
        End If
        lblLight(Index).Visible = True
    Case 4
        If lblOp.Visible = True Then
            dblMem5 = Val(lblOp.Caption)
        ElseIf lblResult.Visible = True Then
            dblMem5 = Val(lblResult.Caption)
        End If
        lblLight(Index).Visible = True
    Case 5
        If lblOp.Visible = True Then
            dblMem6 = Val(lblOp.Caption)
        ElseIf lblResult.Visible = True Then
            dblMem6 = Val(lblResult.Caption)
        End If
        lblLight(Index).Visible = True
    Case 6
        If lblOp.Visible = True Then
            dblMem7 = Val(lblOp.Caption)
        ElseIf lblResult.Visible = True Then
            dblMem7 = Val(lblResult.Caption)
        End If
        lblLight(Index).Visible = True
    Case 7
        If booPlus = False And booSub = False And booMul = False _
            And booDiv = False Then
            If lblOp.Visible = True Then
                lblOp.Caption = dblMem1
            ElseIf lblResult.Visible = True Then
                lblResult.Caption = dblMem1
            End If
        Else
            If lblOp.Visible = True Then
                lblOp.Visible = False
                lblResult.Visible = True
                lblResult.Caption = dblMem1
            Else
                lblResult.Visible = False
                lblOp.Visible = True
                lblOp.Caption = dblMem1
            End If
        End If
    Case 8
        If booPlus = False And booSub = False And booMul = False _
            And booDiv = False Then
            If lblOp.Visible = True Then
                lblOp.Caption = dblMem2
            ElseIf lblResult.Visible = True Then
                lblResult.Caption = dblMem2
            End If
        Else
            If lblOp.Visible = True Then
                lblOp.Visible = False
                lblResult.Visible = True
                lblResult.Caption = dblMem2
            Else
                lblResult.Visible = False
                lblOp.Visible = True
                lblOp.Caption = dblMem2
            End If
        End If
    Case 9
        If booPlus = False And booSub = False And booMul = False _
            And booDiv = False Then
            If lblOp.Visible = True Then
                lblOp.Caption = dblMem3
            ElseIf lblResult.Visible = True Then
                lblResult.Caption = dblMem3
            End If
        Else
            If lblOp.Visible = True Then
                lblOp.Visible = False
                lblResult.Visible = True
                lblResult.Caption = dblMem3
            Else
                lblResult.Visible = False
                lblOp.Visible = True
                lblOp.Caption = dblMem3
            End If
        End If
    Case 10
        If booPlus = False And booSub = False And booMul = False _
            And booDiv = False Then
            If lblOp.Visible = True Then
                lblOp.Caption = dblMem4
            ElseIf lblResult.Visible = True Then
                lblResult.Caption = dblMem4
            End If
        Else
            If lblOp.Visible = True Then
                lblOp.Visible = False
                lblResult.Visible = True
                lblResult.Caption = dblMem4
            Else
                lblResult.Visible = False
                lblOp.Visible = True
                lblOp.Caption = dblMem4
            End If
        End If
    Case 11
        If booPlus = False And booSub = False And booMul = False _
            And booDiv = False Then
            If lblOp.Visible = True Then
                lblOp.Caption = dblMem5
            ElseIf lblResult.Visible = True Then
                lblResult.Caption = dblMem5
            End If
        Else
            If lblOp.Visible = True Then
                lblOp.Visible = False
                lblResult.Visible = True
                lblResult.Caption = dblMem5
            Else
                lblResult.Visible = False
                lblOp.Visible = True
                lblOp.Caption = dblMem5
            End If
        End If
    Case 12
        If booPlus = False And booSub = False And booMul = False _
            And booDiv = False Then
            If lblOp.Visible = True Then
                lblOp.Caption = dblMem6
            ElseIf lblResult.Visible = True Then
                lblResult.Caption = dblMem6
            End If
        Else
            If lblOp.Visible = True Then
                lblOp.Visible = False
                lblResult.Visible = True
                lblResult.Caption = dblMem6
            Else
                lblResult.Visible = False
                lblOp.Visible = True
                lblOp.Caption = dblMem6
            End If
        End If
    Case 13
        If booPlus = False And booSub = False And booMul = False _
            And booDiv = False Then
            If lblOp.Visible = True Then
                lblOp.Caption = dblMem7
            ElseIf lblResult.Visible = True Then
                lblResult.Caption = dblMem7
            End If
        Else
            If lblOp.Visible = True Then
                lblOp.Visible = False
                lblResult.Visible = True
                lblResult.Caption = dblMem7
            Else
                lblResult.Visible = False
                lblOp.Visible = True
                lblOp.Caption = dblMem7
            End If
        End If
End Select
End If
CheckKeyBoardStatus

End Sub

Private Sub cmdMem1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "M1 to M7 Stores Data in Memories and MR1 to MR7 Recalls the data from the Memories"

End Sub

Private Sub cmdMemPlus_Click()

If lblOp.Visible = True Then
    dblMem = dblMem + Val(lblOp.Caption)
ElseIf lblResult.Visible = True Then
    dblMem = dblMem + Val(lblResult.Caption)
End If
booMemPlus = True
CheckKeyBoardStatus

End Sub

Private Sub cmdMemPlus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Add operand or result to Memory"
End If

End Sub

Private Sub cmdMemPlus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Add operand or result to Memory"

End Sub

Private Sub cmdMemPlus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdMMinus_Click()

If lblOp.Visible = True Then
    dblMem = dblMem - Val(lblOp.Caption)
ElseIf lblResult.Visible = True Then
    dblMem = dblMem - Val(lblResult.Caption)
End If
booMemMinus = True
CheckKeyBoardStatus

End Sub

Private Sub cmdMMinus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Subtract operand or result from Memory"
End If

End Sub

Private Sub cmdMMinus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Subtract operand or result from Memory"

End Sub

Private Sub cmdMMinus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdMplus_Click()

If booOn Then
    If lblOp.Visible = True Then
        dblMem = Val(lblOp.Caption)
        lblMem.Caption = "M"
    ElseIf lblResult.Visible = True Then
        dblMem = Val(lblResult.Caption)
        lblMem.Caption = "M"
    End If
    booMem = True
End If
CheckKeyBoardStatus

End Sub

Private Sub cmdMplus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Saves the dispayed number in memory"
End If

End Sub

Private Sub cmdMplus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Saves the dispayed number in memory"

End Sub

Private Sub cmdMplus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdMR_Click()

If booPlus = False And booSub = False And booMul = False _
And booDiv = False Then
  If lblOp.Visible = True Then
     lblOp.Caption = dblMem
  ElseIf lblResult.Visible = True Then
     lblResult.Caption = dblMem
  End If
Else
  If lblOp.Visible = True Then
     lblOp.Visible = False
     lblResult.Visible = True
     lblResult.Caption = dblMem
  Else
     lblResult.Visible = False
     lblOp.Visible = True
     lblOp.Caption = dblMem
  End If
End If
CheckKeyBoardStatus

End Sub

Private Sub cmdMR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Memory Recall"
End If

End Sub

Private Sub cmdMR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Memory Recall"

End Sub

Private Sub cmdMR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdMul_Click()

If booOn Then
    CheckEqualsStatus
    CheckFactorialStatusMultDiv
    If booSub = True Then
        cmdSub_Click
        booSub = False
        lblOp.Caption = "1"
    End If
    If booDiv = True Then
        cmdDiv_Click
        booDiv = False
        lblOp.Caption = "1"
    End If
    If booPlus = True Then
        cmdPlus_Click
        booPlus = False
        lblOp.Caption = "1"
    End If
    If lblResult.Caption = "" Then
       On Error GoTo InvalidInput
       lblResult.Caption = Val(lblOp.Caption)
       lblOp.Caption = "1"
    End If
    If lblResult.Caption = "" Then
       lblResult.Caption = "1"
    End If
    If lblOp.Caption = "" Then
       lblOp.Caption = "1"
    End If
    On Error GoTo InvalidInput
    lblResult.Caption = Val(lblResult.Caption) * Val(lblOp.Caption)
    lblOp.Caption = ""
    lblOp.Visible = False
    lblResult.Visible = True
End If
IniAllBooVar
booMul = True
booDot = False
CheckKeyBoardStatus
Exit Sub
InvalidInput:
    MsgBox "Invalid Input for function", vbCritical + vbDefaultButton1 + vbOKOnly, "Invalid Input"
cmdClear_Click
CheckKeyBoardStatus

End Sub

Private Sub CheckFactorialStatusMultDiv()

If booFactorial = True Then
  lblOp.Caption = "1"
  booFactorial = False
End If

End Sub
Private Sub CheckFactorialStatusAddSub()

If booFactorial = True Then
  lblOp.Caption = "0"
  booFactorial = False
End If

End Sub

Private Sub cmdMul_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Multiplication"
End If

End Sub

Private Sub cmdMul_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Multiplication"

End Sub

Private Sub cmdMul_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdOff_Click()

lblOp.Caption = ""
lblResult.Alignment = 2
lblResult.Caption = "Calculator is Switched Off"
booOn = False
booEquals = False
lblOp.BackColor = &H8000000F
lblResult.BackColor = &H8000000F
cmdKeyBoard.Caption = "Activate &Keyboard"
booActiKeybd = False

End Sub

Private Sub cmdOff_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Switches Calcualtor Off"
End If

End Sub

Private Sub cmdOff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Switches Calcualtor Off"

End Sub

Private Sub cmdOff_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdOn_Click()

booOn = True
cmdKeyBoard.Enabled = True
optDeg.Enabled = True
optRad.Enabled = True
lblOp.Caption = ""
lblResult.Alignment = 1
lblOp.Visible = False
lblResult.Visible = True
lblResult.Caption = ""
lblOp.BackColor = &HFFFFC0
lblResult.BackColor = &HFFFFC0

End Sub

Private Sub cmdOn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Switches Calcualtor On"
End If

End Sub

Private Sub cmdOn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Switches Calcualtor On"

End Sub

Private Sub cmdOn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub ClearSB(Button As Integer)

If Button = 2 Then
    SB.SimpleText = ""
End If

End Sub

Private Sub cmdPi_Click()

If booOn Then
    InitializeStatus
    lblOp.Caption = Pi
End If
Exit Sub
CheckKeyBoardStatus

End Sub

Private Sub cmdPi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Value of Pi"

End Sub

Private Sub cmdPlus_Click()

If booOn Then
    CheckEqualsStatus
    CheckFactorialStatusAddSub
    If booMul = True Then
        cmdMul_Click
        booMul = False
        lblOp.Caption = "0"
    End If
    If booDiv = True Then
        cmdDiv_Click
        booDiv = False
        lblOp.Caption = "0"
    End If
    If booSub = True Then
        cmdSub_Click
        booSub = False
        lblOp.Caption = "0"
    End If
    On Error GoTo InvalidInput
    lblResult.Caption = Val(lblResult.Caption) + Val(lblOp.Caption)
    lblOp.Caption = ""
    lblOp.Visible = False
    lblResult.Visible = True
End If
IniAllBooVar
booPlus = True
booDot = False
CheckKeyBoardStatus
Exit Sub
InvalidInput:
    MsgBox "Invalid Input for function", vbCritical + vbDefaultButton1 + vbOKOnly, "Invalid Input"
cmdClear_Click
CheckKeyBoardStatus

End Sub

Private Sub CheckEqualsStatus()

If booEquals = True Then
    lblOp.Caption = ""
    booEquals = False
End If

End Sub
Private Sub IniAllBooVar()

booPlus = False
booSub = False
booMul = False
booDiv = False

End Sub

Private Sub cmdPlus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Addition"
End If

End Sub

Private Sub cmdPlus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Addition"

End Sub

Private Sub cmdPlus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdPlusMinus_Click()

Dim dblConvertResult As Double
If lblOp.Visible = True Then
    On Error GoTo InvalidInput
    If Val(lblOp.Caption) < 0 Then
        dblConvertResult = Val(lblOp.Caption) - (2 * Val(lblOp.Caption))
        lblOp.Caption = dblConvertResult
    Else
        lblOp.Caption = "-" & lblOp.Caption
        lblOp.Caption = Val(lblOp.Caption)
    End If
ElseIf lblResult.Visible = True Then
    If Val(lblResult.Caption) < 0 Then
        dblConvertResult = Val(lblResult.Caption) - (2 * Val(lblResult.Caption))
        lblResult.Caption = dblConvertResult
    Else
        lblResult.Caption = "-" & lblResult.Caption
        lblResult.Caption = Val(lblResult.Caption)
    End If
End If
CheckKeyBoardStatus
Exit Sub
InvalidInput:
    MsgBox "Invalid Input for function", vbCritical + vbDefaultButton1 + vbOKOnly, "Invalid Input"
cmdClear_Click
CheckKeyBoardStatus

End Sub

Private Sub CheckKeyBoardStatus()

If booActiKeybd = True Then
     txtNum.SetFocus
End If

End Sub

Private Sub cmdPlusMinus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Change sign"
End If

End Sub

Private Sub cmdPlusMinus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Change sign"

End Sub

Private Sub cmdPlusMinus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdRandom_Click()
    
InitializeStatus
lblOp.Caption = ((1 * Rnd) + 0)
CheckKeyBoardStatus

End Sub

Private Sub cmdRandom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Generates Random Numbers between 0 and 1"

End Sub

Private Sub cmdRnd_Click()

If lblOp.Visible = True Then
    lblOp.Caption = Round(Val(lblOp.Caption))
ElseIf lblResult.Visible = True Then
    lblResult.Caption = Round(Val(lblResult.Caption))
End If
CheckKeyBoardStatus

End Sub

Private Sub cmdRnd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Rounds off to the nearest value"
End If

End Sub

Private Sub cmdRnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Rounds off to the nearest value"

End Sub

Private Sub cmdRnd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdSqrt_Click()

If booOn Then
    If lblOp.Caption = "" And lblResult.Caption = "" Then
        Exit Sub
    End If
    If lblOp.Visible = True Then
        lblOp.Caption = Sqr(Val(lblOp.Caption))
    ElseIf lblResult.Visible = True Then
        If lblResult.Caption = "" Then
            lblResult.Caption = lblOp.Caption
        End If
        lblResult.Caption = Sqr(Val(lblResult.Caption))
    End If
End If
booSqrt = True
CheckKeyBoardStatus

End Sub

Private Sub cmdSqrt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Square root"
End If

End Sub

Private Sub cmdSqrt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Square root"

End Sub

Private Sub cmdSqrt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub cmdSub_Click()

If booOn Then
    CheckEqualsStatus
    CheckFactorialStatusAddSub
    If booMul = True Then
        cmdMul_Click
        booMul = False
        lblOp.Caption = "0"
    End If
    If booDiv = True Then
        cmdDiv_Click
        booDiv = False
        lblOp.Caption = "0"
    End If
    If booPlus = True Then
        cmdPlus_Click
        booPlus = False
        lblOp.Caption = "0"
    End If
    If lblResult.Caption = "" Then
        On Error GoTo InvalidInput
        lblResult.Caption = Val(lblOp.Caption)
        lblOp.Caption = ""
    End If
    On Error GoTo InvalidInput
    lblResult.Caption = Val(lblResult.Caption) - Val(lblOp.Caption)
    lblOp.Caption = ""
    lblOp.Visible = False
    lblResult.Visible = True
End If
IniAllBooVar
booSub = True
booDot = False
CheckKeyBoardStatus
Exit Sub
InvalidInput:
    MsgBox "Invalid Input for function", vbCritical + vbDefaultButton1 + vbOKOnly, "Invalid Input"
cmdClear_Click
CheckKeyBoardStatus

End Sub

Private Sub cmdSub_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    SB.SimpleText = "Subtraction"
End If

End Sub

Private Sub cmdSub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Subtraction"
 
End Sub

Private Sub cmdSub_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

ClearSB Button

End Sub

Private Sub Form_Activate()

txtNum.SetFocus

End Sub

Private Sub Form_Load()

lblOp.Caption = ""
lblResult.Alignment = 2
lblResult.Caption = "Please Click the On button to Start"
cmdKeyBoard.Enabled = False
booOn = False
booEquals = False
booActiKeybd = False
optDeg.Value = True
optRad.Value = False
optDeg.Enabled = False
optRad.Enabled = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = ""

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = ""

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = ""

End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = ""

End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = ""

End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = ""

End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = ""

End Sub

Private Sub Frame7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = ""

End Sub

Private Sub lblOp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If booOn And lblOp.Caption <> "" Then
    SB.SimpleText = "The Operand is being displayed"
Else
    SB.SimpleText = "Displays the Operand or the Result"
End If

End Sub

Private Sub lblResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If booOn And lblResult.Caption <> "" Then
    SB.SimpleText = "The Result of the operations is being displayed"
Else
    SB.SimpleText = "Displays the Operand or the Result"
End If

End Sub

Private Sub optDeg_Click()

optDeg.Value = True
optRad.Value = False
CheckKeyBoardStatus

End Sub

Private Sub optDeg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Degrees"

End Sub

Private Sub optRad_Click()

optRad.Value = True
optDeg.Value = False
CheckKeyBoardStatus

End Sub

Private Sub optRad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = "Radians"

End Sub

Private Sub SB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

SB.SimpleText = Format(Now, "                       " _
 & "                                                " _
 & "               dddd, dd mmm yyyy    hh:mm") & " Hrs"

End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)

Const Zero = 48
Const One = 49
Const Two = 50
Const Three = 51
Const Four = 52
Const Five = 53
Const Six = 54
Const Seven = 55
Const Eight = 56
Const Nine = 57
Const Multiply = 42
Const Add = 43
Const Divide = 47
Const Subtract = 45
Const Equals = 61
Const Factorial = 33
Const Deci = 46
Const Back = 8

Select Case KeyAscii
  Case Is = Zero
            cmd0_Click
  Case Is = One
            cmd1_Click
  Case Is = Two
            cmd2_Click
  Case Is = Three
            cmd3_Click
  Case Is = Four
            cmd4_Click
  Case Is = Five
            cmd5_Click
  Case Is = Six
            cmd6_Click
  Case Is = Seven
            cmd7_Click
  Case Is = Eight
            cmd8_Click
  Case Is = Nine
            cmd9_Click
  Case Is = Deci
            cmdDecimal_Click
  Case Is = Multiply
            cmdMul_Click
  Case Is = Add
            cmdPlus_Click
  Case Is = Divide
            cmdDiv_Click
  Case Is = Subtract
            cmdSub_Click
  Case Is = Factorial
            cmdFactorial_Click
  Case Is = Equals
            cmdEquals_Click
Case Is = Back
            cmdBack_Click
End Select

End Sub

Private Sub ValidCharacter(KeyAscii As Integer)

'Ensures that the user can only input the values defined
'in the strValid variable.
Dim strValid As String
strValid = "0123456789."
If KeyAscii > 26 Then 'if it is not a control code
  If InStr(strValid, Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
    Beep
  End If
End If

End Sub
