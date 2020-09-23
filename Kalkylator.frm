VERSION 5.00
Begin VB.Form frmKalkylator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Calculator"
   ClientHeight    =   3540
   ClientLeft      =   6465
   ClientTop       =   1110
   ClientWidth     =   5310
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Kalkylator.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3540
   ScaleWidth      =   5310
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdMuMinus 
      Caption         =   "VAT-"
      Height          =   300
      Left            =   2200
      TabIndex        =   36
      ToolTipText     =   "F12 or Page Down"
      Top             =   2880
      Width           =   550
   End
   Begin VB.CommandButton cmdMuPlus 
      Caption         =   "VAT+"
      Height          =   300
      Left            =   120
      TabIndex        =   35
      ToolTipText     =   "F11 or Page Up"
      Top             =   2880
      Width           =   550
   End
   Begin VB.CommandButton cmdPrint2 
      Caption         =   "Print"
      Height          =   210
      Left            =   0
      TabIndex        =   34
      ToolTipText     =   "P"
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdTillbaka 
      Caption         =   "Back to normal"
      Height          =   210
      Left            =   0
      TabIndex        =   33
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdSlagremsa 
      Caption         =   "Readout only"
      Height          =   210
      Left            =   3250
      TabIndex        =   32
      ToolTipText     =   "F7"
      Top             =   3010
      Width           =   1695
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "To Clipboard"
      Height          =   210
      Left            =   720
      TabIndex        =   31
      ToolTipText     =   "F9"
      Top             =   3120
      Width           =   1400
   End
   Begin VB.CommandButton cmdProc 
      Caption         =   "%"
      Height          =   360
      Left            =   1700
      TabIndex        =   30
      ToolTipText     =   "F4"
      Top             =   2280
      Width           =   420
   End
   Begin VB.CommandButton cmdKvitto 
      Caption         =   "Receipt    >"
      Height          =   210
      Left            =   700
      TabIndex        =   29
      ToolTipText     =   "F5"
      Top             =   2860
      Width           =   1400
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   210
      Left            =   3250
      TabIndex        =   28
      ToolTipText     =   "P"
      Top             =   3250
      Width           =   1695
   End
   Begin VB.CommandButton cmdEjKvitto 
      Caption         =   "<   No receipt"
      Height          =   210
      Left            =   3250
      TabIndex        =   27
      ToolTipText     =   "F6"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.ListBox lstKvitto 
      Height          =   2370
      Left            =   3100
      TabIndex        =   26
      Top             =   240
      Width           =   2100
   End
   Begin VB.CommandButton MemoKey 
      Caption         =   "MC"
      Enabled         =   0   'False
      Height          =   360
      Index           =   3
      Left            =   4800
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2280
      Width           =   420
      Visible         =   0   'False
   End
   Begin VB.CommandButton MemoKey 
      Caption         =   "MR"
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   4800
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1800
      Width           =   420
      Visible         =   0   'False
   End
   Begin VB.CommandButton MemoKey 
      Caption         =   "M-"
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   4800
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1320
      Width           =   420
      Visible         =   0   'False
   End
   Begin VB.CommandButton MemoKey 
      Caption         =   "M+"
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   4800
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   840
      Width           =   420
      Visible         =   0   'False
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   100
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   840
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   600
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   840
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   1100
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   840
      Width           =   420
   End
   Begin VB.CommandButton Cancel 
      BackColor       =   &H00808080&
      Caption         =   "C"
      Height          =   360
      Left            =   1700
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "F2 or Del"
      Top             =   840
      Width           =   420
   End
   Begin VB.CommandButton CancelEntry 
      BackColor       =   &H00808080&
      Caption         =   "CE"
      Height          =   360
      Left            =   2200
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "F3"
      Top             =   840
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   100
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1320
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   600
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1320
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   1100
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1320
      Width           =   420
   End
   Begin VB.CommandButton Operator 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1700
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1320
      Width           =   420
   End
   Begin VB.CommandButton Operator 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2200
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1320
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   100
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1800
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   600
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1800
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   1100
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1800
      Width           =   420
   End
   Begin VB.CommandButton Operator 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   1700
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1800
      Width           =   420
   End
   Begin VB.CommandButton Operator 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   2200
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1800
      Width           =   420
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   100
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2280
      Width           =   900
   End
   Begin VB.CommandButton cmdDecimal 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1100
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2280
      Width           =   420
   End
   Begin VB.CommandButton Operator 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   2200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2280
      Width           =   420
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   250
      TabIndex        =   1
      Top             =   60
      Width           =   2295
      Begin VB.Label Readout 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   60
         TabIndex        =   2
         Top             =   180
         Width           =   2175
      End
   End
   Begin VB.CommandButton CopyButton 
      BackColor       =   &H00808080&
      Caption         =   "<"
      Height          =   315
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Copy to clipboard"
      Top             =   240
      Width           =   315
      Visible         =   0   'False
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      DrawMode        =   6  'Mask Pen Not
      X1              =   2900
      X2              =   2900
      Y1              =   0
      Y2              =   3480
   End
   Begin VB.Label lblMemoFlag 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "M"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4800
      TabIndex        =   25
      ToolTipText     =   "If M, memory  not zero"
      Top             =   50
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.Menu mnuAvsluta 
      Caption         =   "&Exit"
   End
   Begin VB.Menu mnuVisa 
      Caption         =   "&Show"
      Begin VB.Menu mnuMini 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuMiniKvitto 
         Caption         =   "Calculator with receipt"
      End
      Begin VB.Menu mnuKvittoUt 
         Caption         =   "Receipt with readout"
      End
      Begin VB.Menu mnuKvitoUtUt 
         Caption         =   "Receipt"
      End
   End
   Begin VB.Menu mnuHjalp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuOm 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmKalkylator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Original calculator app authored by Herman Lui
'I added some functions:
'Listbox and print
'Helpfunction
'F-keys
'Percentage-key
'VAT-key in Sweden VAT is 25 %
'Possibility to change layout during runtime, se Status.
'Always on top
'Some smaller changes
'Have a nice day

Option Explicit
'Högerjusterar
Private Const LB_SETTABSTOPS = &H192
Private Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Must be a long or integer array
Dim mTabs(0) As Long

'Originaldeklaration
Const Maxdigits = 16        ' After this, scientific notation
Dim Op1 As Variant          ' Prev input operand
Dim Op2 As Variant          ' Further prev input operand
Dim DecimalFlag As Integer  ' Decimal point present yet?
Dim NumOps As Integer       ' Numkey of operands, 0 to 2
Dim LastInput As String     ' Indicate type of last keypress event.
Dim OpFlag As String        ' Indicate pending operation.
Dim PrevReadout As String   ' For restore if "CE"
Dim MemoResult              ' Store result for memo keys
Dim XReadout As String
Dim XOp1 As Variant
Dim XOp2 As Variant
Dim XDecimalFlag As Integer
Dim XNumOps As Integer
Dim XLastInput As String
Dim XOpFlag As String
Dim XCaption As String
Dim XMemoResult
Dim KvittoFlag As String
Dim strTempreadout As String
Dim MinStatus As String
Dim Index As Integer
Dim KnappStatus As Integer
Dim PrevLastInput As String


Private Sub cmdCopy_Click()
    CopyButton_Click
End Sub

Private Sub cmdEjKvitto_Click()
    Call Status("Mini")
End Sub

Private Sub cmdKvitto_Click()
    Call Status("MiniKvitto")
End Sub

Private Sub cmdMuMinus_Click()
Dim moms As String
KnappStatus = 3
    Operator_Click 4
    moms = Readout * 0.2
    Readout = Readout * 0.8
'Aligment wright
SendMessage lstKvitto.hwnd, LB_SETTABSTOPS, 1, mTabs(0)
    lstKvitto.AddItem vbTab & "moms   - " + moms
    lstKvitto.AddItem vbTab & "= " + Readout
    lstKvitto.AddItem " "
    lstKvitto.Selected(lstKvitto.ListCount - 1) = True

End Sub

Private Sub cmdMuPlus_Click()
Dim moms As String
KnappStatus = 3
    Operator_Click 4
    moms = Readout * 0.25
    Readout = Readout * 1.25
'Aligment wright
SendMessage lstKvitto.hwnd, LB_SETTABSTOPS, 1, mTabs(0)
    lstKvitto.AddItem vbTab & "moms   + " + moms
    lstKvitto.AddItem vbTab & "= " + Readout
    lstKvitto.AddItem " "
    lstKvitto.Selected(lstKvitto.ListCount - 1) = True

End Sub

Private Sub cmdPrint_Click()
    Call Status("KvittoUtUt")
       frmKalkylator.lstKvitto.Appearance = 0
       frmKalkylator.BackColor = &H80000005
       cmdTillbaka.Visible = False
       cmdPrint2.Visible = False
       mnuAvsluta.Visible = False
       mnuHjalp.Visible = False
       mnuOm.Visible = False
        lstKvitto.Selected(lstKvitto.ListCount - 1) = False
             frmKalkylator.PrintForm
       frmKalkylator.lstKvitto.Appearance = 1
       frmKalkylator.BackColor = &H8000000F
       cmdTillbaka.Visible = True
       cmdPrint2.Visible = True
       mnuAvsluta.Visible = True
       mnuHjalp.Visible = True
       mnuOm.Visible = True
    Call Status("MiniKvitto")
End Sub

Private Sub cmdPrint2_Click()
       frmKalkylator.lstKvitto.Appearance = 0
       frmKalkylator.BackColor = &H80000005
       cmdTillbaka.Visible = False
       cmdPrint2.Visible = False
       mnuAvsluta.Visible = False
       mnuHjalp.Visible = False
       mnuOm.Visible = False
        lstKvitto.Selected(lstKvitto.ListCount - 1) = False
             frmKalkylator.PrintForm
       frmKalkylator.lstKvitto.Appearance = 1
       frmKalkylator.BackColor = &H8000000F
       cmdTillbaka.Visible = True
       cmdPrint2.Visible = True
       mnuAvsluta.Visible = True
       mnuHjalp.Visible = True
       mnuOm.Visible = True
    Call Status("KvittoUt")
End Sub

Private Sub cmdProc_Click()
KnappStatus = 3
    Operator_Click 4
    Readout = Readout / 100
'Aligment wright
SendMessage lstKvitto.hwnd, LB_SETTABSTOPS, 1, mTabs(0)
    lstKvitto.AddItem vbTab & "% "
    lstKvitto.AddItem vbTab & "= " + Readout
    lstKvitto.AddItem " "
    lstKvitto.Selected(lstKvitto.ListCount - 1) = True
End Sub

Private Sub cmdSlagremsa_Click()
    Call Status("KvittoUt")
End Sub

Private Sub cmdTillbaka_Click()
    Call Status("MiniKvitto")
End Sub

Private Sub Command1_Click()
Call Status("KvittoUtUt")
End Sub

Private Sub Form_Activate()
FormOnTop Me.hwnd, True

End Sub

Private Sub Form_Load()
Call Status("KvittoUt")
ResetStatus
'Alignment wright
'Set the a tab stop to a negative number
mTabs(0) = -65

End Sub


Sub ResetStatus()
    Readout = Format(0, "0")
    PrevReadout = Format(0, "0")
    Op1 = 0
    Op2 = 0
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
    lblMemoFlag.Caption = " "
    MemoResult = 0
End Sub

Sub RestoreStatus()
    Readout = XReadout
    Op1 = XOp1
    Op2 = XOp2
    DecimalFlag = XDecimalFlag
    NumOps = XNumOps
    LastInput = XLastInput
    OpFlag = XOpFlag
    lblMemoFlag.Caption = XCaption
    MemoResult = XMemoResult
End Sub


Sub MarkStatus()
    XReadout = Readout
    XOp1 = Op1
    XOp2 = Op2
    XDecimalFlag = DecimalFlag
    XNumOps = NumOps
    XLastInput = LastInput
    XOpFlag = OpFlag
    XCaption = lblMemoFlag.Caption
    XMemoResult = MemoResult
End Sub


Private Function MaxReached()
    MaxReached = False
    If Len(Readout) >= Maxdigits Then       ' Not allow further Numkey
         MaxReached = True
    End If
End Function


Function HasDecimal(strToRead As String)
    HasDecimal = False
    Dim i As Integer
    For i = Len(strToRead) To 1 Step -1
         If InStr(i, strToRead, ".") Then
             HasDecimal = True
             Exit For
         End If
    Next
End Function

' Copy the "Label" Caption onto the Clipboard.
Private Sub CopyButton_Click()
    Clipboard.SetText Readout
End Sub


Private Sub Cancel_Click()
    ResetStatus
    lstKvitto.Clear
    Operator(4).SetFocus
End Sub


Private Sub CancelEntry_Click()
    RestoreStatus
    LastInput = "CE"
    Operator(4).SetFocus
End Sub




Private Sub cmdDecimal_Click()
    If HasDecimal(Readout) Then             ' One is enough
        Exit Sub
    End If
    If LastInput = "NUMS" Or LastInput = "DIGI" Then
        If Len(Readout) = Maxdigits Then
            MsgBox "Maximalt antal siffror " & Str(Maxdigits - 1) + _
                vbCrLf & "Försök igen", , "  Miniräknare"
                Operator(4).SetFocus
            Exit Sub
        End If
    End If
    
'    Me.cmdDecimal.SetFocus
    MarkStatus
    
    If LastInput = "NEG" Then
        If Abs(Val(Readout)) <> 0 Then
            Readout = Format(0, "-0.")
        End If
    ElseIf LastInput <> "NUMS" And LastInput <> "DIGI" Then
        Readout = Format(0, "0.")
    End If
    
    DecimalFlag = True
    LastInput = "DIGI"
    
    If MaxReached Then
        MsgBox "Maximalt antal siffror " & Str(Maxdigits - 1) + _
           vbCrLf & " Försök igen", , "  Miniräknare"
        RestoreStatus
        Exit Sub
    End If
    Operator(4).SetFocus
End Sub



Private Sub mnuAvsluta_Click()
    End
End Sub

Private Sub mnuHjalp_Click()
 frmHjalpKalk.Show
End Sub

Private Sub mnuKvitoUtUt_Click()
Call Status("KvittoUtUt")
End Sub

Private Sub mnuKvittoUt_Click()
Call Status("KvittoUt")
End Sub

Private Sub mnuMini_Click()
 Call Status("Mini")
End Sub

Private Sub mnuMiniKvitto_Click()
 Call Status("MiniKvitto")
End Sub

Private Sub mnuOm_Click()
   frmAbout.Show vbModal
End Sub

Private Sub Numkey_Click(Index As Integer)
    If LastInput = "NUMS" Or LastInput = "DIGI" Then
        If MaxReached Then
            MsgBox "Maximalt antal siffror " & Str(Maxdigits - 1) + _
               vbCrLf & "Försök igen", , "  Miniräknare"
            Operator(4).SetFocus
            Exit Sub
        End If
    End If
    
'    Me.NumKey(Index).SetFocus
    MarkStatus
    If LastInput <> "NUMS" And LastInput <> "DIGI" Then
        Readout = Format(0, ".")
        DecimalFlag = False
    End If
    If DecimalFlag Then
        Readout = Readout + NumKey(Index).Caption
    Else
        Readout = Left(Readout, InStr(Readout, Format(0, ".")) - 1) + NumKey(Index).Caption + Format(0, ".")
    End If
    If LastInput = "NEG" Then
        Readout = "-" & Readout
    End If
    LastInput = "NUMS"
  KnappStatus = 1
    Operator(4).SetFocus
End Sub

Private Sub Operator_Click(Index As Integer)
'    Me.Operator(Index).SetFocus
    MarkStatus
    
    strTempreadout = Readout
    
    If LastInput = "NUMS" Or LastInput = "DIGI" Then
        NumOps = NumOps + 1
    End If
If OpFlag = "=" Then
KvittoFlag = " "
Else
   KvittoFlag = OpFlag + " "
   End If
    Select Case NumOps
        Case 0
            If Operator(Index).Caption = "-" And LastInput <> "NEG" Then
                If Abs(Val(Readout)) <> 0 Then
                    Readout = "-" & Readout
                    LastInput = "NEG"
                End If
            End If
        Case 1
            Op1 = Readout
            If Operator(Index).Caption = "-" And (LastInput <> "NUMS" _
                    And LastInput <> "DIGI") And OpFlag <> "=" Then
                If Abs(Val(Readout)) <> 0 Then
                    Readout = "-"
                    LastInput = "NEG"
                End If
            End If
        Case 2
            Op2 = strTempreadout
            Select Case OpFlag
                Case "+"
                    Op1 = CDbl(Op1) + CDbl(Op2)
                Case "-"
                    Op1 = CDbl(Op1) - CDbl(Op2)
                Case "*"
                    Op1 = CDbl(Op1) * CDbl(Op2)
                Case "/"
                    If Op2 = 0 Then
                       MsgBox "Division med noll ej möjligt", 48, "  Miniräknare"
                       RestoreStatus
                       Exit Sub
                    Else
                       Op1 = CDbl(Op1) / CDbl(Op2)
                    End If
               Case "="
                    Op1 = CDbl(Op2)
             End Select
             Readout = Op1
             NumOps = 1
             
    End Select
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = Operator(Index).Caption
    End If
    
     ' Be consistent, since we always show a decimal point
    If Not HasDecimal(Readout) Then
        If Abs(Val(Readout)) = 0 Then
           Readout = "0"
        Else
           Readout = Readout '+ "."
        End If
    End If
Call Kvitto
KnappStatus = 2
Operator(4).SetFocus
End Sub
Private Sub MemoKey_Click(Index As Integer)
    MarkStatus
    Select Case Index
       Case 0                    ' Memory Plus
            MemoResult = MemoResult + Val(Readout)
       Case 1                    ' Memory Minus
            MemoResult = MemoResult - Val(Readout)
       Case 2                    ' Memory Recall
            Dim s As String
            s = Str(MemoResult)
            If Not HasDecimal(Str(s)) Then
                s = s + "."
            End If
            Readout = s
       Case 3                    ' Memory Clear
            MemoResult = 0
    End Select
     ' Our system is, if MemoResult is not cleared, show "M"
    If MemoResult <> 0 Then
         lblMemoFlag.Caption = "M"
    Else
         lblMemoFlag.Caption = " "
    End If
    
    LastInput = "OPS"
    NumOps = 1
    Op1 = Readout
    Op2 = 0
    Operator(4).SetFocus
End Sub
' Detect keyboard key
Private Sub Form_KeyPress(keyascii As Integer)
    MarkStatus
    If keyascii < Asc("0") Or keyascii > Asc("9") Then
        If keyascii <> 46 And keyascii <> 43 And _
           keyascii <> 45 And keyascii <> 42 And _
           keyascii <> 47 And keyascii <> 61 And _
           keyascii <> 13 Then
               keyascii = 0
        Else
           Select Case keyascii
             Case 46                   ' "."
               cmdDecimal_Click
             Case 43
               Operator_Click (0)      ' re Property "+"
             Case 45                   ' "-"
               Operator_Click (1)
             Case 42                   ' "*"
               Operator_Click (2)
             Case 47                   ' "/"
               Operator_Click (3)
             Case 61                   ' "="
               Operator_Click (4)
             Case 13                   ' As "=" (if Windows allows Enter)
               Operator_Click (4)
           End Select
        End If
    Else
        Numkey_Click (Val(Chr(keyascii)))
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyDecimal
    cmdDecimal_Click
Case vbKeyDelete
    Cancel_Click
Case vbKeyEscape
    mnuAvsluta_Click
Case vbKeyEnd
mnuAvsluta_Click
Case vbKeyF1
    mnuHjalp_Click
Case vbKeyF2
    Cancel_Click
Case vbKeyF3
    CancelEntry_Click
Case vbKeyF4
    cmdProc_Click
Case vbKeyHome
    cmdProc_Click
Case vbKeyF5
    cmdKvitto_Click
Case vbKeyF6
    cmdEjKvitto_Click
Case vbKeyF7
    cmdSlagremsa_Click
Case vbKeyF8
    Call Status("KvittoUtUt")
Case vbKeyF9
    CopyButton_Click
Case vbKeyF11
    cmdMuPlus_Click
Case vbKeyPageUp
    cmdMuPlus_Click
Case vbKeyF12
    cmdMuMinus_Click
Case vbKeyPageDown
    cmdMuMinus_Click
Case vbKeyP
    cmdPrint_Click
End Select
End Sub

Private Sub Kvitto()
'Alignment wright
SendMessage lstKvitto.hwnd, LB_SETTABSTOPS, 1, mTabs(0)
If KnappStatus = 2 And OpFlag = "=" Then
    lstKvitto.AddItem vbTab & "= " + Readout
   Else
If KnappStatus = 2 Then
    KvittoFlag = "= "
    Else
If KnappStatus = 3 Then
    lstKvitto.AddItem vbTab & KvittoFlag + strTempreadout
    Else
If PrevLastInput = "NEG" And OpFlag = "=" Then
    lstKvitto.AddItem vbTab & strTempreadout
    lstKvitto.AddItem vbTab & "= " + Readout
    lstKvitto.AddItem vbTab & "  "
    Else
If LastInput = "OPS" And OpFlag = "=" Then
    lstKvitto.AddItem vbTab & KvittoFlag + strTempreadout
    lstKvitto.AddItem vbTab & "= " + Readout
    lstKvitto.AddItem vbTab & "  "
    Else
    lstKvitto.AddItem vbTab & KvittoFlag + strTempreadout
End If
End If
End If
End If
End If
PrevLastInput = LastInput
    lstKvitto.Selected(lstKvitto.ListCount - 1) = True
End Sub

Private Sub Status(AppStatus As String)
MinStatus = AppStatus
Select Case AppStatus
Case "MiniKvitto"
    Operator(1).Visible = True
    Operator(3).Visible = True
    Operator(4).Left = 2200
    Frame1.Top = 60
    Frame1.Left = 250
    frmKalkylator.Height = 4150
    frmKalkylator.Width = 5400
    lstKvitto.Height = 2370
    lstKvitto.Left = 3100
    lstKvitto.Top = 240
    NumKey(Index).Visible = True
    cmdDecimal.Visible = True
    Operator(Index).Visible = True
    Cancel.Visible = True
    CancelEntry.Visible = True
    cmdProc.Visible = True
    cmdKvitto.Visible = False
    cmdCopy.Visible = True
    cmdEjKvitto.Visible = True
    cmdPrint.Visible = True
    cmdTillbaka.Visible = False
    cmdPrint2.Visible = False
    Line1.Visible = True
    Readout.Visible = True
    Frame1.Visible = True
    lstKvitto.Visible = True
    cmdMuPlus.Visible = True
    cmdMuMinus.Visible = True
Case "Mini"
    Operator(1).Visible = True
    Operator(3).Visible = True
    Operator(4).Left = 2200
    Frame1.Top = 60
    Frame1.Left = 250
    frmKalkylator.Height = 4150
    frmKalkylator.Width = 2900
    lstKvitto.Height = 2370
    lstKvitto.Left = 3100
    lstKvitto.Top = 240
    NumKey(Index).Visible = True
    cmdDecimal.Visible = True
    Operator(Index).Visible = True
    Cancel.Visible = True
    CancelEntry.Visible = True
    cmdProc.Visible = True
    cmdKvitto.Visible = True
    cmdCopy.Visible = True
    cmdEjKvitto.Visible = True
    cmdPrint.Visible = True
    cmdTillbaka.Visible = False
    cmdPrint2.Visible = False
    Line1.Visible = False
    Readout.Visible = True
    Frame1.Visible = True
    lstKvitto.Visible = True
    cmdMuPlus.Visible = True
    cmdMuMinus.Visible = True
    
Case "KvittoUt"
    Operator(1).Visible = False
    Operator(3).Visible = False
    Operator(4).Left = 1000
    lstKvitto.Height = 5500
    lstKvitto.Left = 100
    lstKvitto.Top = 650
    frmKalkylator.Height = 7200
    frmKalkylator.Width = 2395
    NumKey(Index).Visible = False
    cmdDecimal.Visible = False
    Operator(Index).Visible = False
    Cancel.Visible = False
    CancelEntry.Visible = False
    cmdProc.Visible = False
    cmdKvitto.Visible = False
    cmdCopy.Visible = False
    cmdEjKvitto.Visible = False
    cmdPrint.Visible = False
    cmdTillbaka.Visible = True
    cmdTillbaka.Top = 6050
    cmdTillbaka.Left = 250
    cmdPrint2.Visible = True
    cmdPrint2.Top = 6300
    cmdPrint2.Left = 250
    Line1.Visible = False
    Frame1.Top = 0
    Frame1.Visible = True
    Frame1.Left = 0
    lstKvitto.Visible = True
    cmdMuPlus.Visible = False
    cmdMuMinus.Visible = False
Case "KvittoUtUt"
    Operator(1).Visible = False
    Operator(3).Visible = False
    Operator(4).Left = 1000
    lstKvitto.Height = 6130
    lstKvitto.Left = 100
    lstKvitto.Top = 0
    frmKalkylator.Height = 7200
    frmKalkylator.Width = 2395
    NumKey(Index).Visible = False
    cmdDecimal.Visible = False
    Operator(Index).Visible = False
    Cancel.Visible = False
    CancelEntry.Visible = False
    cmdProc.Visible = False
    cmdKvitto.Visible = False
    cmdCopy.Visible = False
    cmdEjKvitto.Visible = False
    cmdPrint.Visible = False
    cmdTillbaka.Visible = True
    cmdTillbaka.Top = 6050
    cmdTillbaka.Left = 250
    cmdPrint2.Visible = True
    cmdPrint2.Top = 6300
    cmdPrint2.Left = 250
    Line1.Visible = False
    Frame1.Visible = False
    lstKvitto.Visible = True
    cmdMuPlus.Visible = False
    cmdMuMinus.Visible = False

End Select
frmKalkylator.Refresh
frmKalkylator.ResetStatus

End Sub

