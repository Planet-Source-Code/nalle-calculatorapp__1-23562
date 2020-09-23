VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "        ""Calculator"""
   ClientHeight    =   4185
   ClientLeft      =   570
   ClientTop       =   2250
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picScroll 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      Height          =   2535
      Left            =   360
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   341
      TabIndex        =   0
      Top             =   720
      Width           =   5175
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   4800
      TabIndex        =   2
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Calulator with receiptfunction"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   2040
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code and form was downloaded from PSC

Option Explicit
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Const DT_BOTTOM As Long = &H8
Const DT_CALCRECT As Long = &H400
Const DT_CENTER As Long = &H1
Const DT_EXPANDTABS As Long = &H40
Const DT_EXTERNALLEADING As Long = &H200
Const DT_LEFT As Long = &H0
Const DT_NOCLIP As Long = &H100
Const DT_NOPREFIX As Long = &H800
Const DT_RIGHT As Long = &H2
Const DT_SINGLELINE As Long = &H20
Const DT_TABSTOP As Long = &H80
Const DT_TOP As Long = &H0
Const DT_VCENTER As Long = &H4
Const DT_WORDBREAK As Long = &H10

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Const ScrollText As String = "Calculator" & vbCrLf & _
                                vbCrLf & vbCrLf & _
                             "Producer: Björn Johansson" & vbCrLf & _
                                vbCrLf & _
                             "Do you miss any functionality ???" & _
                                vbCrLf & "Send me a mail !!!" & _
                                vbCrLf & "d1001.johansson@swipnet.se " & _
                                vbCrLf & vbCrLf & _
                                vbCrLf & "Muuuuuuuuuuu "
                             
Dim EndingFlag As Boolean

Private Sub Form_Activate()
RunMain
End Sub

Private Sub Form_Load()
    picScroll.ForeColor = vbYellow
    picScroll.FontSize = 14
End Sub

Private Sub RunMain()
Dim LastFrameTime As Long
Const IntervalTime As Long = 40
Dim rt As Long
Dim DrawingRect As RECT
Dim UpperX As Long, UpperY As Long 'Upper left point of drawing rect
Dim RectHeight As Long

'show the form
frmAbout.Refresh

'Get the size of the drawing rectangle by suppying the DT_CALCRECT constant
rt = DrawText(picScroll.hdc, ScrollText, -1, DrawingRect, DT_CALCRECT)

If rt = 0 Then 'err
    MsgBox "Error scrolling text", vbExclamation
    EndingFlag = True
Else
    DrawingRect.Top = picScroll.ScaleHeight
    DrawingRect.Left = 0
    DrawingRect.Right = picScroll.ScaleWidth
    'Store the height of The rect
    RectHeight = DrawingRect.Bottom
    DrawingRect.Bottom = DrawingRect.Bottom + picScroll.ScaleHeight
End If


Do While Not EndingFlag
    
    If GetTickCount() - LastFrameTime > IntervalTime Then
                    
        picScroll.Cls
        
        DrawText picScroll.hdc, ScrollText, -1, DrawingRect, DT_CENTER Or DT_WORDBREAK
        
        'update the coordinates of the rectangle
        DrawingRect.Top = DrawingRect.Top - 1
        DrawingRect.Bottom = DrawingRect.Bottom - 1
        
        'control the scolling and reset if it goes out of bounds
        If DrawingRect.Top < -(RectHeight) Then 'time to reset
            DrawingRect.Top = picScroll.ScaleHeight
            DrawingRect.Bottom = RectHeight + picScroll.ScaleHeight
        End If
        
        picScroll.Refresh
        
        LastFrameTime = GetTickCount()
        
    End If
    
    DoEvents
Loop

Unload Me
Set frmAbout = Nothing

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = vbYellow
End Sub

Private Sub Form_Unload(Cancel As Integer)

    EndingFlag = True
   
End Sub

Private Sub lblExit_Click()

Beep

EndingFlag = True

End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblExit.ForeColor = vbRed
End Sub
