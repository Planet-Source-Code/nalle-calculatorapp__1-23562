VERSION 5.00
Begin VB.Form frmHjalpKalk 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "       "
   ClientHeight    =   5205
   ClientLeft      =   1755
   ClientTop       =   3630
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTillbaka 
      Caption         =   "Back to Calculator"
      Height          =   375
      Left            =   2520
      TabIndex        =   23
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print this form"
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Reduce Swedish VAT"
      Height          =   255
      Index           =   14
      Left            =   1320
      TabIndex        =   31
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Swedish VAT"
      Height          =   255
      Index           =   13
      Left            =   1320
      TabIndex        =   30
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label lblFText 
      Height          =   255
      Index           =   12
      Left            =   1320
      TabIndex        =   29
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F12"
      Height          =   255
      Index           =   14
      Left            =   360
      TabIndex        =   28
      Top             =   3480
      Width           =   400
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F11"
      Height          =   255
      Index           =   13
      Left            =   360
      TabIndex        =   27
      Top             =   3240
      Width           =   400
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F10"
      Height          =   255
      Index           =   12
      Left            =   360
      TabIndex        =   26
      Top             =   3000
      Width           =   400
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F9"
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   25
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "To Clipboard if you want to use it in another application"
      Height          =   495
      Index           =   11
      Left            =   1320
      TabIndex        =   24
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      Height          =   255
      Index           =   10
      Left            =   1320
      TabIndex        =   21
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label lblFText 
      Height          =   255
      Index           =   9
      Left            =   1320
      TabIndex        =   20
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "Esc el End"
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   19
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblFKnapp 
      Caption         =   " "
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   18
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Readout only"
      Height          =   255
      Index           =   8
      Left            =   1320
      TabIndex        =   17
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F8"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   16
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Readout and receipt"
      Height          =   255
      Index           =   7
      Left            =   1320
      TabIndex        =   15
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "No receipt"
      Height          =   255
      Index           =   6
      Left            =   1320
      TabIndex        =   14
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt"
      Height          =   255
      Index           =   5
      Left            =   1320
      TabIndex        =   13
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "% - button"
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   12
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "CE - button"
      Height          =   255
      Index           =   3
      Left            =   1320
      TabIndex        =   11
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "C -  button"
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   10
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   9
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   8
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F7"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "Button"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F6"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F5"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F4 or Home"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F3"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 or Del"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F1"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "frmHjalpKalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
    frmHjalpKalk.PrintForm
End Sub

Private Sub cmdTillbaka_Click()
    Unload Me
End Sub
