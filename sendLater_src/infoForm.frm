VERSION 5.00
Begin VB.Form InfoForm 
   Caption         =   "'SendLater!'  Ver1.0  (information)"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label5 
      Caption         =   "Close this window"
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
      Left            =   3285
      TabIndex        =   4
      Top             =   1845
      Width           =   1365
   End
   Begin VB.Label Label4 
      Caption         =   "pl_harish@hotmail.com"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   3
      Top             =   1485
      Width           =   2940
   End
   Begin VB.Label Label3 
      Caption         =   "For more information, email me at,"
      Height          =   240
      Left            =   135
      TabIndex        =   2
      Top             =   1260
      Width           =   2400
   End
   Begin VB.Label Label2 
      Caption         =   "This is a freeWare. Due to size constraints, 'Help ?' and support for this freeware is not provided along with this application."
      Height          =   690
      Left            =   135
      TabIndex        =   1
      Top             =   495
      Width           =   4155
   End
   Begin VB.Label Label1 
      Caption         =   "Thanks for using SendLater! Ver1.0"
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   2670
   End
End
Attribute VB_Name = "InfoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label5_Click()
    Me.Hide
    Unload Me
End Sub
