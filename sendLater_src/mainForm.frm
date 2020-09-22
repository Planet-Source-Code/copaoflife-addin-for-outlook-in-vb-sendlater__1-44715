VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form mainForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "'SendLater!'  Ver1.0  (options)"
   ClientHeight    =   3210
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton buttSendImmdly 
      Caption         =   "&Send Immdly!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3690
      TabIndex        =   0
      Top             =   2655
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Send Later !"
      Height          =   1770
      Left            =   135
      TabIndex        =   5
      Top             =   720
      Width           =   5550
      Begin VB.ComboBox sMeridien 
         Height          =   315
         ItemData        =   "mainForm.frx":0000
         Left            =   2430
         List            =   "mainForm.frx":000A
         TabIndex        =   12
         Top             =   540
         Width           =   600
      End
      Begin VB.TextBox sMins 
         Height          =   315
         Left            =   1980
         MaxLength       =   2
         TabIndex        =   11
         Top             =   540
         Width           =   375
      End
      Begin VB.TextBox sHrs 
         Height          =   315
         Left            =   1530
         MaxLength       =   2
         TabIndex        =   10
         Top             =   540
         Width           =   375
      End
      Begin VB.CommandButton buttSendLater 
         Caption         =   "Send Later!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3555
         TabIndex        =   2
         Top             =   1305
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker sDate 
         Height          =   315
         Left            =   135
         TabIndex        =   1
         Top             =   540
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22740993
         CurrentDate     =   37724
      End
      Begin VB.Label Label7 
         Caption         =   "The mail is sent on the date, and time specified.      System time is taken as the reference."
         Height          =   1005
         Left            =   3375
         TabIndex        =   7
         Top             =   180
         Width           =   2085
      End
      Begin VB.Label Label3 
         Caption         =   "Hrs : Mins"
         Height          =   195
         Left            =   1620
         TabIndex        =   6
         Top             =   360
         Width           =   1410
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Help ?"
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
      Left            =   4275
      TabIndex        =   9
      Top             =   225
      Width           =   600
   End
   Begin VB.Label Label8 
      Caption         =   "About SendLater!"
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
      Left            =   4275
      TabIndex        =   8
      Top             =   0
      Width           =   1365
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5670
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label2 
      Caption         =   "Press 'ALT + S' to send immediately."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Top             =   2685
      Width           =   3300
   End
   Begin VB.Label Label1 
      Caption         =   "You can select to send this mail later, 'OR', send it immediately."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   135
      TabIndex        =   3
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'API call
'for bringing form to foreground for popupmenu to work well
Private Declare Function SetForegroundWindow Lib "user32" _
(ByVal hwnd As Long) As Long 'sets a window active, in front

'Public VBInstance As VBIDE.VBE
Public Connect As Connect
Public objApplication As New Outlook.Application
Private frmInfoForm As New InfoForm

Private Sub buttSendImmdly_Click()
    Connect.sDateTimeToSend = ""
    Me.Hide
End Sub

Private Sub buttSendLater_Click()
   
    Connect.sDateTimeToSend = ""
    
    'validating user input
    If (sHrs > 12) Then
        Call MsgBox("Please enter a valid Hr ( <= 12 )", vbExclamation)
        sHrs.SetFocus
    ElseIf (sMins > 59) Then
        Call MsgBox("Please enter a valid Hr ( <= 59 )", vbExclamation)
        sMins.SetFocus
    End If
    
    'if user selects a date or time which is less than current datetime,
    'just send the mail immdly.
    
    Connect.sDateTimeToSend = CStr(DateValue(sDate.Value) & " " & sHrs & ":" & sMins & ":00" & " " & sMeridien.Text)
    
    Me.Hide
End Sub

Private Sub Form_Load()
    sDate.MinDate = Now
    sDate.MaxDate = DateAdd("m", 3, Now)
    
    sHrs = DatePart("h", Now)
    sMins = DatePart("n", Now)
    
    sMeridien.Clear
    Call sMeridien.AddItem("AM", 0)
    Call sMeridien.AddItem("PM", 1)
    
    If CInt(sHrs) > 12 Then
        sHrs = CInt(sHrs) - 12
        sMeridien.ListIndex = 1
    Else
        sMeridien.ListIndex = 0
    End If
    SetForegroundWindow (Me.hwnd)
End Sub

Private Sub Label8_Click()
    frmInfoForm.Show (1)
End Sub

Private Sub Label9_Click()
    frmInfoForm.Show (1)
End Sub

Private Sub sHrs_KeyDown(KeyCode As Integer, Shift As Integer)
    If checkNumber(KeyCode) Then
        sHrs.Text = sHrs.Text & Chr(KeyCode)
    End If
End Sub

Private Function checkNumber(iKeyCode As Integer, Optional bHideBox As Boolean = True) As Boolean
    If (iKeyCode >= 48 And iKeyCode <= 57) Then
        'no problem
        checkNumber = True
    Else
        If Not bHideBox Then
            Call MsgBox("Please enter a valid number", vbExclamation)
        End If
        checkNumber = False
    End If
End Function

Private Sub sMins_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not checkNumber(KeyCode) Then
        sMins.Text = sMins.Text & Chr(KeyCode)
    End If
End Sub

