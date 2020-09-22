VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9945
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   6585
   _ExtentX        =   11615
   _ExtentY        =   17542
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "Send Later !  Ver1.0"
   AppName         =   "Microsoft Outlook"
   AppVer          =   "Microsoft Outlook 9.0"
   LoadName        =   "Startup"
   LoadBehavior    =   3
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook"
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public objApplication As Outlook.Application
Public WithEvents evOutlookSend As Outlook.Application
Attribute evOutlookSend.VB_VarHelpID = -1
Public sDateTimeToSend As String
Private frmMainform As New mainForm

'this method is called when the addin has to be loaded.
'...when it has to be loaded is defined in the Designer.
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'link the outlook application object to the object in this addin.
    Set objApplication = Application

    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    'Debug.Print objApplication.Name
    
    'link the outlook application's events to handle them in through addin code.
    Set evOutlookSend = objApplication
  
    Exit Sub
    
error_handler:
    'handle errors
    'place debug code
End Sub

'this method is called when the addin has to be disconnected from outlook
'when outlook shuts down.
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    cleanUp
End Sub

'this event is triggerred when the 'send' button in a mail window is clicked.
Private Sub evOutlookSend_ItemSend(ByVal Item As Object, Cancel As Boolean)

    'show form
    Set frmMainform.objApplication = objApplication 'link form to application(outlook).
    Set frmMainform.Connect = Me    'link form's connect object to this object.
    frmMainform.Show (1)
    
    'destroy form
    Unload frmMainform
    Set frmMainform = Nothing
    
    'do process
    If sDateTimeToSend <> "" Then
        Item.DeferredDeliveryTime = CDate(sDateTimeToSend)
    End If
End Sub

Private Sub cleanUp()
    Set evOutlookSend = Nothing
    Set objApplication = Nothing
End Sub
