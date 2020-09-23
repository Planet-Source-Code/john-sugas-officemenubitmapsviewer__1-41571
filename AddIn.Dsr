VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10020
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   11385
   _ExtentX        =   20082
   _ExtentY        =   17674
   _Version        =   393216
   Description     =   "Add-In Project Template"
   DisplayName     =   "OfficeBarBitmaps"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Office Menu Bitmaps VbAdd-In by John Sugas 2002
'Purpose: To Display the built-in bitmaps that are available to the VbIde.
'Use: Compile DLL, GoTo Add-In manager and load the OfficeMenuBitmaps Add-in
'Add-in will now be avail. in the Add-Ins menu. To copy or save, right click the bitmap.

Public VBInstance As VBIDE.VBE
Private mcbMenuCommandBar As Office.CommandBarControl
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

Public cmdButton As Office.CommandBarButton

Private cbMenuCommandBar As Office.CommandBarControl  'command bar object
Private cbMenu As Object

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = Application

    If ConnectMode = ext_cm_External Then
    Else
'        ' Create the button
        Set cmdButton = VBInstance.CommandBars("Code Window").Controls.Add(msoControlButton)
        cmdButton.Tag = "cmdButton"
        cmdButton.Visible = False

        Set mcbMenuCommandBar = AddToAddInCommandBar("OfficeBarBitmaps")
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
  
    Exit Sub
    
error_handler:
    MsgBox Err.Description & vbCrLf & "Sub AddinInstance_OnConnection"
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    Dim i As Long
    
    On Error Resume Next
    For i = VBInstance.CommandBars("Code Window").Controls.Count To 1 Step -1
        If VBInstance.CommandBars("Code Window").Controls(i).Tag = "cmdButton" Then
            VBInstance.CommandBars("Code Window").Controls(i).Delete
        End If
    Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    Set mcbMenuCommandBar = Nothing
    Set MenuHandler = Nothing
    Set Form1.cc = Nothing
    Set Form1 = Nothing
    Set cbMenuCommandBar = Nothing
    Set cbMenu = Nothing
    Set cmdButton = Nothing
    Set VBInstance = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
    End If
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Set Form1.cc = Me
    Form1.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim i As Long
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    'make sure it doesn't already exist
    For i = cbMenu.Controls.Count To 1 Step -1
        If cbMenu.Controls(i).Caption = sCaption Then
            cbMenu.Controls(i).Delete
        End If
    Next
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    ' add the icon
    cmdButton.FaceId = 2170 'grab any bmp for now
    cmdButton.CopyFace 'copy to clipboard
    cbMenuCommandBar.PasteFace
    Set AddToAddInCommandBar = cbMenuCommandBar
    Clipboard.Clear
    Exit Function
AddToAddInCommandBarErr:
    MsgBox Err.Description & vbCrLf & "Function AddToAddInCommandBar"
End Function




