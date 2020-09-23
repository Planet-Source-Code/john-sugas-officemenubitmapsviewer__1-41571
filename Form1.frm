VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Office Menu Bitmaps (Options-Right Click)"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   3930
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5424
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPics 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      DrawWidth       =   2
      Height          =   3930
      Left            =   0
      LinkTimeout     =   0
      ScaleHeight     =   258
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   396
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   6000
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   0
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   435
         Index           =   0
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.VScrollBar VScrollBar1 
      Height          =   3930
      LargeChange     =   38
      Left            =   6000
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuClipboardSave 
         Caption         =   "Save To ClipBoard"
      End
      Begin VB.Menu mnuFileSaveBitmap 
         Caption         =   "Save To .BMP"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cc As Connect
Private iTotal As Long, bmpW As Long, bmpH As Long
Private iClickIndex As Long

Private Sub Form_Load()
    Dim i As Long
    
    cc.cmdButton.FaceId = 2170 'grab any bmp for now
    cc.cmdButton.CopyFace 'copy to clipboard
    Image1(0).Picture = Clipboard.GetData(vbCFBitmap) 'put it in the temp image
    Image1(0).Stretch = True 'true size it
    bmpW = Image1(0).Width 'got the width
    bmpH = Image1(0).Height 'got the height
    Set Picture1.Picture = Image1(0).Picture
    
    On Error Resume Next  'gonna error when we try to access +1 the total pics
    Do
        iTotal = iTotal + 1
        cc.cmdButton.FaceId = iTotal
    Loop While Err.Number = 0 'hit the max
    iTotal = iTotal - 1 'true total
    StatusBar1.Panels(1).Text = "Total:" & iTotal
    picPics.Move 0, 0
    Show 'show the form now so user will not think nothing is happening
    DoEvents
    DisplayBitmaps 'show'em
    Clipboard.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 1 To Image1.Count
        Unload Image1(i) 'unload the image ctrls
    Next
    Set Form1 = Nothing
End Sub

Private Sub DisplayBitmaps()

    Dim iDx As Long, H As Long, W As Long, Ub As Long, iOldH As Long
    Dim iRows As Long, rRows As Single, iCols As Long
    
    Screen.MousePointer = vbHourglass
    iCols = picPics.ScaleWidth \ (bmpW + 6) '# of columns
    rRows = iTotal / iCols '# of rows floating point
    iRows = iTotal \ iCols '# of rows integer
    If rRows > iRows Then
        iRows = iRows + 1 'add 1 so we don't chop off the last row
    End If
    iOldH = picPics.Height 'original height of picBox
    picPics.Height = iRows * (bmpH + 6) 'total height for scrollable picBox
    VScrollBar1.Enabled = 1
    VScrollBar1.Max = (iRows * (bmpH + 6)) - iOldH + 2 'set sb max val.

    H = 3
    W = 3
    VScrollBar1.Value = 0
    
    For iDx = 1 To iTotal
        Load Image1(iDx) 'new image
        cc.cmdButton.FaceId = iDx
        cc.cmdButton.CopyFace 'copy to clipBoard
        Image1(iDx).Picture = Clipboard.GetData(vbCFBitmap) 'get from clipboard
        Image1(iDx).Stretch = True 'set the dimentions
        Image1(iDx).Top = H 'set the position
        Image1(iDx).Left = W
        Image1(iDx).Visible = True
        W = W + bmpW + 6 'add for border
        If (iDx) Mod iCols = 0 Then 'new row
           W = 3
           H = H + bmpH + 6
        End If
    Next

    Screen.MousePointer = vbDefault
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index > iTotal Then Exit Sub
    StatusBar1.Panels(3).Text = "FaceId=" & Index
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    iClickIndex = Index
    If Button = vbRightButton Then
        PopupMenu mnuPopUp
    End If
End Sub

Private Sub mnuClipboardSave_Click()
    cc.cmdButton.FaceId = iClickIndex
    cc.cmdButton.CopyFace 'copy to clipboard
    MsgBox "Bitmap FaceId:" & iClickIndex & "  Has been saved to ClipBoard..."
End Sub

Private Sub mnuFileSaveBitmap_Click()
    SavePicture Image1(iClickIndex).Picture, App.Path & "\FaceId" & iClickIndex & ".bmp"
    MsgBox "Bitmap FaceId:" & iClickIndex & "  Has been saved to File..."
End Sub

Private Sub VScrollBar1_Change()
    picPics.Top = -VScrollBar1.Value
End Sub
Private Sub VScrollBar1_Scroll()
    picPics.Top = -VScrollBar1.Value
End Sub



