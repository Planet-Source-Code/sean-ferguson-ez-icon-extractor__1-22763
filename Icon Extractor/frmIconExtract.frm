VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIconExtract 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Extractor"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7530
   Icon            =   "frmIconExtract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar verticalScroll 
      Enabled         =   0   'False
      Height          =   5535
      Left            =   7260
      Max             =   0
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin MSComDlg.CommonDialog commonDialog 
      Left            =   0
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image picLib 
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   0
      Left            =   -480
      Top             =   120
      Width           =   480
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Icon Library..."
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmIconExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
commonDialog.Flags = &H400 + &H4 + &H8 + &H2 + &H800
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuFileOpen_Click()
commonDialog.FileName = ""
commonDialog.DialogTitle = "Open Image Library"
commonDialog.Filter = "Icon Libraries|*.ico;*.bmp;*.gif;*.jpg;*.exe;*.dll"
commonDialog.ShowOpen
If commonDialog.FileName = "" Then Exit Sub
LoadAllIcons commonDialog.FileName
End Sub

Private Sub LoadAllIcons(iconLibFN As String)
On Error GoTo handleError
Dim errOccured As Boolean
Dim picNum As Integer
If picLib.Count > 1 Then
    For picNum = 1 To picLib.Count
        Unload picLib(picNum)
    Next picNum
End If
picNum = 1
While Not errOccured
    Load picLib(picNum)
    rtnVal = GetIconFromFile(iconLibFN, CLng(picNum - 1), True)
    If rtnVal = 0 Then errOccured = True: GoTo setIcon
    Set picLib(picNum).Picture = GetIconFromFile(iconLibFN, CLng(picNum - 1), True)
    SetIconPos picNum
    picNum = picNum + 1
Wend
Exit Sub

handleError:
Resume Next

setIcon:
Unload picLib(picNum)
If picLib.Count = 1 Then MsgBox "The file you selected does not exist or contains no icons.", vbOKOnly + vbExclamation, "Error": Exit Sub
ShowIcons 1
Exit Sub
End Sub

Private Sub SetIconPos(picNum As Integer)
    If (picLib(picNum - 1).Left + 580) > (verticalScroll.Left - 300) Then
        If (picLib(picNum - 1).Top + 1580) > Me.Height Then
            picLib(picNum).Top = 120
            picLib(picNum).Visible = False
        Else
            picLib(picNum).Top = picLib(picNum - 1).Top + 580
        End If
        picLib(picNum).Left = 120
    Else
        picLib(picNum).Top = picLib(picNum - 1).Top
        picLib(picNum).Left = picLib(picNum - 1).Left + 580
    End If
    If picLib.Count > 109 Then verticalScroll.Enabled = True Else verticalScroll.Enabled = False
    If verticalScroll.Enabled = True Then verticalScroll.Max = Int((picLib.Count - 1) / 108): verticalScroll.Value = 0
End Sub

Private Sub ShowIcons(iconStart As Integer)
    Dim imgNum As Integer
    Dim endVal As Integer
    For imgNum = 0 To (iconStart - 1)
        picLib(imgNum).Visible = False
    Next imgNum
    If (iconStart + 107) >= picLib.Count - 1 Then endVal = (picLib.Count - 1) Else endVal = (iconStart + 107)
    For imgNum = iconStart To endVal
        picLib(imgNum).Visible = True
    Next imgNum
    If imgNum < (picLib.Count - 1) Then
        For imgNum = (iconStart + 109) To (picLib.Count - 1)
            picLib(imgNum).Visible = False
        Next imgNum
    End If
End Sub

Private Sub picLib_DblClick(Index As Integer)
commonDialog.DialogTitle = "Save Icon"
commonDialog.Filter = "Icon (*.ico)|*.ico"
commonDialog.FileName = ""
commonDialog.ShowSave
If commonDialog.FileName = "" Then Exit Sub
SavePicture picLib(Index), commonDialog.FileName
End Sub

Private Sub verticalScroll_Change()
ShowIcons verticalScroll.Value * 109
End Sub
