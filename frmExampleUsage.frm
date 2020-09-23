VERSION 5.00
Begin VB.Form frmExampleUsage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Example Usage"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Run Dialog"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "About"
         Height          =   375
         Left            =   2760
         TabIndex        =   31
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Show Dialog"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
      Begin VB.CheckBox ChReturn 
         Caption         =   "Checkbox return"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4080
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtReturned 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "Returned path"
         Top             =   840
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Settings"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.CheckBox ChSettings 
         Caption         =   "Include files"
         Height          =   255
         Index           =   9
         Left            =   4200
         TabIndex        =   30
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtNewfolder 
         Height          =   285
         Left            =   2400
         TabIndex        =   28
         Text            =   "New folder"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtCancelButton 
         Height          =   285
         Left            =   2400
         TabIndex        =   27
         Text            =   "Cancel"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtOKbutton 
         Height          =   285
         Left            =   2400
         TabIndex        =   26
         Text            =   "OK"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CheckBox ChSettings 
         Caption         =   "Status text"
         Height          =   255
         Index           =   8
         Left            =   4200
         TabIndex        =   23
         Top             =   2600
         Width           =   1335
      End
      Begin VB.CheckBox ChSettings 
         Caption         =   "Edit box (New style)"
         Height          =   255
         Index           =   7
         Left            =   4200
         TabIndex        =   22
         Top             =   2320
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.ComboBox cboSpecial 
         Height          =   315
         ItemData        =   "frmExampleUsage.frx":0000
         Left            =   1800
         List            =   "frmExampleUsage.frx":0049
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox ChSettings 
         Caption         =   "Edit box (Old style)"
         Height          =   255
         Index           =   6
         Left            =   4200
         TabIndex        =   19
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox ChSettings 
         Caption         =   "Checkbox"
         Height          =   255
         Index           =   5
         Left            =   4200
         TabIndex        =   18
         Top             =   1760
         Width           =   1335
      End
      Begin VB.CheckBox ChSettings 
         Caption         =   "New folder button"
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   17
         Top             =   1480
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox ChSettings 
         Caption         =   "Full screen"
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   16
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox ChSettings 
         Caption         =   "Double size"
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   15
         Top             =   920
         Width           =   1695
      End
      Begin VB.CheckBox ChSettings 
         Caption         =   "Center dialog"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   14
         Top             =   640
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox txtCheckBox 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Text            =   "Include subfolders"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtInitdir 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtDescript 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Text            =   "Select a folder"
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtTitlebar 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Text            =   "Browse for Folder"
         Top             =   360
         Width           =   2055
      End
      Begin VB.CheckBox ChSettings 
         Caption         =   "Allow resize"
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "New Folder Button caption :"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   2586
         Width           =   1980
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cancel Button caption :"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   2235
         Width           =   1665
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "OK Button caption :"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   1884
         Width           =   1395
      End
      Begin VB.Label Label5 
         Caption         =   "Root directory :"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1473
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Checkbox caption :"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   2940
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Initial path :"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1122
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User prompt :"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   771
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Titlebar text :"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   420
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmExampleUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboSpecial_Click()
On Error Resume Next
ResetMe
If cboSpecial.ListIndex = 0 Then
    txtInitdir.BackColor = vbWhite
    txtInitdir.Enabled = True
Else
    txtInitdir.BackColor = &H8000000F
    txtInitdir.Enabled = False
End If
cmdBrowse.SetFocus
End Sub
Private Sub ChSettings_Click(Index As Integer)
'just interface stuff here
ResetMe
Select Case Index
    Case 1
        If ChSettings(Index).Value = 1 Then ChSettings(3).Value = 0
    Case 2
        If ChSettings(Index).Value = 1 Then ChSettings(3).Value = 0
    Case 3
        If ChSettings(Index).Value = 1 Then
            ChSettings(1).Value = 0
            ChSettings(2).Value = 0
        End If
    Case 4
        If ChSettings(Index).Value = 1 Then ChSettings(5).Value = 0
    Case 5
        If ChSettings(Index).Value = 1 Then
            ChSettings(4).Value = 0
            txtCheckBox.BackColor = vbWhite
            txtCheckBox.Enabled = True
            ChReturn.Enabled = True
        Else
            txtCheckBox.BackColor = &H8000000F
            txtCheckBox.Enabled = False
            ChReturn.Enabled = False
        End If
    Case 6
        If ChSettings(Index).Value = 1 Then ChSettings(7).Value = 0
    Case 7
        If ChSettings(Index).Value = 1 Then ChSettings(6).Value = 0
End Select
End Sub
Private Sub cmdBrowse_Click()
'Fill in the variables before we make the call
ResetMe
'Dim bb As BoboBrowse
With BB
    'All these settings are optional
    'Leave all of them out and you are
    'left with the default Browse for Folders
    .TitleBar = txtTitlebar.Text
    .Prompt = txtDescript.Text
    .InitDir = txtInitdir.Text
    .CHCaption = txtCheckBox.Text
    .OKCaption = txtOKbutton.Text
    .CancelCaption = txtCancelButton.Text
    .NewFCaption = txtNewfolder.Text
    .RootDir = GetAroot
    .AllowResize = ChSettings(0).Value
    .CenterDlg = ChSettings(1).Value
    .DoubleSizeDlg = ChSettings(2).Value
    .FSDlg = ChSettings(3).Value
    .ShowButton = ChSettings(4).Value
    .ShowCheck = ChSettings(5).Value
    .EditBoxOld = ChSettings(6).Value
    .EditBoxNew = ChSettings(7).Value
    .StatusText = ChSettings(8).Value
    .ShowFiles = ChSettings(9).Value
    .OwnerForm = Me.hwnd
    'call the function
    txtReturned.Text = BrowseFF
    'If you included a checkbox this is where you
    'recieve the users' response
    ChReturn.Value = .CHvalue
End With
End Sub


Private Sub Form_Load()
cboSpecial.ListIndex = 0
txtInitdir.Text = CurDir
End Sub
Private Sub ResetMe()
ChReturn.Value = 0
txtReturned.Text = "Returned path"
End Sub
Private Sub txtCheckBox_KeyPress(KeyAscii As Integer)
ResetMe
End Sub
Private Sub txtCheckBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ResetMe
End Sub
Private Sub txtDescript_KeyPress(KeyAscii As Integer)
ResetMe
End Sub
Private Sub txtDescript_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ResetMe
End Sub
Private Sub txtInitdir_KeyPress(KeyAscii As Integer)
ResetMe
End Sub
Private Sub txtInitdir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ResetMe
End Sub
Private Sub txtTitlebar_KeyPress(KeyAscii As Integer)
ResetMe
End Sub
Private Sub txtTitlebar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ResetMe
End Sub
Private Function GetAroot() As Long
Select Case cboSpecial.Text 'special folders
    Case "Default"
        GetAroot = 0
    Case "Programs"
        GetAroot = 2
    Case "Control Panel"
        GetAroot = 3
    Case "Printers"
        GetAroot = 4
    Case "My Documents"
        GetAroot = 5
    Case "Favourites"
        GetAroot = 6
    Case "StartUp"
        GetAroot = 7
    Case "Recent"
        GetAroot = 8
    Case "SendTo"
        GetAroot = 9
    Case "Recycle Bin"
        GetAroot = 10
    Case "Start Menu"
        GetAroot = 11
    Case "Desktop"
        GetAroot = 16
    Case "My Computer"
        GetAroot = 17
    Case "Network"
        GetAroot = 18
    Case "NetHood"
        GetAroot = 19
    Case "Fonts"
        GetAroot = 20
    Case "Templates"
        GetAroot = 21
    Case "All users \ desktop"
        GetAroot = 25
    Case "Application Data"
        GetAroot = 26
    Case "PrintHood"
        GetAroot = 27
    Case "Temporary Internet Files"
        GetAroot = 32
    Case "Cookies"
        GetAroot = 33
    Case "History"
        GetAroot = 34
End Select
End Function
Private Sub Command1_Click()

MsgBox " BROWSE FOR FOLDERS ISSUES" + vbCrLf + String(29, "¯") + vbCrLf + _
"This module attempts to get around the following problems." + vbCrLf + vbCrLf + _
"1. No create new folder button under Win95/98." + vbCrLf + _
"2. The edit box in the standard Browse for Folder has limitations." + vbCrLf + _
"3. No method of displaying any extra parameters to the user." + vbCrLf + _
"4. No Resizing." + vbCrLf + vbCrLf + _
"Under Win2k you can use resize, have a new folder button and a sensible editbox," + vbCrLf + _
"however if your app is loaded under Win 95/98 no such luck! Its to do with the" + vbCrLf + _
"SHELL32.DLL Version 4.71 as opposed to the newer SHELL32.DLL Version 5.0." + vbCrLf + _
"This module attempts to duplicate the functionality of SHELL32.DLL Version 5.0" + vbCrLf + _
"when using SHELL32.DLL Version 4.71 (Win95/98)." + vbCrLf + vbCrLf + _
"This module falls short of SHELL32.DLL Version 5.0 in these areas :" + vbCrLf + _
"1. No context help for new items." + vbCrLf + _
"2. No context menu in the treeview." + vbCrLf + vbCrLf + _
"This module improves on SHELL32.DLL Version 5.0 in these areas :" + vbCrLf + _
"1. Customise button captions, titlebar caption" + vbCrLf + _
"2. Double size/Full screen" + vbCrLf + _
"3. Adds a checkbox", vbInformation, "Bobo Enterprises"
End Sub

