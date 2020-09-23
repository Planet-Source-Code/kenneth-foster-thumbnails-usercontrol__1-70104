VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Demo of  ucThumbNails...."
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Picture Dimensions"
      Height          =   210
      Left            =   6315
      TabIndex        =   21
      Top             =   2805
      Width           =   2085
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Rounded Corners"
      Height          =   195
      Left            =   6315
      TabIndex        =   20
      Top             =   3300
      Width           =   1650
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Show Folder Info"
      Height          =   225
      Left            =   6315
      TabIndex        =   16
      Top             =   3045
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse For Folder"
      Height          =   660
      Left            =   6315
      TabIndex        =   3
      Top             =   1020
      Width           =   2475
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enter above Path"
      Height          =   600
      Left            =   6315
      TabIndex        =   2
      Top             =   2160
      Width           =   2505
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6315
      TabIndex        =   1
      Top             =   1845
      Width           =   2505
   End
   Begin Project1.ucThumbNails ucThumbNails1 
      Height          =   7110
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   12541
      PicBoxBorderColor=   16711680
      FontColor       =   255
      BackColor       =   15790320
      ShowFolderInfo  =   0   'False
      ProBarColor     =   16744576
      PicDimen        =   0   'False
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on thumbnail to view preview. Double click on preview to close.      Left click and hold ,to drag preview."
      Height          =   615
      Left            =   6240
      TabIndex        =   19
      Top             =   7995
      Width           =   2610
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   6300
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   135
      Width           =   2535
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "PicBoxBackColor"
      Height          =   195
      Left            =   6990
      TabIndex        =   18
      Top             =   4980
      Width           =   1275
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6285
      TabIndex        =   17
      Top             =   4905
      Width           =   615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "BackColor"
      Height          =   210
      Left            =   7005
      TabIndex        =   15
      Top             =   4215
      Width           =   765
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6285
      TabIndex        =   14
      Top             =   4155
      Width           =   615
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Border Color"
      Height          =   255
      Left            =   6990
      TabIndex        =   13
      Top             =   4605
      Width           =   945
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "FontColor"
      Height          =   165
      Left            =   7005
      TabIndex        =   12
      Top             =   3855
      Width           =   735
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6285
      TabIndex        =   11
      Top             =   4530
      Width           =   615
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6285
      TabIndex        =   10
      Top             =   3795
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Folder Path"
      Height          =   195
      Left            =   7035
      TabIndex        =   9
      Top             =   6015
      Width           =   930
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   6240
      TabIndex        =   8
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected File"
      Height          =   180
      Left            =   6960
      TabIndex        =   7
      Top             =   5415
      Width           =   930
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Path"
      Height          =   210
      Left            =   7185
      TabIndex        =   6
      Top             =   6720
      Width           =   690
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1020
      Left            =   6240
      TabIndex        =   5
      Top             =   6930
      Width           =   2640
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6240
      TabIndex        =   4
      Top             =   5640
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  'FOR DEMO ONLY
 Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
   
   Private Type CHOOSECOLOR
   lStructSize As Long
   hwndOwner As Long
   hInstance As Long
   rgbResult As Long
   lpCustColors As String
   flags As Long
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type

Dim cc As CHOOSECOLOR

Private Type BROWSEINFO
    lngHwnd        As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32.dll" _
    (ByVal hMem As Long)

Private Declare Function lStrCat Lib "kernel32" _
   Alias "lstrcatA" (ByVal lpString1 As String, _
   ByVal lpString2 As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" _
   (lpbi As BROWSEINFO) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
   (ByVal pidList As Long, ByVal lpBuffer As String) As Long
   
Public Function BrowseForFolder(ByVal lngHwnd As Long, ByVal strPrompt As String) As String

    On Error GoTo ehBrowseForFolder 'Trap for errors

    Dim intNull As Integer
    Dim lngIDList As Long, lngResult As Long
    Dim strPath As String
    Dim udtBI As BROWSEINFO
    
    'Set API properties (housed in a UDT)
    With udtBI
        .lngHwnd = lngHwnd
        .lpszTitle = lStrCat(strPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Display the browse folder...
    lngIDList = SHBrowseForFolder(udtBI)

    If lngIDList <> 0 Then
        'Create string of nulls so it will fill in with the path
        strPath = String(MAX_PATH, 0)

        'Retrieves the path selected, places in the null
         'character filled string
        lngResult = SHGetPathFromIDList(lngIDList, strPath)

        'Frees memory
        Call CoTaskMemFree(lngIDList)

        'Find the first instance of a null character,
         'so we can get just the path
        intNull = InStr(strPath, vbNullChar)
        'Greater than 0 means the path exists...
        If intNull > 0 Then
            'Set the value
            strPath = Left(strPath, intNull - 1)
        End If
    End If

    'Return the path name
    BrowseForFolder = strPath
    Exit Function 'Abort

ehBrowseForFolder:

    'Return no value
    BrowseForFolder = Empty
 
End Function

Private Function ShowColor() As Long
   
   'set the structure size
   cc.lStructSize = Len(cc)
   'Set the owner
   cc.hwndOwner = Form1.HWND
   'set the application's instance
   cc.hInstance = App.hInstance
   'set the custom colors (converted to Unicode)
   cc.lpCustColors = ""
   'no extra flags
   cc.flags = 0  'set to 0 = define custom colors unselected. 2= define custom colors selected
   
   'Show the 'Select Color'-dialog
   If CHOOSECOLOR(cc) <> 0 Then
      ShowColor = (cc.rgbResult)
   Else
      ShowColor = -1
   End If
   
End Function

Private Sub Check1_Click()
   If Check1.Value = Checked Then
      ucThumbNails1.ShowFolderInfo = True
   Else
      ucThumbNails1.ShowFolderInfo = False
   End If
End Sub

Private Sub Check2_Click()
   If Check2.Value = Checked Then
     ucThumbNails1.RndCorners = True
   Else
     ucThumbNails1.RndCorners = False
   End If
End Sub

Private Sub Check4_Click()
   If Check4.Value = Checked Then
     ucThumbNails1.PicDimen = True
   Else
     ucThumbNails1.PicDimen = False
   End If
End Sub

Private Sub Command1_Click()
   Label2.Caption = ""
   Label3.Caption = ""
   Label6.Caption = ""
   ucThumbNails1.FolderPath = BrowseForFolder(Me.HWND, "Select a Folder")
End Sub

Private Sub Command2_Click()
   ucThumbNails1.FolderPath = Text1.Text
End Sub

Private Sub Form_Load()
   Label8.BackColor = ucThumbNails1.FontColor
   Label9.BackColor = ucThumbNails1.PicBoxBorderColor
   Label1.BackColor = ucThumbNails1.BackColor
   Label13.BackColor = ucThumbNails1.PicBoxBackColor
   Check1.Value = 1
   Check4.Value = 1
End Sub

Private Sub Label1_Click()  'FOR DEMO ONLY
Dim sure As Long
sure = ShowColor
If sure = -1 Then Exit Sub
Label1.BackColor = sure
ucThumbNails1.BackColor = sure
End Sub

Private Sub Label13_Click()   'FOR DEMO ONLY
Dim sure As Long
sure = ShowColor
If sure = -1 Then Exit Sub
Label13.BackColor = sure
ucThumbNails1.PicBoxBackColor = sure
End Sub

Private Sub Label8_Click()   'FOR DEMO ONLY
Dim sure As Long
sure = ShowColor
If sure = -1 Then Exit Sub
Label8.BackColor = sure
ucThumbNails1.FontColor = sure
End Sub

Private Sub Label9_Click()  'FOR DEMO ONLY
Dim sure As Long
sure = ShowColor
If sure = -1 Then Exit Sub
Label9.BackColor = sure
ucThumbNails1.PicBoxBorderColor = sure
End Sub

Private Sub ucThumbNails1_Click()
   Label2.Caption = ucThumbNails1.SelectedFile
   Label3.Caption = ucThumbNails1.FullPath
   Label6.Caption = ucThumbNails1.FolderPath
End Sub

