VERSION 5.00
Begin VB.UserControl ucThumbNails 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   ScaleHeight     =   7770
   ScaleWidth      =   8595
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3630
      Left            =   150
      ScaleHeight     =   3600
      ScaleWidth      =   5520
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   5550
      Begin VB.Image Image1 
         Height          =   1605
         Left            =   1155
         Top             =   1035
         Width           =   1740
      End
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   6330
      TabIndex        =   6
      Top             =   270
      Width           =   1635
   End
   Begin VB.PictureBox picPR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   6120
      TabIndex        =   5
      Top             =   6870
      Width           =   6150
      Begin VB.Shape sPrBar 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H0000C000&
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   300
         Left            =   -120
         Top             =   -45
         Width           =   120
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   0
      ScaleHeight     =   6825
      ScaleWidth      =   6120
      TabIndex        =   0
      Top             =   0
      Width           =   6150
      Begin VB.PictureBox picDisplay 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7665
         Left            =   -15
         ScaleHeight     =   7665
         ScaleWidth      =   5850
         TabIndex        =   2
         Top             =   -15
         Visible         =   0   'False
         Width           =   5850
         Begin VB.PictureBox picThumbNail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1170
            Index           =   0
            Left            =   165
            ScaleHeight     =   1170
            ScaleWidth      =   1185
            TabIndex        =   3
            Top             =   90
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label lblSorry 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Sorry...No pictures available."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   750
            TabIndex        =   18
            Top             =   480
            Visible         =   0   'False
            Width           =   5070
         End
         Begin VB.Label lblTDim 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   75
            TabIndex        =   17
            Top             =   1515
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Image imgPreSize 
            Height          =   330
            Left            =   4050
            Top             =   2115
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Shape sFrame 
            BorderColor     =   &H00000000&
            Height          =   1215
            Index           =   0
            Left            =   135
            Shape           =   4  'Rounded Rectangle
            Top             =   75
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.Label lblThumb 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   75
            TabIndex        =   4
            Top             =   1290
            Visible         =   0   'False
            Width           =   1350
         End
      End
      Begin VB.VScrollBar VS1 
         Enabled         =   0   'False
         Height          =   6975
         Left            =   5850
         Max             =   100
         TabIndex        =   1
         Top             =   0
         Width           =   285
      End
      Begin VB.Label lblLoading 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading...Please Wait"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1350
         TabIndex        =   7
         Top             =   3255
         Visible         =   0   'False
         Width           =   3255
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File Size"
      Height          =   195
      Left            =   15
      TabIndex        =   16
      Top             =   7470
      Width           =   645
   End
   Begin VB.Label lblFileSize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   645
      TabIndex        =   15
      Top             =   7455
      Width           =   930
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Folder Size"
      Height          =   210
      Left            =   2130
      TabIndex        =   13
      Top             =   7470
      Width           =   795
   End
   Begin VB.Label lblFolderSz 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2940
      TabIndex        =   12
      Top             =   7455
      Width           =   1050
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Files"
      Height          =   255
      Left            =   4770
      TabIndex        =   11
      Top             =   7485
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Folder"
      Height          =   195
      Left            =   15
      TabIndex        =   10
      Top             =   7215
      Width           =   1125
   End
   Begin VB.Label lblCurFold 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1065
      TabIndex        =   9
      Top             =   7200
      Width           =   5070
   End
   Begin VB.Label lblFileCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5520
      TabIndex        =   8
      Top             =   7455
      Width           =   615
   End
   Begin VB.Image imgSizing 
      Height          =   1095
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1110
   End
End
Attribute VB_Name = "ucThumbNails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************
'**                          ThumbNail User Control
'**                               Version 1.1.5
'**                               By Ken Foster
'**                                 January 2008
'**                     Freeware--- no copyrights claimed
'*******************************************************************

'THIS CONTROL IS USING GDI+  SO MAKE SURE YOU HAVE IT ON YOUR SYSTEM

'Credit goes to Miltiadis Kritikos for the PNG code

'*******************************************************************

'***************** Table of Procedures *************
'   Private Sub UserControl_Initialize
'   Private Sub UserControl_ReadProperties
'   Private Sub UserControl_Resize
'   Private Sub UserControl_WriteProperties
'   Private Sub ClearAll
'   Public Sub CreateThumbNail
'   Public Sub FormDrag
'   Private Sub lblThumb_Click
'   Private Sub LoadList
'   Private Sub LoadPics
'   Private Sub picPreview_DblClick
'   Private Sub picPreview_MouseDown
'   Private Sub picThumbNail_MouseDown
'   Private Sub SmoothForm
'   Private Sub VS1_Change
'   Private Sub VS1_Scroll
'   Public Function GetPathSize
'   Private Function ShortName
'   Public Property Get BackColor
'   Public Property Let BackColor
'   Public Property Let FolderPath
'   Public Property Get FolderPath
'   Public Property Get FolderSize
'   Public Property Get FontColor
'   Public Property Let FontColor
'   Public Property Get FullPath
'   Public Property Get PicBoxBackColor
'   Public Property Let PicBoxBackColor
'   Public Property Get PicBoxBorderColor
'   Public Property Let PicBoxBorderColor
'   Public Property Get PicDimen
'   Public Property Let PicDimen
'   Public Property Get ProBarColor
'   Public Property Let ProBarColor
'   Public Property Get RndCorners
'   Public Property Let RndCorners
'   Public Property Get SelectedFile
'   Public Property Get SelectedFolder
'   Public Property Let ShowFolderInfo
'   Public Property Get ShowFolderInfo
'   Public Function InitGDIPlus
'   Public Sub FreeGDIPlus
'   Public Function LoadPictureGDIPlus
'   Private Sub InitDC
'   Private Sub gdipResize
'   Private Sub GetBitmap
'   Private Function CreatePicture
'   Public Function Resize
'   Private Sub FillInWmfHeader

'GDI CODE BY MILTIADIS KRITIKOS
'   Public Function InitGDIPlus
'   Public Sub FreeGDIPlus
'   Public Function LoadPictureGDIPlus
'   Private Sub InitDC
'   Private Sub gdipResize
'   Private Sub GetBitmap
'   Private Function CreatePicture
'   Public Function Resize
'   Private Sub FillInWmfHeader
'***************** End of Table ********************


Private Declare Function SetWindowRgn Lib "user32" ( _
      ByVal HWND As Long, _
      ByVal hRgn As Long, _
      ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" ( _
      ByVal X1 As Long, _
      ByVal Y1 As Long, _
      ByVal X2 As Long, _
      ByVal Y2 As Long, _
      ByVal X3 As Long, _
      ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" _
      Alias "SendMessageA" ( _
      ByVal HWND As Long, _
      ByVal wMsg As Long, _
      ByVal wParam As Long, _
      lParam As Any) As Long

Event Click()
Event DblClick()

Private Const m_def_FolderPath = ""
Private Const m_def_SelectedFile = ""
Private Const m_def_SelectedFolder = ""
Private Const m_def_FullPath = ""
Private Const m_def_PicBoxBorderColor = vbBlack
Private Const m_def_FontColor = vbBlack
Private Const m_def_BackColor = vbWhite
Private Const m_def_ShowFolderInfo = True
Private Const m_def_FolderSize = ""
Private Const m_def_PicBoxBackColor = vbWhite
Private Const m_def_ProBarColor = vbGreen
Private Const m_def_RndCorners = False
Private Const m_def_IncludePNGs = True
Private Const m_def_PicDimen = True

Private mvalue As Integer
Private m_FolderPath As String
Private m_SelectedFile As String
Private m_SelectedFolder As String
Private m_FullPath As String
Private m_PicBoxBorderColor As OLE_COLOR
Private m_FontColor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_ShowFolderInfo As Boolean
Private m_FolderSize As String
Private m_PicBoxBackColor As OLE_COLOR
Private m_ProBarColor As OLE_COLOR
Private m_RndCorners As Boolean
Dim m_IncludePNGs As Boolean
Dim m_PicDimen As Boolean

'---------------------For PNG files-------------------
Private Type GUID
   Data1    As Long
   Data2    As Integer
   Data3    As Integer
   Data4(7) As Byte
End Type

Private Type PICTDESC
   size     As Long
   Type     As Long
   hBmp     As Long
   hPal     As Long
   Reserved As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type PWMFRect16
    Left   As Integer
    Top    As Integer
    Right  As Integer
    Bottom As Integer
End Type

Private Type wmfPlaceableFileHeader
    Key         As Long
    hMf         As Integer
    BoundingBox As PWMFRect16
    Inch        As Integer
    Reserved    As Long
    CheckSum    As Integer
End Type

' GDI Functions
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

' GDI+ functions
Private Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal FileName As Long, GpImage As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (token As Long, gdipInput As GdiplusStartupInput, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hDC As Long, GpGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal Img As Long, ByVal x As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hBmp As Long, ByVal hPal As Long, GpBitmap As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipCreateMetafileFromWmf Lib "gdiplus.dll" (ByVal hWmf As Long, ByVal deleteWmf As Long, WmfHeader As wmfPlaceableFileHeader, Metafile As Long) As Long
Private Declare Function GdipCreateMetafileFromEmf Lib "gdiplus.dll" (ByVal hEmf As Long, ByVal deleteEmf As Long, Metafile As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "gdiplus.dll" (ByVal hIcon As Long, GpBitmap As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal GpImage As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal token As Long)

' GDI and GDI+ constants
Private Const PLANES = 14            '  Number of planes
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Const PATCOPY = &HF00021     ' (DWORD) dest = pattern
Private Const PICTYPE_BITMAP = 1     ' Bitmap type
Private Const InterpolationModeHighQualityBicubic = 7
Private Const GDIP_WMF_PLACEABLEKEY = &H9AC6CDD7
Private Const UnitPixel = 2
'----------------------end PNG files----------------------------

Private tmpPath As String
Private prev As Boolean
Dim stg As String
Dim pngwidth As Long
Dim pngheight As Long

Private Sub UserControl_Initialize()

   PicBoxBorderColor = m_def_PicBoxBorderColor
   FontColor = m_def_FontColor
   BackColor = m_def_BackColor
   PicBoxBackColor = m_def_PicBoxBackColor
   ShowFolderInfo = m_def_ShowFolderInfo
   ProBarColor = m_def_ProBarColor
   RndCorners = m_def_RndCorners
   PicDimen = m_def_PicDimen
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   With PropBag
      PicBoxBorderColor = .ReadProperty("PicBoxBorderColor", m_def_PicBoxBorderColor)
      FontColor = .ReadProperty("FontColor", m_def_FontColor)
      BackColor = .ReadProperty("BackColor", m_def_BackColor)
      ShowFolderInfo = .ReadProperty("ShowFolderInfo", m_def_ShowFolderInfo)
      PicBoxBackColor = .ReadProperty("PicBoxBackColor", m_def_PicBoxBackColor)
      ProBarColor = .ReadProperty("ProBarColor", m_def_ProBarColor)
      RndCorners = .ReadProperty("RndCorners", m_def_RndCorners)
      PicDimen = .ReadProperty("PicDimen", m_def_PicDimen)
   End With

End Sub

Private Sub UserControl_Resize()

   If ShowFolderInfo = True Then
      UserControl.Width = picDisplay.Width + VS1.Width + 10
      UserControl.Height = picMain.Height + picPR.Height + lblFolderSz.Height + 400
    Else
      UserControl.Width = picDisplay.Width + VS1.Width + 10
      UserControl.Height = picMain.Height + picPR.Height + 30
   End If
   VS1.Height = picMain.Height
   picDisplay.Top = picMain.Top - 10
   picDisplay.Left = picMain.Left
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "PicBoxBorderColor", m_PicBoxBorderColor, m_def_PicBoxBorderColor
      .WriteProperty "FontColor", m_FontColor, m_def_FontColor
      .WriteProperty "BackColor", m_BackColor, m_def_BackColor
      .WriteProperty "ShowFolderInfo", m_ShowFolderInfo, m_def_ShowFolderInfo
      .WriteProperty "PicBoxBackColor", m_PicBoxBackColor, m_def_PicBoxBackColor
      .WriteProperty "ProBarColor", m_ProBarColor, m_def_ProBarColor
      .WriteProperty "RndCorners", m_RndCorners, m_def_RndCorners
      .WriteProperty "PicDimen", m_PicDimen, m_def_PicDimen
   End With

End Sub

Private Sub ClearAll()

  Dim x As Integer

   'reset display window and scrollbar to top
   picDisplay.Top = 0
   picDisplay.Height = 0
   VS1.Value = 0
   
   'clear all thumbnails and labels
   For x = 0 To List1.ListCount - 1
      picThumbNail(x).Picture = LoadPicture
      imgSizing.Picture = LoadPicture
      picThumbNail(x).Visible = False
      sFrame(x).Visible = False
      lblThumb(x).Caption = ""
      lblTDim(x).Caption = ""
   Next x

   List1.Clear          'clear the files list
   
End Sub

Public Sub CreateThumbNail(PicBox As Object, _
                           ByVal picImg As StdPicture, _
                           ByVal MaxHeight As Long, _
                           ByVal MaxWidth As Long, _
                           Center As Boolean, _
                           Optional ByVal PicTop As Integer, _
                           Optional ByVal PicLeft As Integer)

  Dim NewH As Long                                                            'New Height
  Dim NewW As Long                                                           'New Width

 
      NewH = picImg.Height                                                   'actual image height
      NewW = picImg.Width                                                   'actual image width

   If NewH > MaxHeight Or NewW > MaxWidth Then            'picture is too large
      If NewH > NewW Then                                                   'height is greater than width
         NewW = Fix((NewW / NewH) * MaxHeight)                 'rescale height
         NewH = MaxHeight                                                      'set max height
       ElseIf NewW > NewH Then                                            'width is greater than height
         NewH = Fix((NewH / NewW) * MaxWidth)                   'rescale width
         NewW = MaxWidth                                                      'set max Width
         Debug.Print "Width>"
       Else                                                                                'image is square
         NewH = MaxHeight
         NewW = MaxWidth
      End If

   End If

   'check if centered
   If Center = True Then                                                         'center picture
      PicTop = (PicBox.Height / 2) - (NewH / 2) - 10
      PicLeft = (PicBox.Width / 2) - (NewW / 2) - 10
    Else                                                                                    'if Optional variables are
                                                                                              'missing  and center=false
      PicTop = 0                                                                        'Default top position
      PicLeft = 0                                                                        'Default left position
   End If

   If prev = True Then
      picPreview.Width = NewW
      picPreview.Height = NewH
      prev = False
   End If

   'Draw newly scaled picture
   With PicBox
      .AutoRedraw = True                                                            'set needed properties
      .Cls                                                                                      'clear picture box
      .PaintPicture picImg, PicLeft, PicTop, NewW, NewH             'paint new picture size in                                                                    '   picturebox
   End With

End Sub

Public Sub FormDrag(TheForm As Object)

   On Local Error Resume Next
   ReleaseCapture
   SendMessage TheForm.HWND, &HA1, 2, 0&

End Sub

Private Sub lblThumb_Click(Index As Integer)

   m_SelectedFile = List1.List(Index)
   m_SelectedFolder = tmpPath
   m_FullPath = tmpPath & List1.List(Index)
   lblFileSize.Caption = FormatNumber(FileLen(m_FullPath) / 1024, 0) & "KB"
   RaiseEvent Click
   RaiseEvent DblClick

End Sub

Private Sub LoadList()

  Dim sTemp As String

   If Right$(FolderPath, 1) = "\" Then                             'determine weather theres a "\"
                                                                                      'at the end of the path
      tmpPath = FolderPath
    Else
      tmpPath = FolderPath & "\"
   End If

   'load the filenames into the list
   sTemp = Dir$(tmpPath & "*.*")

   While sTemp <> ""
      List1.AddItem sTemp
      sTemp = Dir$
   Wend

   lblFileCount.Caption = List1.ListCount                         'show number of files
   lblCurFold.Caption = FolderPath
   lblFolderSz.Caption = FormatNumber(GetPathSize(FolderPath) / 1024, 0) & "KB"
   Image1.Visible = False                                               'just in case it was still visible
End Sub

Private Sub LoadPics()

  Dim x As Integer
  Dim Y As Integer
  Dim token As Long
  Dim pgct As Single
  
   On Error Resume Next
   picPreview.Visible = False
   picDisplay.Visible = False
   lblSorry.Visible = False
   lblLoading.Visible = True                                   'Loading...Please Wait
   DoEvents

   VS1.Enabled = False                                         'disable scrollbar while loading
   
   'calculate height of display window
   pgct = List1.ListCount / 16                              '16 is the number of images displayed per page
   If pgct <= 1 Then
      picDisplay.Height = picMain.Height
      VS1.Max = 0
   Else
      picDisplay.Height = (pgct * (picMain.Height * 1.1))
      VS1.Max = picDisplay.Height / picMain.Height * 2
   End If
   
   VS1_Change
   
   For x = 0 To List1.ListCount - 1
      'dynamically load all controls
      Load picThumbNail(x)
      Load lblThumb(x)
      Load lblTDim(x)
      Load sFrame(x)

      If RndCorners = True Then
         SmoothForm picThumbNail(x), 15
         sFrame(x).Shape = 4
       Else
         sFrame(x).Shape = 0
      End If

      picThumbNail(x).BackColor = PicBoxBackColor

      With picThumbNail(x)
         .Left = picThumbNail(x - 1).Left + picThumbNail(x - 1).Width + 245
         .Top = picThumbNail(x - 1).Top
         .Visible = True
         .Picture = LoadPicture()
      End With

      With lblThumb(x)
         .Left = picThumbNail(x).Left - 100
         .Top = picThumbNail(x).Top + picThumbNail(x).Height + 20
         .ForeColor = FontColor
         .Caption = ShortName(List1.List(x))
         .Visible = True
      End With
      
      With lblTDim(x)
         .Left = picThumbNail(x).Left - 100
         .Top = lblThumb(x).Top + lblThumb(x).Height - 40
         .ForeColor = FontColor
         .Caption = ""
         .Visible = True
      End With
      With sFrame(x)
         .Left = picThumbNail(x).Left - 30
         .Top = picThumbNail(x).Top - 30
         .BorderColor = PicBoxBorderColor
         .Visible = True
      End With

      If Y = 4 Then                                                 'allow 4 thumbnails per row
         Y = 0
         'position the first picture and label in new row

         With picThumbNail(x)
            .Left = picThumbNail(0).Left
            .Top = picThumbNail(x - 1).Top + lblThumb(x - 1).Height + picThumbNail(x - 1).Height + _
               300
         End With

         With lblThumb(x)
            .Left = picThumbNail(x).Left - 100
            .Top = picThumbNail(x).Top + picThumbNail(x).Height + 20
         End With
         
         With lblTDim(x)
            .Left = lblThumb(x).Left
            .Top = lblThumb(x).Top + lblThumb(x).Height - 40
         End With
         
         With sFrame(x)
            .Left = picThumbNail(x).Left - 30
            .Top = picThumbNail(x).Top - 30
         End With

      End If

      Y = Y + 1                                                                  'advance the row counter by one
                                                                       
      'for checking extension
      stg = LCase(Right$(List1.List(x), 4))                       'change any upper case extensions to lower case
      
      If stg = ".bmp" Or stg = ".jpg" Or stg = "jpeg" Or stg = ".ico" Or stg = ".gif" Or stg = ".wmf" Or stg = _
         ".avi" Or stg = ".png" Then                                  'excepted formats ... process
         
         If stg = ".png" = True Then
            token = InitGDIPlus
            imgSizing.Picture = LoadPictureGDIPlus(FolderPath & "\" & List1.List(x), , , vbWhite)
            CreateThumbNail picThumbNail(x), imgSizing.Picture, picThumbNail(x).Width, picThumbNail(x).Height, True
            FreeGDIPlus token
               If PicDimen = True Then                                   'get picture dimensions
                  lblTDim(x).Caption = pngwidth & " X " & pngheight
               End If
         Else
            imgSizing.Picture = LoadPicture(FolderPath & "\" & List1.List(x))
            CreateThumbNail picThumbNail(x), imgSizing.Picture, picThumbNail(x).Width, picThumbNail(x).Height, True
               If PicDimen = True Then                                      'get picture dimensions
                   lblTDim(x).Caption = imgSizing.Width / Screen.TwipsPerPixelX & " X " & imgSizing.Height / Screen.TwipsPerPixelY
               End If
         End If
         
      Else
         'process non-accepted extensions
         picThumbNail(x).ForeColor = FontColor
         picThumbNail(x).CurrentX = 200
         picThumbNail(x).CurrentY = 400
         picThumbNail(x).Print "Format Not"
         picThumbNail(x).CurrentX = 200
         picThumbNail(x).CurrentY = 600
         picThumbNail(x).Print "Supported"
      End If

      'calculate the length of progress bar
      mvalue = mvalue + 1
      sPrBar.Width = ((picPR.Width / List1.ListCount) * mvalue) + 500
      
   Next x
   picDisplay.Visible = True
   
   'if no pics in folder , show label
   If List1.ListCount = 0 Then
      lblSorry.Visible = True
    Else
      lblSorry.Visible = False
   End If

   lblLoading.Visible = False
   picDisplay.Top = -10
   If picDisplay.Visible = True Then VS1.Enabled = True

   'reset progress bar to zero
   sPrBar.Width = 0
   mvalue = 0

End Sub

Private Sub picPreview_DblClick()
   picPreview.Visible = False
   lblFileSize.Caption = ""
End Sub

Private Sub picPreview_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   FormDrag picPreview
End Sub

Private Sub picThumbNail_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim token As Long

   'show preview picture
   If Button = 1 Then
      stg = LCase(Right$(List1.List(Index), 4))
      m_SelectedFile = List1.List(Index)
      m_SelectedFolder = tmpPath
      m_FullPath = tmpPath & List1.List(Index)
      lblFileSize.Caption = FormatNumber(FileLen(m_FullPath) / 1024, 0) & "KB"
    
      If stg <> ".png" Then
          imgPreSize.Picture = LoadPicture(tmpPath & List1.List(Index))
          prev = True
          picPreview.Visible = True
          CreateThumbNail picPreview, imgPreSize.Picture, 5000, 5000, False
       Else
          token = InitGDIPlus
          
          prev = True
          picPreview.Visible = True
             If (FileLen(tmpPath & List1.List(Index)) / 1024) > 3 Then
                Image1.Visible = False
                picPreview.Width = 5550
                picPreview.Height = 3630
                picPreview.Picture = LoadPictureGDIPlus(tmpPath & List1.List(Index), picPreview.Width / Screen.TwipsPerPixelX, picPreview.Height / Screen.TwipsPerPixelY, vbWhite, True)
             Else
                picPreview.Width = 3000
                picPreview.Height = 3000
                picPreview.Picture = LoadPicture()
                Image1.Visible = True                                       'I'm using image1 to center small png's in preview window
                Image1.Picture = LoadPictureGDIPlus(tmpPath & List1.List(Index), picPreview.Width / Screen.TwipsPerPixelX / 4, picPreview.Height / Screen.TwipsPerPixelY / 4, vbWhite, True)
             End If
       FreeGDIPlus token
       End If
   End If

   RaiseEvent Click
   RaiseEvent DblClick

End Sub

Private Sub SmoothForm(Frm As PictureBox, Optional ByVal Curvature As Double = 25)

  Dim hRgn As Long
  Dim X1 As Long
  Dim Y1 As Long

   X1 = Frm.Width / Screen.TwipsPerPixelX
   Y1 = Frm.Height / Screen.TwipsPerPixelY
   hRgn = CreateRoundRectRgn(0, 0, X1, Y1, Curvature, Curvature)
   SetWindowRgn Frm.HWND, hRgn, True
   DeleteObject hRgn

End Sub

Private Sub VS1_Change()
   
   If picDisplay.Height < picMain.Height Then picDisplay.Height = picMain.Height
   picDisplay.Top = picMain.Top - ((picDisplay.Height / (VS1.Max + 3) * VS1.Value))
   picDisplay.SetFocus                                   'keeps bar from flashing all the time

End Sub

Private Sub VS1_Scroll()
   VS1_Change
End Sub

Public Function GetPathSize(ByRef sPathName As String) As Double

  Dim sFileName As String
  Dim dSize As Double
  Dim asFileName() As String
  Dim i As Long

   On Error Resume Next
   ' Enumerate DirNames and FileNames
   If StrComp(Right$(sPathName, 1), "\", vbBinaryCompare) <> 0 Then sPathName = sPathName & "\"
   sFileName = Dir$(sPathName, vbDirectory + vbHidden + vbSystem + vbReadOnly)

   Do While Len(sFileName) > 0

      If StrComp(sFileName, ".", vbBinaryCompare) <> 0 And StrComp(sFileName, "..", _
         vbBinaryCompare) <> 0 Then
         ReDim Preserve asFileName(i)
         asFileName(i) = sPathName & sFileName
         i = i + 1
      End If

      sFileName = Dir
   Loop

   If i > 0 Then

      For i = 0 To UBound(asFileName)

         If (GetAttr(asFileName(i)) And vbDirectory) = vbDirectory Then
            ' Add dir size
            dSize = dSize + GetPathSize(asFileName(i))
          Else
            ' Add file size
            dSize = dSize + FileLen(asFileName(i))
         End If

      Next
   End If

   GetPathSize = dSize

End Function

Private Function ShortName(sName As String)

  Dim LName As String
  Dim fName As String
  Dim RName As String

   If Len(sName) >= 14 Then                                        'this seems to be a good
                                                                                    'compromise on name length
      LName = Left$(sName, 7)
      fName = "---"
      RName = Right$(sName, 4)
      ShortName = LName & fName & RName
    Else
      ShortName = sName
   End If

End Function

Public Property Get BackColor() As OLE_COLOR
   BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
   m_BackColor = NewBackColor
   picDisplay.BackColor = m_BackColor
   picMain.BackColor = m_BackColor
   UserControl.BackColor = m_BackColor
   PropertyChanged "BackColor"
End Property

Public Property Let FolderPath(ByVal NewFolderPath As String)
   m_FolderPath = NewFolderPath
   If NewFolderPath = "" Then Exit Property
   ClearAll
   tmpPath = m_FolderPath
   LoadList
   LoadPics
   PropertyChanged "FolderPath"
End Property

Public Property Get FolderPath() As String
   FolderPath = m_FolderPath
End Property

Public Property Get FolderSize() As String
   FolderSize = m_FolderSize
End Property

Public Property Get FontColor() As OLE_COLOR
   Let FontColor = m_FontColor
End Property

Public Property Let FontColor(ByVal NewFontColor As OLE_COLOR)
   m_FontColor = NewFontColor
   Label1.ForeColor = FontColor
   Label3.ForeColor = FontColor
   Label4.ForeColor = FontColor
   Label6.ForeColor = FontColor
   FontCC
   PropertyChanged "FontColor"
End Property

Public Property Get FullPath() As String
   FullPath = m_FullPath
End Property

Public Property Get PicBoxBackColor() As OLE_COLOR
   PicBoxBackColor = m_PicBoxBackColor
End Property

Public Property Let PicBoxBackColor(ByVal NewPicBoxBackColor As OLE_COLOR)
   m_PicBoxBackColor = NewPicBoxBackColor
   PropertyChanged "PicBoxBackColor"
End Property

Public Property Get PicBoxBorderColor() As OLE_COLOR
   PicBoxBorderColor = m_PicBoxBorderColor
End Property

Public Property Let PicBoxBorderColor(ByVal NewPicBoxBorderColor As OLE_COLOR)
   m_PicBoxBorderColor = NewPicBoxBorderColor
   BorderCC
   PropertyChanged "PicBoxBorderColor"
End Property

Public Property Get PicDimen() As Boolean
   Let PicDimen = m_PicDimen
End Property

Public Property Let PicDimen(ByVal NewPicDimen As Boolean)
   Let m_PicDimen = NewPicDimen
   PDV
   PropertyChanged "PicDimen"
End Property

Public Property Get ProBarColor() As OLE_COLOR
   ProBarColor = m_ProBarColor
End Property

Public Property Let ProBarColor(ByVal NewProBarColor As OLE_COLOR)
   m_ProBarColor = NewProBarColor
   sPrBar.BorderColor = m_ProBarColor
   sPrBar.FillColor = m_ProBarColor
   PropertyChanged "ProBarColor"
End Property

Public Property Get RndCorners() As Boolean
   RndCorners = m_RndCorners
End Property

Public Property Let RndCorners(ByVal NewRndCorners As Boolean)
   m_RndCorners = NewRndCorners
   RC
   PropertyChanged "RndCorners"
End Property

Public Property Get SelectedFile() As String
   SelectedFile = m_SelectedFile
End Property

Public Property Get SelectedFolder() As String
   SelectedFolder = m_SelectedFolder
End Property

Public Property Let ShowFolderInfo(ByVal NewShowFolderInfo As Boolean)
   m_ShowFolderInfo = NewShowFolderInfo
   UserControl_Resize
   PropertyChanged "ShowFolderInfo"
End Property

Public Property Get ShowFolderInfo() As Boolean
   ShowFolderInfo = m_ShowFolderInfo
End Property

'------------------------ start of GDI+ for the PNG files-----------------------------------
Public Function InitGDIPlus() As Long
    Dim token    As Long
    Dim gdipInit As GdiplusStartupInput
    
    gdipInit.GdiplusVersion = 1
    GdiplusStartup token, gdipInit, ByVal 0&
    InitGDIPlus = token
End Function

' Frees GDI Plus
Public Sub FreeGDIPlus(token As Long)
    GdiplusShutdown token
End Sub

' Loads the picture (optionally resized)
Public Function LoadPictureGDIPlus(PicFile As String, Optional Width As Long = -1, Optional Height As Long = -1, Optional ByVal BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
    Dim hDC     As Long
    Dim hBitmap As Long
    Dim Img     As Long
        
    ' Load the image
    If GdipLoadImageFromFile(StrPtr(PicFile), Img) <> 0 Then
        Err.Raise 999, "GDI+ Module", "Error loading picture " & PicFile
        Exit Function
    End If
    
    ' Calculate picture's width and height if not specified
    If Width = -1 Or Height = -1 Then
        GdipGetImageWidth Img, Width
        pngwidth = Width
        GdipGetImageHeight Img, Height
        pngheight = Height
    End If
    
    ' Initialise the hDC
    InitDC hDC, hBitmap, BackColor, Width, Height

    ' Resize the picture
    gdipResize Img, hDC, Width, Height, RetainRatio
    GdipDisposeImage Img
    
    ' Get the bitmap back
    GetBitmap hDC, hBitmap

    ' Create the picture
    Set LoadPictureGDIPlus = CreatePicture(hBitmap)
End Function

' Initialises the hDC to draw
Private Sub InitDC(hDC As Long, hBitmap As Long, BackColor As Long, Width As Long, Height As Long)
    Dim hBrush As Long
        
    ' Create a memory DC and select a bitmap into it, fill it in with the backcolor
    hDC = CreateCompatibleDC(ByVal 0&)
    hBitmap = CreateBitmap(Width, Height, GetDeviceCaps(hDC, PLANES), GetDeviceCaps(hDC, BITSPIXEL), ByVal 0&)
    hBitmap = SelectObject(hDC, hBitmap)
    hBrush = CreateSolidBrush(BackColor)
    hBrush = SelectObject(hDC, hBrush)
    PatBlt hDC, 0, 0, Width, Height, PATCOPY
    DeleteObject SelectObject(hDC, hBrush)
End Sub

' Resize the picture using GDI plus
Private Sub gdipResize(Img As Long, hDC As Long, Width As Long, Height As Long, Optional RetainRatio As Boolean = False)
    Dim Graphics   As Long      ' Graphics Object Pointer
    Dim OrWidth    As Long      ' Original Image Width
    Dim OrHeight   As Long      ' Original Image Height
    Dim OrRatio    As Double    ' Original Image Ratio
    Dim DesRatio   As Double    ' Destination rect Ratio
    Dim DestX      As Long      ' Destination image X
    Dim DestY      As Long      ' Destination image Y
    Dim DestWidth  As Long      ' Destination image Width
    Dim DestHeight As Long      ' Destination image Height
    
    GdipCreateFromHDC hDC, Graphics
    GdipSetInterpolationMode Graphics, InterpolationModeHighQualityBicubic
    
    If RetainRatio Then
        GdipGetImageWidth Img, OrWidth
        GdipGetImageHeight Img, OrHeight
        
        OrRatio = OrWidth / OrHeight
        DesRatio = Width / Height
        
        ' Calculate destination coordinates
        DestWidth = IIf(DesRatio < OrRatio, Width, Height * OrRatio)
        DestHeight = IIf(DesRatio < OrRatio, Width / OrRatio, Height)
        DestX = (Width - DestWidth) / 2
        DestY = (Height - DestHeight) / 2

        GdipDrawImageRectRectI Graphics, Img, DestX, DestY, DestWidth, DestHeight, 0, 0, OrWidth, OrHeight, UnitPixel, 0, 0, 0
    Else
        GdipDrawImageRectI Graphics, Img, 0, 0, Width, Height
    End If
    GdipDeleteGraphics Graphics
End Sub

' Replaces the old bitmap of the hDC, Returns the bitmap and Deletes the hDC
Private Sub GetBitmap(hDC As Long, hBitmap As Long)
    hBitmap = SelectObject(hDC, hBitmap)
    DeleteDC hDC
End Sub

' Creates a Picture Object from a handle to a bitmap
Private Function CreatePicture(hBitmap As Long) As IPicture
    Dim IID_IDispatch As GUID
    Dim Pic           As PICTDESC
    Dim IPic          As IPicture
    
    ' Fill in OLE IDispatch Interface ID
    IID_IDispatch.Data1 = &H20400
    IID_IDispatch.Data4(0) = &HC0
    IID_IDispatch.Data4(7) = &H46
        
    ' Fill Pic with necessary parts
    Pic.size = Len(Pic)        ' Length of structure
    Pic.Type = PICTYPE_BITMAP  ' Type of Picture (bitmap)
    Pic.hBmp = hBitmap         ' Handle to bitmap

    ' Create the picture
    OleCreatePictureIndirect Pic, IID_IDispatch, True, IPic
    Set CreatePicture = IPic
End Function

' Returns a resized version of the picture
Public Function Resize(Handle As Long, PicType As PictureTypeConstants, Width As Long, Height As Long, Optional BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
    Dim Img       As Long
    Dim hDC       As Long
    Dim hBitmap   As Long
    Dim WmfHeader As wmfPlaceableFileHeader
    
    ' Determine pictyre type
    Select Case PicType
    Case vbPicTypeBitmap
         GdipCreateBitmapFromHBITMAP Handle, ByVal 0&, Img
    Case vbPicTypeMetafile
         FillInWmfHeader WmfHeader, Width, Height
         GdipCreateMetafileFromWmf Handle, False, WmfHeader, Img
    Case vbPicTypeEMetafile
         GdipCreateMetafileFromEmf Handle, False, Img
    Case vbPicTypeIcon
         ' Does not return a valid Image object
         GdipCreateBitmapFromHICON Handle, Img
    End Select
    
    ' Continue with resizing only if we have a valid image object
    If Img Then
        InitDC hDC, hBitmap, BackColor, Width, Height
        gdipResize Img, hDC, Width, Height, RetainRatio
        GdipDisposeImage Img
        GetBitmap hDC, hBitmap
        Set Resize = CreatePicture(hBitmap)
    End If
End Function

' Fills in the wmfPlacable header
Private Sub FillInWmfHeader(WmfHeader As wmfPlaceableFileHeader, Width As Long, Height As Long)
    WmfHeader.BoundingBox.Right = Width
    WmfHeader.BoundingBox.Bottom = Height
    WmfHeader.Inch = 1440
    WmfHeader.Key = GDIP_WMF_PLACEABLEKEY
End Sub

Private Sub FontCC()
   Dim x As Integer
   
   For x = 0 To List1.ListCount - 1
      With lblThumb(x)
         .ForeColor = FontColor
         .Caption = ShortName(List1.List(x))
      End With
      
      With lblTDim(x)
         .ForeColor = FontColor
      End With
      Next x
End Sub

Private Sub BorderCC()
   Dim x As Integer
   
   For x = 0 To List1.ListCount - 1
      With sFrame(x)
         .BorderColor = PicBoxBorderColor
      End With
   Next x
End Sub

Private Sub RC()
   Dim x As Integer

   For x = 0 To List1.ListCount - 1
      If RndCorners = True Then
         SmoothForm picThumbNail(x), 15
         sFrame(x).Shape = 4
      Else
         sFrame(x).Shape = 0
      End If
   Next x
End Sub

Private Sub PDV()
   Dim x As Integer

   For x = 0 To List1.ListCount - 1
      If PicDimen = True Then
         lblTDim(x).Visible = True
      Else
         lblTDim(x).Visible = False
      End If
   Next x
End Sub
