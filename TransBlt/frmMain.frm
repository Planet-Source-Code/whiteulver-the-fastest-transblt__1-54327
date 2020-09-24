VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trans Blt"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   2400
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1455
      Left            =   105
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   2
      Top             =   90
      Width           =   1050
   End
   Begin VB.PictureBox picDest 
      AutoRedraw      =   -1  'True
      Height          =   1455
      Left            =   1185
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   1
      Top             =   90
      Width           =   1050
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "TransBlt"
      Height          =   495
      Left            =   555
      TabIndex        =   0
      Top             =   1635
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------
'Fast Transblt by Whiteulver
'-----------------------------


Private Sub cmdTrans_Click()

    Dim FileNo As Integer
    Dim ScreenDC As Long
    Dim FileHead As BITMAPFILEHEADER
    Dim bmpInfo_8 As BITMAPINFO_8
    Dim bmpInfo_8M As BITMAPINFO_8
    Dim bmpInfoHead As BITMAPINFOHEADER
    Dim PicBytes() As Byte
    Dim RetVal As Long
    Dim PicOffset As Long, PicLen As Long
    Dim PicPrevBmp As Long, PicBmp As Long
    Dim PicPrevBmpM As Long, PicBmp_M As Long
    Dim filename As String
    
    filename = App.Path & "\test.bmp"
    'Loads a filename into a picturebox
    picSource.Picture = LoadPicture(filename)

    picSource.Refresh
    'Loads a filename into a picture in memory.
    
    'Read in the file data
    FileNo = FreeFile
    Open filename For Binary Access Read As #FileNo
    'Retrieve the bitmap information
    Get #FileNo, , FileHead
    Get #FileNo, 15, bmpInfoHead
    
    PicLen = FileHead.bfSize
    
    'Load the picture into the memory bitmap
    Select Case bmpInfoHead.biBitCount
        Case Is = 8
            Get #FileNo, 15, bmpInfo_8
        Case Is = 24
            MsgBox "Only 8-bit bitmap support!"
            Exit Sub
    End Select
    'Copy bitmap palette to mask bitmap palette
    bmpInfo_8M = bmpInfo_8
    
    'Convert Trans color to black at main bitmap
    With bmpInfo_8.bmiColors(0)
        .rgbRed = 0
        .rgbGreen = 0
        .rgbBlue = 0
    End With
    'Convert Trans color to white at mask bitmap
    With bmpInfo_8M.bmiColors(0)
        .rgbRed = 255
        .rgbGreen = 255
        .rgbBlue = 255
    End With

    ReDim PicBytes(0 To PicLen - 44)
    
    Get #FileNo, , PicBytes
    
    'Create the memory bitmap
    ScreenDC = GetDC(0)
    PicDC = CreateCompatibleDC(ScreenDC)
    PicDC_M = CreateCompatibleDC(ScreenDC)
    picWidth = bmpInfoHead.biWidth
    picHeight = bmpInfoHead.biHeight
    PicBmp = CreateCompatibleBitmap(ScreenDC, picWidth, picHeight)
    PicBmp_M = CreateCompatibleBitmap(ScreenDC, picWidth, picHeight)
    PicPrevBmp = SelectObject(PicDC, PicBmp)
    PicPrevBmpM = SelectObject(PicDC_M, PicBmp_M)
    
    'Load the picture into the memory bitmap
    Select Case bmpInfoHead.biBitCount
        Case Is = 8
            RetVal = SetDIBits_8(PicDC, PicBmp, 0, picHeight, PicBytes(0), bmpInfo_8, DIB_RGB_COLORS)
            RetVal = SetDIBits_8(PicDC_M, PicBmp_M, 0, picHeight, PicBytes(0), bmpInfo_8M, DIB_RGB_COLORS)
    End Select
    'Clean up the bitmap
    RetVal = ReleaseDC(0, ScreenDC)
    
    
    'Close the file
    Close #FileNo
    
    DeleteObject PicBmp
    DeleteObject PicBmp_M
    
    'Draw bitmap
    BitBlt picDest.hdc, 0, 0, picWidth, picHeight, PicDC_M, 0, 0, vbSrcAnd
    BitBlt picDest.hdc, 0, 0, picWidth, picHeight, PicDC, 0, 0, vbSrcPaint
    picDest.Refresh
    
End Sub

