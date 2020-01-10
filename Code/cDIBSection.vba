
Option Compare Database
Option Explicit


Private Type RECT
    Left As Long
    top As Long
    right As Long
    Bottom As Long
End Type

Private Type SIZEL
    cx As Long
    cy As Long
End Type

Private Type ENHMETAHEADER
        iType As Long
        nSize As Long
        rclBounds As RECT
        rclFrame As RECT
        dSignature As Long
        nVersion As Long
        nBytes As Long
        nRecords As Long
        nHandles As Integer
        sReserved As Integer
        nDescription As Long
        offDescription As Long
        nPalEntries As Long
        szlDevice As SIZEL
        szlMillimeters As SIZEL
End Type


Private Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgblReterved As Byte
End Type


'Private Enum ERGBCompression
 Private Const BI_RGB = 0&
  Private Const BI_RLE4 = 2&
  Private Const BI_RLE8 = 1&
  Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
'End Enum


Private Type BITMAPINFOHEADER '40 bytes
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long 'ERGBCompression
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type


Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type


Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Private Type DIBSECTION
    dsBm As BITMAP
    dsBmih As BITMAPINFOHEADER
    dsBitfields(2) As Long
    dshSection As Long
    dsOffset As Long
End Type

' From winuser.h
Private Const IMAGE_BITMAP = 0
Private Const IMAGE_ICON = 1
Private Const IMAGE_CURSOR = 2
Private Const IMAGE_ENHMETAFILE = 3

Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000

Private Const vbSrcCopy = &HCC0020
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const WHITENESS = &HFF0062 ' (DWORD) dest = WHITE
Private Const BLACKNESS = &H42 ' (DWORD) dest = BLACK

' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
'Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
'Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBmp As Long, ByVal uStartScan As Long, ByVal cScanLines As Long, ByVal lpvBits As Long, ByRef lpbi As BITMAPINFO, ByVal uUsage As Long) As Long
'Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInstance As Long, ByVal Name As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
'Private Declare Function apiGetObject Lib "gdi32" Alias "GetObjectA" _
'(ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'Private Declare Sub apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'(Destination As Any, Source As Any, ByVal Length As Long)
'
'Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'
'
'Private Declare Function apiGetDC Lib "user32" _
'  Alias "GetDC" (ByVal hwnd As Long) As Long
'
'Private Declare Function apiReleaseDC Lib "user32" _
'  Alias "ReleaseDC" (ByVal hwnd As Long, _
'  ByVal hDC As Long) As Long
'
'
'Private Declare Function apiGetDeviceCaps Lib "gdi32" _
'Alias "GetDeviceCaps" (ByVal hDC As Long, ByVal nIndex As Long) As Long

' Create an Information Context
'Private Declare Function apiCreateIC Lib "gdi32" Alias "CreateICA" _
'(ByVal lpDriverName As String, ByVal lpDeviceName As String, _
'ByVal lpOutput As String, lpInitData As Any) As Long
'
'Private Declare Function apiPlayEnhMetaFile Lib "gdi32" Alias "PlayEnhMetaFile" (ByVal hDC As Long, ByVal hEMF As Long, lpRect As RECT) As Long
'
'
'Private Declare Function apiDeleteEnhMetaFile Lib "gdi32" Alias "DeleteEnhMetaFile" _
'(ByVal hEMF As Long) As Long
'
'Private Declare Function apiCloseEnhMetaFile Lib "gdi32" Alias "CloseEnhMetaFile" _
'(ByVal hDC As Long) As Long
'
'Private Declare Function GetEnhMetaFileHeader Lib "gdi32" _
'(ByVal hEMF As Long, ByVal cbBuffer As Long, lpemh As ENHMETAHEADER) As Long
'
'Private Declare Function apiDeleteDC Lib "gdi32" _
'  Alias "DeleteDC" (ByVal hDC As Long) As Long
'
'Private Declare Function apiCreateSolidBrush Lib "gdi32" Alias "CreateSolidBrush" _
'    (ByVal crColor As Long) As Long
'
'Private Declare Function apiFillRect Lib "user32" Alias "FillRect" _
'(ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long


' Predefined Clipboard Formats
Private Const CF_TEXT = 1
Private Const CF_BITMAP = 2
Private Const CF_METAFILEPICT = 3
Private Const CF_SYLK = 4
Private Const CF_DIF = 5
Private Const CF_TIFF = 6
Private Const CF_OEMTEXT = 7
Private Const CF_DIB = 8
Private Const CF_PALETTE = 9
Private Const CF_PENDATA = 10
Private Const CF_RIFF = 11
Private Const CF_WAVE = 12
Private Const CF_UNICODETEXT = 13
Private Const CF_ENHMETAFILE = 14

'  Device Parameters for GetDeviceCaps()
Private Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

' Handle to the current DIBSection:
Private m_hDib As Long
' Handle to the old bitmap in the DC, for clear up:
Private m_hBmpOld As Long
' Handle to the Device context holding the DIBSection:
Private m_hDC As Long
' Address of memory pointing to the DIBSection's bits:
Private m_lPtr As Long
' Type containing the Bitmap information:
Private m_bmi As BITMAPINFO
' Holds current JPEG's FileName
Private m_CurrentJpegFileName As String
' Array to hold original compressed Jpeg
' to be used for BLOB storage in Table
Private bArray() As Byte

' Temp var
Dim lngRet As Long



Public Function CreateDIB( _
  ByVal lhdc As Long, _
  ByVal lWidth As Long, _
  ByVal lHeight As Long, _
  ByVal lChannels As Long, _
  ByRef hDib As Long _
  ) As Boolean
   
  With m_bmi.bmiHeader
    .biSize = Len(m_bmi.bmiHeader)
    .biWidth = lWidth
    .biHeight = lHeight
    .biPlanes = 1
    ' April 4/2003
    '.biBitCount = lChannels * 8
'    If lChannels = 3 Then
'      .biBitCount = 24
'    Else
'      .biBitCount = 24 '32
'    End If
     .biBitCount = 24
    .biCompression = BI_RGB
    .biSizeImage = BytesPerScanLine * .biHeight
  End With
  
  'The m_lPtr is passed in byref.. so that it returns the the pointer to the bitmapinfo bits
  'the m_lptr is then stored as a reference to the uncompressed image data
  'the m_lptr is filled with image data when the ijlread method is invoked.
  hDib = CreateDIBSection(lhdc, m_bmi, DIB_RGB_COLORS, m_lPtr, 0, 0)
  
  CreateDIB = (hDib <> 0)

End Function


Public Function Create(ByVal lWidth As Long, ByVal lHeight As Long, Optional ByVal lChannels As Long = 3) As Boolean
  
  CleanUp
  
  m_hDC = CreateCompatibleDC(0)
  
  If (m_hDC <> 0) Then
    If (CreateDIB(m_hDC, lWidth, lHeight, lChannels, m_hDib)) Then
      m_hBmpOld = SelectObject(m_hDC, m_hDib)
      Create = True
    Else
      Call DeleteObject(m_hDC)
      m_hDC = 0
    End If
  End If

End Function


Public Function Load(ByVal Name As String) As Boolean
  Dim hBmp As Long
  Dim pName As Long
  Dim aName As String

  Load = False

  CleanUp

  m_hDC = CreateCompatibleDC(0)
  If m_hDC = 0 Then
    Exit Function
  End If

  aName = StrConv(Name, vbFromUnicode)
  pName = StrPtr(aName)

  hBmp = LoadImage(0, pName, IMAGE_BITMAP, 0, 0, (LR_CREATEDIBSECTION Or LR_LOADFROMFILE))
  If hBmp = 0 Then
    Call DeleteObject(m_hDC)
    m_hDC = 0
    MsgBox "Can't load BMP image"
    Exit Function
  End If

  m_bmi.bmiHeader.biSize = Len(m_bmi.bmiHeader)

  ' get image sizes
  Call GetDIBits(m_hDC, hBmp, 0, 0, 0, m_bmi, DIB_RGB_COLORS)

  ' make 24 bpp dib section
  m_bmi.bmiHeader.biBitCount = 24
  m_bmi.bmiHeader.biCompression = BI_RGB
  m_bmi.bmiHeader.biClrUsed = 0
  m_bmi.bmiHeader.biClrImportant = 0
  
  m_hDib = CreateDIBSection(m_hDC, m_bmi, DIB_RGB_COLORS, m_lPtr, 0, 0)
  If m_hDib = 0 Then
    Call DeleteObject(hBmp)
    Call DeleteObject(m_hDC)
    m_hDC = 0
    Exit Function
  End If

  m_hBmpOld = SelectObject(m_hDC, m_hDib)

  m_bmi.bmiHeader.biSize = Len(m_bmi.bmiHeader)

  ' get image data in 24 bpp format (convert if need)
  Call GetDIBits(m_hDC, hBmp, 0, m_bmi.bmiHeader.biHeight, m_lPtr, m_bmi, DIB_RGB_COLORS)

  Call DeleteObject(hBmp)

  Load = True

End Function

Public Function PictureDataToDIB(ctl As Control) As Boolean
 Dim hBmp As Long
  Dim pName As Long
  Dim aName As String

Dim hDCtemp As Long
Dim lngIC As Long

Dim hBMPtemp As Long
Dim lImageType As Long

' Instance of EMF Header structure
 Dim mh As ENHMETAHEADER
 
' Current Screen Resolution
Dim lngXdpi As Long

' Used to convert Metafile dimensions to pixels
Dim sngConvertX As Single
Dim sngConvertY As Single
Dim sngMetaResolutionX As Single
Dim sngMetaResolutionY As Single

Dim rc As RECT
' Init our vars
  CleanUp

  m_hDC = CreateCompatibleDC(0)
  hDCtemp = CreateCompatibleDC(0)
  
  If m_hDC = 0 Then
    Exit Function
  End If


  'lngRet = FPictureDataToClipBoard(ctl)
  'hBmp = GetClipBoard(lImageType)

  'hBmp = LoadImage(0, pName, IMAGE_BITMAP, 0, 0, (LR_CREATEDIBSECTION Or LR_LOADFROMFILE))
  If hBmp = 0 Then
    Call DeleteObject(m_hDC)
    m_hDC = 0
    MsgBox "Can't get Bitmap from ClipBoard"
    Exit Function
  End If

  m_bmi.bmiHeader.biSize = Len(m_bmi.bmiHeader)

    Select Case lImageType
  Case CF_BITMAP
  ' get image sizes
  Call GetDIBits(m_hDC, hBmp, 0, 0, 0, m_bmi, DIB_RGB_COLORS)

    Case CF_ENHMETAFILE
    
lngRet = GetEnhMetaFileHeader(hBmp, Len(mh), mh)

With mh.rclFrame
    ' The rclFrame member Specifies the dimensions,
    ' in .01 millimeter units, of a rectangle that surrounds
    ' the picture stored in the metafile.
    ' I'll show this as seperate steps to aid in understanding
    ' the conversion process.
    
' Convert to MM
sngConvertX = (.right - .Left) * 0.01
sngConvertY = (.Bottom - .top) * 0.01
 End With
 
' Convert to CM
sngConvertX = sngConvertX * 0.1
sngConvertY = sngConvertY * 0.1
' Convert to Inches
sngConvertX = sngConvertX / 2.54
sngConvertY = sngConvertY / 2.54






' Get current Screen DPI
lngIC = apiCreateIC("DISPLAY", vbNullString, vbNullString, vbNullString)
    'If the call to CreateIC didn't fail, then get the Screen X resolution.
    If lngIC <> 0 Then
        lngXdpi = apiGetDeviceCaps(lngIC, LOGPIXELSX)
        'Release the information context.
        apiDeleteDC (lngIC)
    Else
        ' Something has gone wrong. Assume an average value.
        lngXdpi = 120
    End If
    'End If
'нннннннннннннннннннннннннннннннннннннннннна

' Convert the szlMillimeters to inches. This member
' Specifies the resolution of the reference device, in millimeters.
' Convert Inches to Pixels
'sngMetaResolutionX = (mh.szlMillimeters.cx * 0.01) / 2.54
sngMetaResolutionX = (mh.szlDevice.cx / ((mh.szlMillimeters.cx * 0.1) / 2.54))
sngMetaResolutionY = (mh.szlDevice.cy / ((mh.szlMillimeters.cy * 0.1) / 2.54))


m_bmi.bmiHeader.biWidth = CLng(sngConvertX * sngMetaResolutionX)
m_bmi.bmiHeader.biHeight = CLng(sngConvertY * sngMetaResolutionY)

    
    Case Else
    
    End Select

  ' make 24 bpp dib section
  m_bmi.bmiHeader.biBitCount = 24
  m_bmi.bmiHeader.biCompression = BI_RGB
  m_bmi.bmiHeader.biClrUsed = 0
  m_bmi.bmiHeader.biClrImportant = 0
  m_bmi.bmiHeader.biPlanes = 1
  
  
  
  
  m_hDib = CreateDIBSection(m_hDC, m_bmi, DIB_RGB_COLORS, m_lPtr, 0, 0)
  If m_hDib = 0 Then
    Call DeleteObject(hBmp)
    Call DeleteObject(m_hDC)
    m_hDC = 0
    PictureDataToDIB = False
    Exit Function
  End If

' Select the DIBSection into the DC
m_hBmpOld = SelectObject(m_hDC, m_hDib)

' Clear the DIBSection to WHITE
Dim hnewbrush As Long
    
            
' Use White
hnewbrush = apiCreateSolidBrush(RGB(255, 255, 255))
rc.Left = 0
rc.top = 0
rc.right = m_bmi.bmiHeader.biWidth
rc.Bottom = m_bmi.bmiHeader.biHeight

Call apiFillRect(m_hDC, rc, hnewbrush)
Call DeleteObject(hnewbrush)


 Select Case lImageType
  Case CF_BITMAP
  '
 hBMPtemp = SelectObject(hDCtemp, hBmp)

  m_bmi.bmiHeader.biSize = Len(m_bmi.bmiHeader)

  ' get image data in 24 bpp format (convert if need)
 'lngRet = GetDIBits(m_hDC, hBmp, 0, m_bmi.bmiHeader.biHeight, m_lPtr, m_bmi, DIB_RGB_COLORS)


lngRet = BitBlt(m_hDC, 0, 0, m_bmi.bmiHeader.biWidth, m_bmi.bmiHeader.biHeight _
, hDCtemp, 0, 0, vbSrcCopy)

hBmp = SelectObject(hDCtemp, hBMPtemp)
  Call DeleteObject(hBmp)
  Call DeleteDC(hDCtemp)
  
  
 Case CF_ENHMETAFILE
 ' If it is an Enhanced Metafile then we
 ' Need to  "PLAY" the Metafile
 ' back into the Device COntext instead
 ' of using the SelectObject API

  rc.top = 0
 rc.Left = 0
 rc.Bottom = m_bmi.bmiHeader.biHeight
 rc.right = m_bmi.bmiHeader.biWidth
 lngRet = apiPlayEnhMetaFile(m_hDC, hBmp, rc)
 
' Delete the EMF
lngRet = apiDeleteEnhMetaFile(hBmp)
    
Case Else

End Select

  PictureDataToDIB = True
End Function



Public Property Get BytesPerScanLine() As Long
  ' Scans must align on dword boundaries:
  BytesPerScanLine = (m_bmi.bmiHeader.biWidth * (m_bmi.bmiHeader.biBitCount / 8) + 3) And &HFFFFFFFC
End Property


Public Property Get dib_width() As Long
  dib_width = m_bmi.bmiHeader.biWidth
End Property


Public Property Get dib_height() As Long
  dib_height = m_bmi.bmiHeader.biHeight
End Property


Public Property Get dib_channels() As Long
  dib_channels = m_bmi.bmiHeader.biBitCount / 8
End Property

Public Property Get CurrentJpegFileName() As String
CurrentJpegFileName = m_CurrentJpegFileName
End Property

Public Sub PaintPicture( _
  ByVal lhdc As Long, _
  Optional ByVal lDestLeft As Long = 0, _
  Optional ByVal lDestTop As Long = 0, _
  Optional ByVal lDestWidth As Long = -1, _
  Optional ByVal lDestHeight As Long = -1, _
  Optional ByVal lSrcLeft As Long = 0, _
  Optional ByVal lSrcTop As Long = 0, _
  Optional ByVal eRop As Long) ' = vbSrcCopy)

  If (lDestWidth < 0) Then lDestWidth = m_bmi.bmiHeader.biWidth
  If (lDestHeight < 0) Then lDestHeight = m_bmi.bmiHeader.biHeight
Dim lngRet As Long
Dim hDCtemp As Long
'hDCtemp = apiGetDC(0&)

 lngRet = BitBlt(m_hDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, lhdc, lSrcLeft, lSrcTop, vbSrcCopy)
'lngRet = BitBlt(lhdc, lDestLeft, lDestTop, lDestWidth, lDestHeight, hDCtemp, lSrcLeft, lSrcTop, vbSrcCopy)
'lngRet = apiReleaseDC(0&, hDCtemp)
End Sub

Public Function LoadJpegFileIntoArray() As Boolean

On Error GoTo Err_CmdLoad_Click

Dim blRet As Boolean

 ' jpg_scale = 1
  Dim strfName As String
  strfName = Me.CurrentJpegFileName  ' m_cDib.FileDialog 'c:\test2.jpg"
  ' Read JPEG image
 
Dim lPtr As Long
Dim lSize As Long
Dim iFile As Integer
Dim sFile As String
'Dim bArray() As Byte
    
   ' Copy the current Jpeg file data directly to the buffer
   iFile = FreeFile
   Open strfName For Binary Access Read Lock Write As #iFile
   lSize = LOF(iFile)
   ReDim bArray(0 To lSize - 1) As Byte
   Get #iFile, , bArray()
   Close #iFile
   
      
    LoadJpegFileIntoArray = True
Exit_CmdLoad_Click:
    Exit Function

Err_CmdLoad_Click:
LoadJpegFileIntoArray = False
    MsgBox Err.Description
    Resume Exit_CmdLoad_Click
    
End Function



Public Property Get JPegAsByteArray() As Variant
JPegAsByteArray = bArray

End Property

Public Property Get hDC() As Long
  hDC = m_hDC
End Property


Public Property Get hDib() As Long
  hDib = m_hDib
End Property


Public Property Get DIBSectionBitsPtr() As Long
  DIBSectionBitsPtr = m_lPtr
End Property

'Public Property Get DIBBitmapinfo(bm As BITMAPINFO) As Boolean
'DIBBitmapinfo = True
'End Property


Public Function DIBtoPictureData(ctl As Control)
 Dim lngRet As Long
 Dim ds As DIBSECTION
     lngRet = apiGetObject(hDib, Len(ds), ds)
     
    
      '.bfSize = Len(FileHeader) + Len(ds.dsBmih) + ds.dsBmih.biSizeImage
            
       ' Update the Image Control display
        ' We do this by simply copying the mBitmapAdd's contents to
        ' the control's PictureData prop
        
        Dim varTemp() As Byte
        ReDim varTemp(ds.dsBmih.biSizeImage + 40)
        apiCopyMemory varTemp(40), ByVal Me.DIBSectionBitsPtr, ds.dsBmih.biSizeImage
        apiCopyMemory varTemp(0), ds.dsBmih, 40
        
         ctl.PictureData = varTemp


End Function

Public Sub CleanUp()
  
  If (m_hDC <> 0) Then
    If (m_hDib <> 0) Then
      Call SelectObject(m_hDC, m_hBmpOld)
      Call DeleteObject(m_hDib)
    End If
    Call DeleteObject(m_hDC)
  End If
  
  m_hDC = 0
  m_hDib = 0
  m_hBmpOld = 0
  m_lPtr = 0

  m_bmi.bmiColors.rgbBlue = 0
  m_bmi.bmiColors.rgbGreen = 0
  m_bmi.bmiColors.rgbRed = 0
  m_bmi.bmiColors.rgblReterved = 0
  m_bmi.bmiHeader.biSize = Len(m_bmi.bmiHeader)
  m_bmi.bmiHeader.biWidth = 0
  m_bmi.bmiHeader.biHeight = 0
  m_bmi.bmiHeader.biPlanes = 0
  m_bmi.bmiHeader.biBitCount = 0
  m_bmi.bmiHeader.biClrUsed = 0
  m_bmi.bmiHeader.biClrImportant = 0
  m_bmi.bmiHeader.biCompression = 0

End Sub


Private Sub Class_Terminate()
  CleanUp
End Sub


Public Function FileDialog(LoadSave As Boolean) As String
' Calls the API File Dialog Window
' Returns full path to new File.
' If LoadSave = TRUE then call File Load Dialog

On Error GoTo Err_fFileDialog

' Call the File Common Dialog Window
Dim clsDialog As Object
Dim strTemp As String
Dim strfName As String

Set clsDialog = New clsCommonDialog

' Fill in our structure
' I'll leave in how to select Jpeg to
' show you how to build the Filter
clsDialog.Filter = "Bitmap (*.BMP)" & Chr$(0) & "*.BMP" & Chr$(0)

'clsDialog.Filter = "JPEG (*.JPG)" & Chr$(0) & "*.JPG" & Chr$(0)
'clsDialog.Filter = clsDialog.Filter & "Jpe (*.JPE)" & Chr$(0) & "*.JPE" & Chr$(0)
'clsDialog.Filter = clsDialog.Filter & "Jpeg (*.JPEG)" & Chr$(0) & "*.JPEG" & Chr$(0)
'clsDialog.Filter = clsDialog.Filter & "ALL (*.*)" & Chr$(0) & "*.*" & Chr$(0)

'clsDialog.Filter = clsDialog.Filter & "Gif (*.GIF)" & Chr$(0) & "*.GIF" & Chr$(0)


If LoadSave Then
' Display the Open File Dialog
clsDialog.DialogTitle = "Please Select a Bitmap File to Load"
clsDialog.ShowOpen
Else
clsDialog.DialogTitle = "Please Enter/Select a FileName to save the Bitmap File"
clsDialog.ShowSave
End If

' See if user clicked Cancel or even selected
' the very same file already selected
strfName = clsDialog.FileName
If Len(strfName & vbNullString) = 0 Then
Set clsDialog = Nothing
Exit Function
'' Raise the exception
 ' Err.Raise vbObjectError + 513, "clsPrintToFit.fFileDialog", _
 ' "Please type in a Name for a New File"
End If

' Return File Path and Name
FileDialog = strfName
' Update our property
m_CurrentJpegFileName = strfName

Exit_fFileDialog:

Err.Clear
Set clsDialog = Nothing
Exit Function

Err_fFileDialog:
FileDialog = ""
m_CurrentJpegFileName = ""
MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
Resume Exit_fFileDialog

End Function

