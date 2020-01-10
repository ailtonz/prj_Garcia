
Option Compare Database
Option Explicit


' This code is from the Microsoft Knowledge Base.


''API function called by ChooseColor method
'Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
'
''API function called by ShowOpen method
'Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFilename) As Long
'
''API function called by ShowSave method
'Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OpenFilename) As Long
'
''API function to retrieve extended error information
'Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
'
''API memory functions
'Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
'         hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
 

'constants for API memory functions
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
 
 
'data buffer for the ChooseColor function
Private Type ChooseColor
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        Flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

'data buffer for the GetOpenFileName and GetSaveFileName functions
Private Type OpenFilename
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        iFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type


'internal property buffers

Private iAction As Integer         'internal buffer for Action property
Private bCancelError As Boolean    'internal buffer for CancelError property
Private lColor As Long             'internal buffer for Color property
Private lCopies As Long            'internal buffer for lCopies property
Private sDefaultExt As String      'internal buffer for sDefaultExt property
Private sDialogTitle As String     'internal buffer for DialogTitle property
Private sFileName As String        'internal buffer for FileName property
Private sFileTitle As String       'internal buffer for FileTitle property
Private sFilter As String          'internal buffer for Filter property
Private iFilterIndex As Integer    'internal buffer for FilterIndex property
Private lFlags As Long             'internal buffer for Flags property
Private lhdc As Long               'internal buffer for hdc property
Private sInitDir As String         'internal buffer for InitDir property
Private lMax As Long               'internal buffer for Max property
Private lMaxFileSize As Long       'internal buffer for MaxFileSize property
Private lMin As Long               'internal buffer for Min property
Private objObject As Object        'internal buffer for Object property

Private lApiReturn As Long          'internal buffer for APIReturn property
Private lExtendedError As Long      'internal buffer for ExtendedError property



'constants for color dialog

Private Const CDERR_DIALOGFAILURE = &HFFFF
Private Const CDERR_FINDRESFAILURE = &H6
Private Const CDERR_GENERALCODES = &H0
Private Const CDERR_INITIALIZATION = &H2
Private Const CDERR_LOADRESFAILURE = &H7
Private Const CDERR_LOADSTRFAILURE = &H5
Private Const CDERR_LOCKRESFAILURE = &H8
Private Const CDERR_MEMALLOCFAILURE = &H9
Private Const CDERR_MEMLOCKFAILURE = &HA
Private Const CDERR_NOHINSTANCE = &H4
Private Const CDERR_NOHOOK = &HB
Private Const CDERR_NOTEMPLATE = &H3
Private Const CDERR_REGISTERMSGFAIL = &HC
Private Const CDERR_STRUCTSIZE = &H1


'constants for file dialog

Private Const FNERR_BUFFERTOOSMALL = &H3003
Private Const FNERR_FILENAMECODES = &H3000
Private Const FNERR_INVALIDFILENAME = &H3002
Private Const FNERR_SUBCLASSFAILURE = &H3001

Public Property Get Filter() As String
    'return object's Filter property
    Filter = sFilter
End Property

Public Sub ShowColor()
    'display the color dialog box
    
    Dim tChooseColor As ChooseColor
    Dim alCustomColors(15) As Long
    Dim lCustomColorSize As Long
    Dim lCustomColorAddress As Long
    Dim lMemHandle As Long
    
    Dim n As Integer
        
    On Error GoTo ShowColorError
    
    
    '***    init property buffers
    
    iAction = 3  'Action property - ShowColor
    lApiReturn = 0  'APIReturn property
    lExtendedError = 0  'ExtendedError property
    
    
    '***    prepare tChooseColor data
    
    'lStructSize As Long
    tChooseColor.lStructSize = Len(tChooseColor)
    
    'hwndOwner As Long
    tChooseColor.hwndOwner = Application.hWndAccessApp

    'hInstance As Long
    
    'rgbResult As Long
    tChooseColor.rgbResult = lColor
    
    'lpCustColors As Long
    ' Fill custom colors array with all white
    For n = 0 To UBound(alCustomColors)
        alCustomColors(n) = &HFFFFFF
    Next
    ' Get size of memory needed for custom colors
    lCustomColorSize = Len(alCustomColors(0)) * 16
    ' Get a global memory block to hold a copy of the custom colors
    lMemHandle = GlobalAlloc(GHND, lCustomColorSize)
    
    If lMemHandle = 0 Then
        Exit Sub
    End If
    ' Lock the custom color's global memory block
    lCustomColorAddress = GlobalLock(lMemHandle)
    If lCustomColorAddress = 0 Then
        Exit Sub
    End If
    ' Copy custom colors to the global memory block
    Call CopyMemory(ByVal lCustomColorAddress, alCustomColors(0), lCustomColorSize)
 
    tChooseColor.lpCustColors = lCustomColorAddress
    
    'flags As Long
    tChooseColor.Flags = lFlags
        
    'lCustData As Long
    'lpfnHook As Long
    'lpTemplateName As String
    
    
    '***    call the ChooseColor API function
    lApiReturn = ChooseColor(tChooseColor)
    
    
    '***    handle return from ChooseColor API function
    Select Case lApiReturn
        
        Case 0  'user canceled
        If bCancelError = True Then
            'generate an error
            On Error GoTo 0
            Err.Raise Number:=vbObjectError + 894, _
                Description:="Cancel Pressed"
            Exit Sub
        End If
        
        Case 1  'user selected a color
            'update property buffer
            lColor = tChooseColor.rgbResult
        
        Case Else   'an error occured
            'call CommDlgExtendedError
            lExtendedError = CommDlgExtendedError
        
    End Select

Exit Sub

ShowColorError:
    Exit Sub
End Sub

Public Sub ShowOpen()
    'display the file open dialog box
    ShowFileDialog (1)  'Action property - ShowOpen
End Sub

Public Sub ShowSave()
    'display the file save dialog box
    ShowFileDialog (2)  'Action property - ShowSave
End Sub

Public Property Get FileName() As String
    'return object's FileName property
    FileName = sFileName
End Property

Public Property Let FileName(vNewValue As String)
    'assign object's FileName property
    sFileName = vNewValue
End Property

Public Property Let Filter(vNewValue As String)
    'assign object's Filter property
    sFilter = vNewValue
End Property

Private Function sLeftOfNull(ByVal sIn As String)
    'returns the part of sIn preceding Chr$(0)
    Dim lNullPos As Long
    
    'init output
    sLeftOfNull = sIn
    
    'get position of first Chr$(0) in sIn
    lNullPos = InStr(sIn, Chr$(0))
    
    'return part of sIn to left of first Chr$(0) if found
    If lNullPos > 0 Then
        sLeftOfNull = Mid$(sIn, 1, lNullPos - 1)
    End If
    
End Function


Public Property Get Action() As Integer
    'Return object's Action property
    Action = iAction
End Property

Private Function sAPIFilter(sIn)
    'prepares sIn for use as a filter string in API common dialog functions
    Dim lChrNdx As Long
    Dim sOneChr As String
    Dim sOutStr As String
    
    'convert any | characters to nulls
    For lChrNdx = 1 To Len(sIn)
        sOneChr = Mid$(sIn, lChrNdx, 1)
        If sOneChr = "|" Then
            sOutStr = sOutStr & Chr$(0)
        Else
            sOutStr = sOutStr & sOneChr
        End If
    Next
    
    'add a null to the end
    sOutStr = sOutStr & Chr$(0)
    
    'return sOutStr
    sAPIFilter = sOutStr
    
End Function

Public Property Get FilterIndex() As Integer
    'return object's FilterIndex property
    FilterIndex = iFilterIndex
End Property

Public Property Let FilterIndex(vNewValue As Integer)
    iFilterIndex = vNewValue
End Property

Public Property Get CancelError() As Boolean
    'Return object's CancelError property
    CancelError = bCancelError
End Property

Public Property Let CancelError(vNewValue As Boolean)
    'Assign object's CancelError property
    bCancelError = vNewValue
End Property

Public Property Get Color() As Long
    'return object's Color property
    Color = lColor
End Property

Public Property Let Color(vNewValue As Long)
    'assign object's Color property
    lColor = vNewValue
End Property

Public Property Get DefaultExt() As String
    'return object's DefaultExt property
    DefaultExt = sDefaultExt
End Property

Public Property Let DefaultExt(vNewValue As String)
    'assign object's DefaultExt property
    sDefaultExt = vNewValue
End Property

Public Property Get DialogTitle() As String
    'return object's FileName property
    DialogTitle = sDialogTitle
End Property

Public Property Let DialogTitle(vNewValue As String)
    'assign object's DialogTitle property
    sDialogTitle = vNewValue
End Property

Public Property Get Flags() As Long
    'return object's Flags property
    Flags = lFlags
End Property

Public Property Let Flags(vNewValue As Long)
    'assign object's Flags property
    lFlags = vNewValue
End Property

Public Property Get hDC() As Long
    'Return object's hDC property
    hDC = lhdc
End Property

Public Property Let hDC(vNewValue As Long)
    'Assign object's hDC property
    lhdc = vNewValue
End Property

Public Property Get InitDir() As String
    'Return object's InitDir property
    InitDir = sInitDir
End Property

Public Property Let InitDir(vNewValue As String)
    'Assign object's InitDir property
    sInitDir = vNewValue
End Property

Public Property Get Max() As Long
    'Return object's Max property
    Max = lMax
End Property

Public Property Let Max(vNewValue As Long)
    'Assign object's - property
    lMax = vNewValue
End Property

Public Property Get MaxFileSize() As Long
    'Return object's MaxFileSize property
    MaxFileSize = lMaxFileSize
End Property

Public Property Let MaxFileSize(vNewValue As Long)
    'Assign object's MaxFileSize property
    lMaxFileSize = vNewValue
End Property

Public Property Get Min() As Long
    'Return object's Min property
    Min = lMin
End Property

Public Property Let Min(vNewValue As Long)
    'Assign object's Min property
    lMin = vNewValue
End Property

Public Property Get Object() As Object
    'Return object's Object property
    Object = objObject
End Property

Public Property Let Object(vNewValue As Object)
    'Assign object's Object property
    objObject = vNewValue
End Property

Public Property Get FileTitle() As String
    'return object's FileTitle property
    FileTitle = sFileTitle
End Property

Public Property Let FileTitle(vNewValue As String)
    'assign object's FileTitle property
    sFileTitle = vNewValue
End Property

Public Property Get APIReturn() As Long
    'return object's APIReturn property
    APIReturn = lApiReturn
End Property

Public Property Get ExtendedError() As Long
    'return object's ExtendedError property
    ExtendedError = lExtendedError
End Property


Private Function sByteArrayToString(abBytes() As Byte) As String
    'return a string from a byte array
    Dim lBytePoint As Long
    Dim lByteVal As Long
    Dim sOut As String
    
    'init array pointer
    lBytePoint = LBound(abBytes)
    
    'fill sOut with characters in array
    While lBytePoint <= UBound(abBytes)
        
        lByteVal = abBytes(lBytePoint)
        
        'return sOut and stop if Chr$(0) is encountered
        If lByteVal = 0 Then
            sByteArrayToString = sOut
            Exit Function
        Else
            sOut = sOut & Chr$(lByteVal)
        End If
        
        lBytePoint = lBytePoint + 1
    
    Wend
    
    'return sOut if Chr$(0) wasn't encountered
    sByteArrayToString = sOut
    
End Function
Private Sub ShowFileDialog(ByVal iAction As Integer)
    
    'display the file dialog for ShowOpen or ShowSave
    
    Dim tOpenFile As OpenFilename
    Dim lMaxSize As Long
    Dim sFileNameBuff As String
    Dim sFileTitleBuff As String
    
    On Error GoTo ShowFileDialogError
    
    
    '***    init property buffers
    
    iAction = iAction  'Action property
    lApiReturn = 0  'APIReturn property
    lExtendedError = 0  'ExtendedError property
        
    
    '***    prepare tOpenFile data
    
    'tOpenFile.lStructSize As Long
    tOpenFile.lStructSize = Len(tOpenFile)
    
    'tOpenFile.hWndOwner As Long - init from hdc property
    tOpenFile.hwndOwner = Application.hWndAccessApp
    
    'tOpenFile.lpstrFilter As String - init from Filter property
    tOpenFile.lpstrFilter = sAPIFilter(sFilter)
        
    'tOpenFile.iFilterIndex As Long - init from FilterIndex property
    tOpenFile.iFilterIndex = iFilterIndex
    
    'tOpenFile.lpstrFile As String
        'determine size of buffer from MaxFileSize property
        If lMaxFileSize > 0 Then
            lMaxSize = lMaxFileSize
        Else
            lMaxSize = 256
        End If
    
    'tOpenFile.lpstrFile As Long - init from FileName property
        'prepare sFileNameBuff
        sFileNameBuff = sFileName
        'pad with spaces
        While Len(sFileNameBuff) < lMaxSize - 1
            sFileNameBuff = sFileNameBuff & " "
        Wend
        'trim to length of lMaxFileSize - 1
       sFileNameBuff = Mid$(sFileNameBuff, 1, lMaxFileSize - 1)
        'null terminate
        sFileNameBuff = sFileNameBuff & Chr$(0)
    tOpenFile.lpstrFile = sFileNameBuff
    
    'nMaxFile As Long - init from MaxFileSize property
    If lMaxFileSize <> 255 Then  'default is 255
        tOpenFile.nMaxFile = lMaxFileSize
    End If
            
    'lpstrFileTitle As String - init from FileTitle property
        'prepare sFileTitleBuff
        sFileTitleBuff = sFileTitle
        'pad with spaces
        While Len(sFileTitleBuff) < lMaxSize - 1
            sFileTitleBuff = sFileTitleBuff & " "
        Wend
        'trim to length of lMaxFileSize - 1
        sFileTitleBuff = Mid$(sFileTitleBuff, 1, lMaxFileSize - 1)
        'null terminate
        sFileTitleBuff = sFileTitleBuff & Chr$(0)
    tOpenFile.lpstrFileTitle = sFileTitleBuff
        
    'tOpenFile.lpstrInitialDir As String - init from InitDir property
    tOpenFile.lpstrInitialDir = sInitDir
    
    'tOpenFile.lpstrTitle As String - init from DialogTitle property
    tOpenFile.lpstrTitle = sDialogTitle
    
    'tOpenFile.flags As Long - init from Flags property
    tOpenFile.Flags = lFlags
        
    'tOpenFile.lpstrDefExt As String - init from DefaultExt property
    tOpenFile.lpstrDefExt = sDefaultExt
    
    
    '***    call the GetOpenFileName API function
    Select Case iAction
        Case 1  'ShowOpen
            lApiReturn = GetOpenFileName(tOpenFile)
        Case 2  'ShowSave
            lApiReturn = GetSaveFileName(tOpenFile)
        Case Else   'unknown action
            Exit Sub
    End Select
    
    
    '***    handle return from GetOpenFileName API function
    Select Case lApiReturn
        
        Case 0  'user canceled
        If bCancelError = True Then
            'generate an error
            Err.Raise (2001)
            Exit Sub
        End If
        
        Case 1  'user selected or entered a file
            'sFileName gets part of tOpenFile.lpstrFile to the left of first Chr$(0)
            sFileName = sLeftOfNull(tOpenFile.lpstrFile)
            sFileTitle = sLeftOfNull(tOpenFile.lpstrFileTitle)
        
        Case Else   'an error occured
            'call CommDlgExtendedError
            lExtendedError = CommDlgExtendedError
        
    End Select
    

Exit Sub

ShowFileDialogError:
    
    Exit Sub

End Sub


Private Sub Class_Initialize()
lMaxFileSize = 256
End Sub
