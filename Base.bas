Attribute VB_Name = "Base"
Option Private Module
Option Explicit

Private Const ADB_ADB_PATH As String = "C:\Users\M\Documents\Application\platform-tools\adb.exe"
'Private Const WSCRIPT_RUN_WINDOW_STATE As Integer = 7
Private Const WSCRIPT_RUN_WINDOW_STATE As Integer = 0

Private Type Point: x As Long: Y As Long: End Type
Private Type Rectangle: Left As Long: Top As Long: Right As Long: Bottom As Long: End Type
Private Type SystemTime: Year As Integer: Month As Integer: DayOfWeek As Integer: Day As Integer: Hour As Integer: Minute As Integer: Second As Integer: Milliseconds As Integer: End Type
Private Type FileTime: LowDateTime As Long: HighDateTime As Long: End Type
Private Type UUID: Data1 As Long: Data2 As Integer: Data3 As Integer: Data4(7) As Byte: End Type
Private Type PicDesc: cbSizeOfStruct As Long: PICTYPE As Long: bmp As Long: icon As Long: emf As Long: End Type
Private Type bitmap: Type As Long: Width As Long: Height As Long: WidthBytes As Long: Planes As Integer: BitsPixel As Integer: Bits As Long: End Type

Private Type LONG_TYPE: val As Long: End Type
Private Type BYTE_TYPE: val(3) As Byte: End Type

#If VBA7 And Win64 Then
Private Type GdiplusStartupInput: GdiplusVersion As Long: DebugEventCallback As LongPtr: SuppressBackgroundThread As Long: SuppressExternalCodecs As Long: End Type
#Else
Private Type GdiplusStartupInput: GdiplusVersion As Long: DebugEventCallback As Long: SuppressBackgroundThread As Long: SuppressExternalCodecs As Long: End Type
#End If

Public Enum MouseButton
    Left = 2
    Right = 8
End Enum

Public Enum WindowState
    Nomal = 0
    min = 1
    max = 2
End Enum

Public Enum AccessibleState
    Disable = &H1
    Selected = &H2
    Focus = &H4
    Pressed = &H8
    Checked = &H10
    Mixed = &H20
    ReadOnly = &H40
    HotTracked = &H80
    Default = &H100
    Expanded = &H200
    Collapsed = &H400
    Busy = &H800
    Floating = &H1000
    Marqueed = &H2000
    Animated = &H4000
    Hidden = &H8000
    Offscreen = &H10000
    Sizeable = &H20000
    Moveable = &H40000
    SelfVoicing = &H80000FF
    Focusable = &H100000
    Selectable = &H200000
    Linked = &H400000
    Traversed = &H800000
    MultiSelectable = &H1000000
    ExtSelectable = &H2000000
    AlertLow = &H4000000
    AlertMedium = &H8000000
    AlertHigh = &H10000000
    HasPopup = &H40000000
    Valid = &H1FFFFFFF
End Enum

Public Enum Providers
    Sqloledb = 1
    Msdasql = 2
    MicrosoftExcelDriver = 4
    MicrosoftTextDriver = 8
End Enum

Public Enum Exceptions
    ArgumentOutOfRangeException = 5
    IndexOutOfRangeException = 9
    ArrayTemporarilyLockedException = 10
    InvalidTypeException = 13
    FileNotFoundException = 53
    FileExistsException = 58
    DuplicateKeyException = 457
    FeatureNotAvailableException = 3251
End Enum

Private Const DOCUMENT_IMAGE_ID       As Long = 0
Private Const DOCUMENT_IMAGE_TEXT     As Long = 1
Private Const DOCUMENT_IMAGE_LEFT     As Long = 2
Private Const DOCUMENT_IMAGE_RIGHT    As Long = 3
Private Const DOCUMENT_IMAGE_TOP      As Long = 4
Private Const DOCUMENT_IMAGE_BOTTOM   As Long = 5
Private Const DOCUMENT_IMAGE_LINEID   As Long = 6
Private Const DOCUMENT_IMAGE_REGIONID As Long = 7

Private TEMP_ARRAY As Variant

Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal x As Long, ByVal Y As Long, ByVal cButtons As Long, ByVal dwExtraInfo As LongPtr)
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As Point) As Long
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Private Declare PtrSafe Function MapVirtualKeyA Lib "user32" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal ClassName As String, ByVal WindowName As String) As Long
Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, lParam As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function GetWindowTextA Lib "user32" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindowTextLengthA Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetClassNameA Lib "user32" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rectangle) As Long
Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function LoadImageA Lib "user32" (ByVal hinst As Long, ByVal lpszName As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Declare PtrSafe Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FileTime, lpLastAccessTime As FileTime, lpLastWriteTime As FileTime) As Long
Private Declare PtrSafe Function SystemTimeToFileTime Lib "kernel32" (ByRef lpSystemTime As SystemTime, ByRef lpFileTime As FileTime) As Long
Private Declare PtrSafe Function LocalFileTimeToFileTime Lib "kernel32" (ByRef lpLocalFileTime As FileTime, ByRef lpFileTime As FileTime) As Long
Private Declare PtrSafe Function CreateFileA Lib "kernel32" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare PtrSafe Function GetTempFileNameA Lib "kernel32" (ByVal DirName As String, ByVal Prefix As String, ByVal Unique As Long, ByVal TempFile As String) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function AccessibleObjectFromPoint Lib "oleacc" (ByVal Pos As LongPtr, ByRef accObject As Any, ByRef accChild As Variant) As Long
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hWnd As Long, ByVal dwObjectID As Long, ByRef riid As UUID, ByRef ppvObject As Any) As Long
Private Declare PtrSafe Function AccessibleChildren Lib "oleacc" (ByVal paccContainer As Object, ByVal iChildStart As Long, ByVal cChildren As Long, ByRef rgvarChildren As Any, ByRef pcObtained As Long) As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (ByRef lpPictDesc As PicDesc, ByRef RefIID As UUID, ByVal fPictureOwnHandle As LongPtr, ByRef IPic As IPicture) As Long
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare PtrSafe Function GetObjectA Lib "gdi32" (ByVal hObject As Long, ByVal nCount As Integer, ByRef lpObject As Any) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef inputBuf As GdiplusStartupInput, Optional ByVal outputBuf As Long = 0) As Long
Private Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As LongPtr, ByRef Image As Long) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As Long, hbmReturn As Long, ByVal background As Long) As Long
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare PtrSafe Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Long

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long


Public Function IO≫File≫IsExists(ByVal Path As String) As Boolean
    On Error Resume Next
        Call VBA.GetAttr(Path)
        IO≫File≫IsExists = (VBA.Err.Number = 0)
    On Error GoTo 0
End Function

Public Function IO≫File≫Move(ByVal sourceFileName As String, ByVal destFileName As String) As String
    Name sourceFileName As destFileName
    IO≫File≫Move = destFileName
End Function

Public Function IO≫File≫Copy(ByVal sourceFileName As String, ByVal destFileName As String) As String
    VBA.FileCopy sourceFileName, destFileName
    IO≫File≫Copy = destFileName
End Function

Public Function IO≫File≫CreateTempFile(ByVal Path As String, Optional ByVal Prefix As String = "") As String
    Dim n As String * &HFF
    GetTempFileNameA Path, Prefix, 0, n
    IO≫File≫CreateTempFile = VBA.Replace(n, VBA.Right(n, 1), VBA.vbNullString)
End Function

Public Sub IO≫File≫CreateFile(ByVal Path As String)
    If IO≫File≫IsExists(Path) Then
        Error FileExistsException
    Else
        IO≫File≫WriteText Path, vbNullString
    End If
End Sub

Public Sub IO≫File≫CreateDirectory(ByVal Path As String)
    If IO≫File≫IsExists(Path) Then
        VBA.Error FileExistsException
    Else
        VBA.MkDir Path
    End If
End Sub

Public Sub IO≫File≫Delete(ByVal Path As String)
    VBA.Kill Path
End Sub

Public Function IO≫File≫IsDirectory(ByVal Path As String) As Boolean
    IO≫File≫IsDirectory = (VBA.GetAttr(Path) And VBA.VbFileAttribute.vbDirectory) = VBA.VbFileAttribute.vbDirectory
End Function

Public Function IO≫File≫GetFiles(ByVal Path As String, Optional ByVal ChildDir As Boolean = False, Optional ByVal FileAttribute As VBA.VbFileAttribute = vbNormal, Optional ByVal FileFilter As String = VBA.vbNullString) As String()
    
    Dim i, j As Long
    
    Dim result(), d() As String
    ReDim d(0) As String
    d(0) = Path
    
    If FileAttribute = VBA.VbFileAttribute.vbNormal Then _
        FileAttribute = VBA.VbFileAttribute.vbAlias + _
                        VBA.VbFileAttribute.vbArchive + _
                        VBA.VbFileAttribute.vbDirectory + _
                        VBA.VbFileAttribute.vbHidden + _
                        VBA.VbFileAttribute.vbNormal + _
                        VBA.VbFileAttribute.vbReadOnly + _
                        VBA.VbFileAttribute.vbSystem + _
                        VBA.VbFileAttribute.vbVolume
    
    Do
        Dim strDirPath As String
        strDirPath = VBA.Dir(BuildPath(d(i), FileFilter), VBA.VbFileAttribute.vbDirectory)
        
        Do _
            While strDirPath <> VBA.vbNullString
            
            If strDirPath <> "." And strDirPath <> ".." Then
                
                strDirPath = BuildPath(d(i), strDirPath)
                
                Dim a As VBA.VbFileAttribute
                a = VBA.GetAttr(strDirPath)
                
                If (a And VBA.VbFileAttribute.vbDirectory) = VBA.VbFileAttribute.vbDirectory Then
                    j = j + 1
                    ReDim Preserve d(j)
                    d(j) = strDirPath
                End If
                
                If a And FileAttribute Then _
                    Array≫AddItem result, strDirPath
                
            End If
            strDirPath = Dir
        Loop
        
        If Not ChildDir Then Exit Do
        i = i + 1
        
    Loop _
    While i <= j
    
    IO≫File≫GetFiles = result
    
End Function

Public Function IO≫File≫Grep(ByVal Path As String, ByVal Pattern As String, Optional ByVal IgnoreCase As Boolean = False, Optional ChildDir As Boolean = True) As VBA.Collection
    
    Dim result As New VBA.Collection
    Dim f As Variant
    
    Static a As VBA.VbFileAttribute
    If a = VBA.VbFileAttribute.vbNormal Then _
        a = VBA.VbFileAttribute.vbAlias + _
            VBA.VbFileAttribute.vbArchive + _
            VBA.VbFileAttribute.vbDirectory + _
            VBA.VbFileAttribute.vbHidden + _
            VBA.VbFileAttribute.vbNormal + _
            VBA.VbFileAttribute.vbReadOnly + _
            VBA.VbFileAttribute.vbSystem + _
            VBA.VbFileAttribute.vbVolume
    
    For Each f In IO≫File≫GetFiles(Path, ChildDir, a - vbDirectory)
        Dim m As Variant
        Set m = IO≫Text≫Match(IO≫Text≫ReadAllText(f), Pattern, IgnoreCase)
        If m.Matches.Count > 0 Then result.Add m, f
    Next
    
    Set IO≫File≫Grep = result
    
End Function

Public Function IO≫File≫FileInfo≫GetLastModifiedDate(ByVal Path As String) As Date
    IO≫File≫FileInfo≫GetLastModifiedDate = GetFsoFileObject(Path).DateLastModified
End Function

Public Function IO≫File≫FileInfo≫GetLastAccessDate(ByVal Path As String) As Date
    IO≫File≫FileInfo≫GetLastAccessDate = GetFsoFileObject(Path).DateLastAccessed
End Function

Public Function IO≫File≫FileInfo≫IsReadOnly(ByVal Path As String) As Boolean
    IO≫File≫FileInfo≫IsReadOnly = (VBA.GetAttr(Path) And VBA.VbFileAttribute.vbReadOnly) = VBA.VbFileAttribute.vbReadOnly
End Function

Public Function IO≫File≫FileInfo≫IsSystemFile(ByVal Path As String) As Boolean
    IO≫File≫FileInfo≫IsSystemFile = (VBA.GetAttr(Path) And VBA.VbFileAttribute.bSystem) = VBA.VbFileAttribute.vbSystem
End Function

Public Function IO≫File≫FileInfo≫IsHidden(ByVal Path As String) As Boolean
    IO≫File≫FileInfo≫IsHidden = (VBA.GetAttr(Path) And VBA.VbFileAttribute.vbHidden) = VBA.VbFileAttribute.vbHidden
End Function

Public Function IO≫File≫FileInfo≫IsArchive(ByVal Path As String) As Boolean
    IO≫File≫FileInfo≫IsArchive = (VBA.GetAttr(Path) And VBA.VbFileAttribute.vbArchive) = VBA.VbFileAttribute.vbArchive
End Function

Public Sub IO≫File≫FileInfo≫SetReadOnly(ByVal Path As String, ByVal ReadOnly As Boolean)
    SetAttributes Path, VBA.VbFileAttribute.vbReadOnly, ReadOnly
End Sub

Public Sub IO≫File≫FileInfo≫SetSystemFile(ByVal Path As String, ByVal SystemFile As Boolean)
    SetAttributes Path, VBA.VbFileAttribute.vbSystem, SystemFile
End Sub

Public Sub IO≫File≫FileInfo≫SetHidden(ByVal Path As String, ByVal Hidden As Boolean)
    SetAttributes Path, VBA.VbFileAttribute.vbHidden, Hidden
End Sub

Public Sub IO≫File≫FileInfo≫SetArchive(ByVal Path As String, ByVal Archive As Boolean)
    SetAttributes Path, VBA.VbFileAttribute.vbArchive, Archive
End Sub

Public Sub IO≫File≫FileInfo≫SetCreateDate(ByVal Path As String, ByVal DateTime As Date)
    SetFileTimes Path, DateTime, 0, 0
End Sub

Public Sub IO≫File≫FileInfo≫SetModifiedDate(ByVal Path As String, ByVal DateTime As Date)
    SetFileTimes Path, 0, DateTime, 0
End Sub

Public Sub IO≫File≫FileInfo≫SetAccessDate(ByVal Path As String, ByVal DateTime As Date)
    SetFileTimes Path, 0, 0, DateTime
End Sub

Public Function IO≫File≫FileInfo≫GetCreateDate(ByVal Path As String) As Date
    IO≫Path≫GetCreateDate = GetFsoFileObject(Path).DateCreated
End Function

Public Function IO≫Path≫GetFileName(ByVal Path As String) As String
    IO≫Path≫GetFileName = TryMid(Path, InStrRev(Path, "\") + 1)
End Function

Public Function IO≫Path≫GetFolderPath(ByVal Path As String) As String
    IO≫Path≫GetFolderPath = TryMid(Path, 1, InStrRev(Path, "\") - 1)
End Function

Public Function IO≫Path≫GetExtension(ByVal Path As String) As String
    IO≫Path≫GetExtension = TryMid(Path, 1, InStrRev(Path, ".") - 1)
End Function

Public Function IO≫Path≫GetTempPath()
    IO≫Path≫GetTempPath = VBA.Environ("TEMP")
End Function

Public Function IO≫Text≫ReadAllText(ByVal Path As String) As String
    With VBA.CreateObject("Scripting.FileSystemObject").GetFile(Path)
        If .Size > 0 Then
            With .OpenAsTextStream(1, -2)
                IO≫Text≫ReadAllText = .ReadAll
                .Close
            End With
        End If
    End With
End Function

Public Sub IO≫Text≫WriteText(ByVal Path As String, ByVal str As String, Optional ByRef FileNo As Long = 0)
    If FileNo = 0 Then
        FileNo = FreeFile
        Open Path For Append As #FileNo
        Print #FileNo, str
        Close #FileNo
    Else
        Print #FileNo, str
    End If
End Sub

Public Sub IO≫Text≫ClearText(ByVal Path As String, Optional ByRef FileNo As Long = 0)
    Dim i As Long
    If FileNo = 0 Then
        i = FreeFile
    Else
        i = FileNo
        Close #i
    End If
    Open Path For Output As #i
    Print #i, VBA.vbNullString
    Close #i
    If FileNo = 0 Then Open Path For Append As #i
End Sub

Public Function IO≫Text≫Match(ByVal str As String, ByVal Pattern As String, Optional ByVal IgnoreCase As Boolean = False) As Object
    With VBA.CreateObject("VBScript.RegExp")
        .Global = True
        .IgnoreCase = IgnoreCase
        .Pattern = Pattern
        Set IO≫Text≫Match = .Execute(str)
    End With
End Function

Private Function TryMid(ByVal str As String, ByVal Start As Long, Optional ByVal Length As Long = -1) As String
On Error GoTo Exception
        If Length >= 0 Then
            TryMid = VBA.Mid(str, Start)
        Else
            TryMid = VBA.Mid(str, Start, Length)
        End If
        Exit Function
Exception:
        If VBA.Err.Number = ArgumentOutOfRangeException Then
            TryMid = VBA.vbNullString
            On Error GoTo 0
        Else
            VBA.Err.Raise VBA.Err.Number, VBA.Err.Source, VBA.Err.Description, VBA.Err.HelpFile, VBA.Err.HelpContext
        End If
End Function

Private Function BuildPath(ByVal Path As String, ByVal FileName As String) As String
    BuildPath = Path & "\" & FileName
End Function

Private Sub SetAttributes(ByVal Path As String, ByVal Attr As VbFileAttribute, ByVal Value As Boolean)
    
    Dim a As VBA.VbFileAttribute
    a = VBA.GetAttr(Path)
    
    If ((a And Attr) = Attr) <> Value Then
        VBA.SetAttr VBA.IIf(Value, a + Attr, a - Attr)
    End If
    
End Sub

Private Function GetFsoFileObject(ByVal Path As String) As Object
    If IO≫File≫IsDirectory(Path) Then
        Set GetFsoFileObject = VBA.CreateObject("Scripting.FileSystemObject").GetFolder(Path)
    Else
        Set GetFsoFileObject = VBA.CreateObject("Scripting.FileSystemObject").GetFile(Path)
    End If
End Function

Private Sub SetFileTimes(ByVal Path As String, ByVal CreateDate As Date, ByVal ModifiedDate As Date, ByVal AccessDate As Date)
    
    Dim ct, mt, at As FileTime
    If CreateDate <> 0 Then ct = GetFileTime(CreateDate)
    If ModifiedDate <> 0 Then mt = GetFileTime(ModifiedDate)
    If AccessDate <> 0 Then ac = GetFileTime(AccessDate)
    
    Dim h As Long
    h = CreateFileA(Path, &H80000000 Or &H40000000, &H1, 0, 3, &H80, 0)
    
    If h > 0 Then
        On Error Resume Next
            Call SetFileTime(h, ct, at, mt)
            Call CloseHandle(h)
        On Error GoTo 0
    End If
    
End Sub

Private Function GetFileTime(ByVal d As Date) As FileTime
    
    Dim st As SystemTime
    Dim lt, ft As FileTime
    
    With st
        .Year = VBA.Year(d)
        .Month = VBA.Month(d)
        .DayOfWeek = VBA.Weekday(d)
        .Day = VBA.Day(d)
        .Hour = VBA.Hour(d)
        .Minute = VBA.Minute(DateTime)
        .Second = VBA.Second(d)
    End With
    
    Call SystemTimeToFileTime(st, lt)
    Call LocalFileTimeToFileTime(lt, ft)
    
    GetFileTime = ft
    
End Function

Public Function IO≫Image≫ConvertImageFormat(ByVal Before As String, ByVal After As String) As String
    
    Const WIA_FORMAT_UNDEFINED = "{00000000-0000-0000-0000-000000000000}"
    Const WIA_FORMAT_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
    
    Dim img As Object
    Set img = VBA.CreateObject("WIA.ImageFile")
    img.LoadFile Before
    
    Select Case img.FormatId
        Case WIA_FORMAT_UNDEFINED
        Case WIA_FORMAT_BMP
        Case Else
            Dim ip As Object
            Set ip = VBA.CreateObject("WIA.ImageProcess")
            ip.Filters.Add ip.FilterInfos("Convert").FilterID
            ip.Filters(1).Properties("FormatID").Value = WIA_FORMAT_BMP
            ip.Apply img
            img.SaveFile After
            Set ip = Nothing
    End Select
    
    IO≫Image≫ConvertImageFormat = After
    Set img = Nothing
    
End Function

Public Function IO≫Image≫CropImage(ByVal LoadImage As String, ByVal SaveImage As String, ByVal Left As Long, ByVal Right As Long, ByVal Top As Long, ByVal Bottom As Long) As String
    
    Dim img As Object
    Set img = VBA.CreateObject("WIA.ImageFile")
    img.LoadFile LoadImage
    
    With VBA.CreateObject("WIA.ImageProcess")
        .Filters.Add .FilterInfos("Crop").FilterID
        .Filters(1).Properties("Left") = Left
        .Filters(1).Properties("Top") = Top
        .Filters(1).Properties("Right") = img.Width - Right
        .Filters(1).Properties("Bottom") = img.Height - Bottom
        Set img = .Apply(img)
        img.SaveFile SaveImage
    End With
    
    IO≫Image≫CropImage = SaveImage
    Set img = Nothing
    
End Function

Public Function IO≫Image≫GetBitmapFromPng(ByVal Path As String) As StdPicture
    
    Dim g As GdiplusStartupInput
    g.GdiplusVersion = 1
    
    Dim t As Long
    GdiplusStartup t, g
    
    Dim i As Long
    GdipLoadImageFromFile ByVal StrPtr(Path), i
    
    Dim b As Long
    GdipCreateHBITMAPFromBitmap i, b, 0&
    
    Dim p As IPictureDisp
    Set IO≫Image≫GetBitmapFromPng = CreateBitmap(b)
    
    GdipDisposeImage i
    GdiplusShutdown t
    
End Function

Public Function IO≫Image≫GetBitmapFromFile(ByVal Path As String) As StdPicture
    
    Dim i As Long
    i = LoadImageA(0, Path, 0, 0, 0, &H10)
    
    Dim hdc1 As Long
    hdc1 = CreateCompatibleDC(0)
    
    Dim hdc2 As Long
    hdc2 = SelectObject(hdc1, i)
    
    Dim b As bitmap
    GetObjectA i, Len(b), b
    
    If i = 0 Then
        Set IO≫Image≫GetBitmapFromFile = Nothing
    Else
        Set IO≫Image≫GetBitmapFromFile = CreateBitmap(i)
    End If
    
    SelectObject hdc1, hdc2
    DeleteDC hdc1
    
End Function

Public Function IO≫Image≫GetBitmapFromWindow(ByVal hWndOrWindowTitle) As StdPicture
    Dim i As Long
    i = GetBitmapHandleFromWindow(UI≫Window≫GetWindowHandle(hWndOrWindowTitle))
    If i = 0 Then
        Set IO≫Image≫GetBitmapFromWindow = Nothing
    Else
        Set IO≫Image≫GetBitmapFromWindow = CreateBitmap(i)
    End If
End Function

Private Function GetBitmapHandleFromWindow(ByVal hWnd As Long) As Long
    
    Dim r As VBA.Collection
    If hWnd = 0 Then
        Set r = GetDesktopSize()
    Else
        Set r = UI≫Window≫GetWindowRectangle(hWnd)
    End If
    
    Dim w As Long: w = r("Right") - r("Left")
    Dim h As Long: h = r("Bottom") - r("Top")
    
    Dim hdc1 As Long
    'hdc1 = GetWindowDC(hWnd)
    hdc1 = GetDC(0)
    
    Dim hdc2 As Long
    hdc2 = CreateCompatibleDC(hdc1)
    
    Dim bmp1 As Long
    bmp1 = CreateCompatibleBitmap(hdc1, w, h)
    
    Dim bmp2 As Long
    bmp2 = SelectObject(hdc2, bmp1)
    
    'Debug.Print w & "," & h&; "," & r("Left") & "," & r("Top")
    BitBlt hdc2, 0, 0, w, h, hdc1, r("Left"), r("Top"), &HCC0020
    GetBitmapHandleFromWindow = bmp1
    
    ReleaseDC hWnd, hdc1
    SelectObject hdc2, bmp2
    DeleteDC hdc2
    
End Function

Private Function CreateBitmap(ByVal BitmapHandle As Long) As StdPicture
    
    Dim u As UUID
    Dim p As PicDesc
    Dim i As StdPicture
    
    With p
        .cbSizeOfStruct = Len(p)
        .PICTYPE = 1
        .bmp = BitmapHandle
    End With
    
    With u
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    OleCreatePictureIndirect p, u, 0&, i
    Set CreateBitmap = i
    
End Function

Public Function IO≫Image≫CompareImage(ByRef Image1 As StdPicture, ByRef Image2 As StdPicture, Optional ByVal Left1 As Long = 1, Optional ByVal Top1 As Long = 1, Optional ByVal Right1, Optional ByVal Bottom1, Optional ByVal Left2, Optional ByVal Top2, Optional ByVal Right2, Optional ByVal Bottom2) As Boolean
    
    IO≫Image≫CompareImage = False
    
    If VBA.IsError(Right1) Then Right1 = VBA.CLng(Image1.Width * (SYS≫Environment≫GetDpi() / 2540))
    If VBA.IsError(Bottom1) Then Bottom1 = VBA.CLng(Image1.Height * (SYS≫Environment≫GetDpi() / 2540))
    If VBA.IsError(Left2) Then Left2 = Left1
    If VBA.IsError(Top2) Then Top2 = Top1
    If VBA.IsError(Right2) Then Right2 = Right1
    If VBA.IsError(Bottom2) Then Bottom2 = Bottom1
    
    Dim img1() As Variant: img1 = IO≫Image≫GetPixelFromImage(Image1, Left1, Top1, Right1, Bottom1)
    Dim img2() As Variant: img2 = IO≫Image≫GetPixelFromImage(Image2, Left2, Top2, Right2, Bottom2)
    
    Dim i As Long: i = UBound(img1, 1)
    If i <> UBound(img2, 1) Then Exit Function
    
    Dim j As Long: j = UBound(img1, 2)
    If j <> UBound(img2, 2) Then Exit Function
    
    For i = i To 0 Step -1
        For j = j To 0 Step -1
            If img1(i, j) <> img2(i, j) Then Exit Function
        Next
    Next
    
    IO≫Image≫CompareImage = True
    
End Function

Public Function IO≫Image≫GetPixelFromImage(ByRef Picture As StdPicture, Optional ByVal Left As Long = 1, Optional ByVal Top As Long = 1, Optional ByVal Right, Optional ByVal Bottom) As Long()
    
    If VBA.IsError(Right) Then Right = VBA.CLng(Picture.Width * (SYS≫Environment≫GetDpi() / 2540))
    If VBA.IsError(Bottom) Then Bottom = VBA.CLng(Picture.Height * (SYS≫Environment≫GetDpi() / 2540))
    
    Dim hdc As Long
    hdc = CreateCompatibleDC(0&)
    
    Dim bmp1 As Long
    bmp1 = SelectObject(hdc, Picture.Handle)
    
    Dim bmp As bitmap
    GetObjectA Picture.Handle, Len(bmp), bmp
    
    Dim r() As Long
    ReDim r(Right - Left, Bottom - Top)
    
    Dim i As Long
    Dim j As Long
    For i = Left To Right
        For j = Top To Bottom
            r(i - Left, j - Top) = GetPixel(hdc, i - 1, j - 1)
        Next
    Next
    
    SelectObject hdc, bmp1
    DeleteDC hdc
    'DeleteObject Picture.Handle
    
    IO≫Image≫GetPixelFromImage = r
    
End Function

Public Function IO≫Image≫GetImageFromPicture(ByVal pic As Picture) As StdPicture
    pic.Copy
    Set IO≫Image≫GetImageFromPicture = IO≫Clipboard≫GetBitmapFromClipboard()
End Function


Public Function IO≫Image≫GetDocumentImage(ByVal Path As String, Optional ByVal Language As Long = &H11, Optional ByVal IsVertical As Boolean = False) As Variant()
    
    Dim InpWords() As Variant
    Dim OutWords() As Variant
    Dim ErrWords1() As Variant
    Dim ErrWords2() As Variant
    
    On Error GoTo Finally
    If Not IO≫File≫IsExists(Path) Then Error FileNotFoundException
    InpWords = GetDocumentImageStep1(Path, Language)
        
    GetDocumentImageStep2 InpWords, OutWords, ErrWords1, IsVertical, 1, DOCUMENT_IMAGE_ID, DOCUMENT_IMAGE_LINEID, DOCUMENT_IMAGE_LINEID, 0, DOCUMENT_IMAGE_REGIONID, DOCUMENT_IMAGE_REGIONID, 0
    
    If Not Not ErrWords1 Then
        Array≫Sort≫Sort2DArray ErrWords1, True, True, DOCUMENT_IMAGE_TOP
        GetDocumentImageStep2 ErrWords1, OutWords, ErrWords2, IsVertical, 0, VBA.IIf(IsVertical, DOCUMENT_IMAGE_TOP, DOCUMENT_IMAGE_LEFT), VBA.IIf(IsVertical, DOCUMENT_IMAGE_LEFT, DOCUMENT_IMAGE_TOP), VBA.IIf(IsVertical, DOCUMENT_IMAGE_RIGHT, DOCUMENT_IMAGE_BOTTOM), -1, VBA.IIf(IsVertical, DOCUMENT_IMAGE_RIGHT, DOCUMENT_IMAGE_BOTTOM), VBA.IIf(IsVertical, DOCUMENT_IMAGE_LEFT, DOCUMENT_IMAGE_TOP), 1
    End If
    
    Dim t As Variant
    If Not Not ErrWords2 Then
        For Each t In ErrWords2
            Array≫AddItem OutWords, t
        Next
    End If
    
Finally:
    IO≫Image≫Image≫GetDocumentImage = OutWords
    
End Function

Private Function GetDocumentImageStep1(ByVal Path As String, ByVal Language As Long) As Variant()
    
'    Const src As String = _
'        "CreateObject(""Scripting.FileSystemObject"").DeleteFile(Wscript.ScriptFullName):" & _
'        "Set d=CreateObject(""MODI.Document""):d.Create ""'ImagePath'"":d.OCR 'Language',0,0:If d.Images.Count=0 Then Wscript.Quit:End If:" & _
'        "With d.Images(0):Set l=.Layout:Redim v(l.Words.Count):n=1:t=vbTab:v(0)=.PixelWidth&t&.PixelHeight&t&l.Language:End With:" & _
'        "For each w In l.Words:With w.Rects(0):v(n)=w.Id&t&w.Text&t&.Left&t&.Right&t&.Top&t&.Bottom&t&w.LineId&t&w.RegionId:End With:n=n+1:Next:" & _
'        "GetObject(""'BookPath'"").Application.Run ""MultiSplit"",Join(v,vbNewLine),t:Set d=Nothing:Set l=d"
'
'    Dim params As New Collection
'    params.Add Path, "ImagePath"
'    params.Add Language, "Language"
'    params.Add ThisWorkbook.FullName, "BookPath"
'
'    Dim tmpPath As String
'    tmpPath =IO≫File≫CreateTempFile IO≫Path≫GetTempPath)
'    tmpPath = IO≫File≫Move(tmpPath, tmpPath & ".vbs")
'
'    IO≫File≫WriteText tmpPath, SetParameters(src, params)
'    Process≫Run tmpPath, 0, True
'
'    GetDocumentImageStep1 = TEMP_ARRAY
    
    With VBA.CreateObject("MODI.Document")
        
        Dim result() As Variant
        
        .Create Path
        .Ocr Language, 0, 0
        
        If .Images.Count > 0 Then
            
            With .Images(0)
                Array≫AddItem result, Array(.PixelWidth, .PixelHeight, .Layout.Language)
            End With
            
            Dim w As Variant
            For Each w In .Images(0).Layout.Words
                Array≫AddItem result, Array(w.id, w.Text, 0, 0, 0, 0, w.LineId, w.RegionId)
            Next
            
        End If
        
    End With
    
    GetDocumentImageStep1 = result
    
End Function

Private Sub GetDocumentImageStep2(ByRef InputWords() As Variant, ByRef OutputWords() As Variant, ByRef ErrorWords() As Variant, ByVal IsVertical As Boolean, ByVal StartIndex As Long, ByVal SortKeyNo As Long, ByVal GroupingKeyNo1A As Long, ByVal GroupingKeyNo1B As Long, ByVal GroupingCompareOperator1 As Integer, ByVal GroupingKeyNo2A As Long, ByVal GroupingKeyNo2B As Long, ByVal GroupingCompareOperator2 As Integer)
    
    Dim i As Long: i = StartIndex
    Dim cnt As Long: cnt = Array≫GetLength(InputWords) - 1
    
    Erase ErrorWords
    
    Do
        Dim t As Variant
        Dim tmp() As Variant
        
        Dim GroupingKey1 As Long: GroupingKey1 = InputWords(i)(GroupingKeyNo1A)
        Dim GroupingKey2 As Long: GroupingKey2 = InputWords(i)(GroupingKeyNo2A)
        
        ReDim tmp(0): tmp(0) = InputWords(i)
        Dim tmpLeft As Long: tmpLeft = InputWords(i)(DOCUMENT_IMAGE_LEFT)
        Dim tmpRight As Long: tmpRight = InputWords(i)(DOCUMENT_IMAGE_RIGHT)
        Dim tmpTop As Long: tmpTop = InputWords(i)(DOCUMENT_IMAGE_TOP)
        Dim tmpBottom As Long: tmpBottom = InputWords(i)(DOCUMENT_IMAGE_BOTTOM)
        Dim tmpLineId As Long: tmpLineId = InputWords(i)(DOCUMENT_IMAGE_LINEID)
        Dim tmpRegionId As Long: tmpRegionId = InputWords(i)(DOCUMENT_IMAGE_REGIONID)
        i = i + 1
        
        Do _
        While i <= cnt
            If _
            ( _
                Compare(GroupingKey1, InputWords(i)(GroupingKeyNo1B)) = GroupingCompareOperator1 And _
                Compare(GroupingKey2, InputWords(i)(GroupingKeyNo2B)) = GroupingCompareOperator2 _
            ) _
            Then
                If tmpLeft > InputWords(i)(DOCUMENT_IMAGE_LEFT) Then tmpLeft = InputWords(i)(DOCUMENT_IMAGE_LEFT)
                If tmpRight < InputWords(i)(DOCUMENT_IMAGE_RIGHT) Then tmpRight = InputWords(i)(DOCUMENT_IMAGE_RIGHT)
                If tmpTop > InputWords(i)(DOCUMENT_IMAGE_TOP) Then tmpTop = InputWords(i)(DOCUMENT_IMAGE_TOP)
                If tmpBottom < InputWords(i)(DOCUMENT_IMAGE_BOTTOM) Then tmpBottom = InputWords(i)(DOCUMENT_IMAGE_BOTTOM)
                Array≫AddItem tmp, InputWords(i)
                i = i + 1
            Else
                Exit Do
            End If
        Loop
        
        If ((tmpRight - tmpLeft) < (tmpBottom - tmpTop)) = IsVertical Then
            
            Array≫Sort≫Sort2DArray tmp, True, True, SortKeyNo
            
            Dim tmpText As String
            tmpText = vbNullString
            
            For Each t In tmp
                tmpText = tmpText & t(DOCUMENT_IMAGE_TEXT)
            Next
            
            Array≫AddItem _
                OutputWords, _
                Array(Array≫GetLength(OutputWords), tmpText, tmpLeft, tmpRight, tmpTop, tmpBottom, tmpLineId, tmpRegionId)
        Else
            
            For Each t In tmp
                
                Dim k As Long
                k = VBA.Len(t(DOCUMENT_IMAGE_TEXT))
                
                If k < 2 Then
                    Array≫AddItem ErrorWords, t
                Else
                    Dim tmpCharWidth As Long
                    Dim tmpCharHeight As Long

                    If IsVertical Then
                        tmpCharWidth = Int((t(DOCUMENT_IMAGE_RIGHT) - t(DOCUMENT_IMAGE_LEFT)) / k)
                        tmpCharHeight = t(DOCUMENT_IMAGE_BOTTOM) - t(DOCUMENT_IMAGE_TOP)
                    Else
                        tmpCharWidth = t(DOCUMENT_IMAGE_RIGHT) - t(DOCUMENT_IMAGE_LEFT)
                        tmpCharHeight = Int((t(DOCUMENT_IMAGE_BOTTOM) - t(DOCUMENT_IMAGE_TOP)) / k)
                    End If

                    For k = k To 1 Step -1
                        Array≫AddItem _
                            ErrorWords, _
                            Array _
                            ( _
                                t(DOCUMENT_IMAGE_ID), _
                                VBA.Mid(t(DOCUMENT_IMAGE_TEXT), k, 1), _
                                VBA.IIf(IsVertical, t(DOCUMENT_IMAGE_LEFT) + ((k - 1) * tmpCharWidth), t(DOCUMENT_IMAGE_LEFT)), _
                                VBA.IIf(IsVertical, t(DOCUMENT_IMAGE_LEFT) + (k * tmpCharWidth) - 1, t(DOCUMENT_IMAGE_RIGHT)), _
                                VBA.IIf(IsVertical, t(DOCUMENT_IMAGE_TOP), t(DOCUMENT_IMAGE_TOP) + ((k - 1) * tmpCharHeight)), _
                                VBA.IIf(IsVertical, t(DOCUMENT_IMAGE_BOTTOM), t(DOCUMENT_IMAGE_TOP) + (k * tmpCharHeight)) - 1, _
                                t(DOCUMENT_IMAGE_LINEID), _
                                t(DOCUMENT_IMAGE_REGIONID) _
                            )
                    Next
                End If
                
            Next
            
        End If
        
    Loop _
    While i <= cnt
        
End Sub

Public Function MultiSplit(ByVal Expression As String, Optional ByVal Delimiter1 As String = ",", Optional ByVal Delimiter2 As String = vbNewLine) As Variant
    
    Dim v As Variant
    v = VBA.Split(Expression, Delimiter2)
    
    Dim i As Long: i = UBound(v)
    ReDim TEMP_ARRAY(i)
    
    For i = i To 0 Step -1
        TEMP_ARRAY(i) = VBA.Split(v(i), Delimiter1)
    Next
    
    MultiSplit = TEMP_ARRAY
    
End Function

Private Function SetParameters(ByVal str As String, ByVal Parameters As VBA.Collection) As String
    
    Dim params As Variant
    Set params = GetParamSurroundedByQuotes(str)
    
    Dim i As Long
    For i = params.Count - 1 To 0 Step -1
        Dim Key As String
        Key = Replace(Mid(params(i).Value, 2, params(i).Length - 2), "''", "'")
        If Len(Key) > 0 Then
            str = Application.WorksheetFunction.Replace(str, params(i).FirstIndex + 1, params(i).Length, Parameters(Key))
        Else
            str = Application.WorksheetFunction.Replace(str, params(i).FirstIndex + 1, params(i).Length, "'")
        End If
    Next
    
    SetParameters = str
    
End Function

Private Function FilterValidation(ByVal obj As Object, ByVal Filter As String) As Boolean
    
    Static fil As Variant
    Static f As String
    
    If Filter <> f Then
        f = Filter
        Set fil = GetParamSurroundedByQuotes(Filter)
    End If
    
    Dim i As Long
    For i = fil.Count - 1 To 0 Step -1
        Dim Key As String
        Key = VBA.Replace(Mid(fil(i).Value, 2, fil(i).Length - 2), "''", "'")
        If VBA.Len(Key) > 0 Then
            Dim v As String: v = VBA.CallByName(obj, Key, VbGet)
            Filter = Application.WorksheetFunction.Replace(Filter, fil(i).FirstIndex + 1, fil(i).Length, VBA.IIf(IsNumeric(v), v, """" & v & """"))
        Else
            Filter = Application.WorksheetFunction.Replace(Filter, fil(i).FirstIndex + 1, fil(i).Length, "'")
        End If
    Next
    
    FilterValidation = Application.Evaluate(Filter)
    
End Function

Private Function GetParamSurroundedByQuotes(ByVal str As String) As Variant
    Const ESCAPE_SINGLE_QUOTE_PATTERN As String = "'([^']*('{2})*[^']*)*'"
    Set GetParamSurroundedByQuotes = IO≫Text≫Match(str, ESCAPE_SINGLE_QUOTE_PATTERN)
End Function

Public Function IO≫Clipboard≫GetBitmapFromClipboard() As StdPicture
    Dim i As Long
    i = GetBitmapHandleFromClipboard
    If i = 0 Then
        Set IO≫Clipboard≫GetBitmapFromClipboard = Nothing
    Else
        Set IO≫Clipboard≫GetBitmapFromClipboard = CreateBitmap(i)
    End If
End Function

Public Sub IO≫Clipboard≫SetClipboard(ByVal Text As String)
    With New DataObject
        .SetText Text
        .PutInClipboard
    End With
End Sub

Public Sub IO≫Clipboard≫ClearClipboard()
    ActiveCell.Copy
    Application.CutCopyMode = False
End Sub

Private Function GetBitmapHandleFromClipboard() As Long
    If OpenClipboard(0&) = 0 Then
        GetBitmapHandleFromClipboard = 0
    Else
        GetBitmapHandleFromClipboard = GetClipboardData(2)
        CloseClipboard
    End If
End Function

Public Sub UI≫Keyboard≫KeyDown(ParamArray Vk()): KeyDown Vk: End Sub
Private Sub KeyDown(ByVal Vk As Variant)
    Dim k As Variant
    For Each k In Vk
        keybd_event k, 0, 0, 0
    Next
End Sub

Public Sub UI≫Keyboard≫KeyUp(ParamArray Vk()): KeyUp Vk: End Sub
Private Sub KeyUp(ByVal Vk As Variant)
    Dim k As Variant
    For Each k In Vk
        keybd_event k, 0, 2, 0
    Next
End Sub

Public Sub UI≫Keyboard≫KeyPress(ParamArray Vk()): KeyPress Vk: End Sub
Public Sub KeyPress(ByVal Vk As Variant)
    Dim i As Long
    KeyDown Vk
    For i = UBound(Vk) To 0 Step -1
        keybd_event Vk(i), 0, 2, 0
    Next
End Sub

Public Sub UI≫Mouse≫MouseMove(Optional ByVal x As Long = 0, Optional ByVal Y As Long = 0)
    SetCursorPos x, Y
End Sub

Public Sub UI≫Mouse≫MouseDown(ByVal btn As MouseButton)
    mouse_event btn * 1, 0, 0, 0, 0
End Sub

Public Sub UI≫Mouse≫MouseUp(ByVal btn As MouseButton)
    mouse_event btn * 2, 0, 0, 0, 0
End Sub

Public Sub UI≫Mouse≫MouseClick(ByVal btn As MouseButton)
    UI≫Mouse≫MouseDown btn
    UI≫Mouse≫MouseUp btn
End Sub

Public Sub UI≫Mouse≫MouseDrag(ByVal btn As MouseButton, ByVal x As Long, ByVal Y As Long)
    UI≫Mouse≫MouseDown btn
    UI≫Mouse≫MouseMove x, Y
    UI≫Mouse≫MouseUp bnt
End Sub

Public Function UI≫Mouse≫GetMousePoint() As VBA.Collection
    
    Dim p As Point
    GetCursorPos p
    
    Dim c As New VBA.Collection
    c.Add p.x, "X"
    c.Add p.Y, "Y"
    
    Set UI≫Mouse≫GetMousePoint = c
    
End Function

Public Sub Util≫Wait(ByVal Milliseconds As Long)
    Application.Wait [Now()] + (Milliseconds / 86400000)
End Sub

Public Function UI≫Window≫Dialog≫SaveFileDialog(Optional ByVal InitialFileName As String = vbNullString, Optional ByVal FileFilter As String = VBA.vbNullString, Optional ByVal FilterIndex As Long = 1, Optional ByVal title As String = VBA.vbNullString, Optional ByVal ButtonText As String = VBA.vbNullString) As String
    UI≫Window≫Dialog≫SaveFileDialog = Application.GetSaveAsFilename(InitialFileName, FileFilter, FilterIndex, title, ButtonText)
End Function

Private Function Compare(ByVal a As Long, ByVal b As Long) As Integer
    If a = b Then
        Compare = 0
    ElseIf a > b Then
        Compare = 1
    ElseIf a < b Then
        Compare = -1
    End If
End Function

Private Function Inc(ByRef Num As Long) As Long
    Num = Num + 1
    Inc = Num
End Function

Private Function dec(ByRef Num As Long) As Long
    Num = Num - 1
    dec = Num
End Function

Private Function CallbackFunction4EnumWindows(ByVal hWnd As Long, lParam As Long) As Boolean
    TEMP_ARRAY.Add hWnd
    CallbackFunction4EnumWindows = True
End Function

Public Function UI≫Window≫IsVisible(ByVal hWnd As Long) As Boolean
    UI≫Window≫IsVisible = IsWindowVisible(hWnd)
End Function

Public Function UI≫Window≫GetWindows() As VBA.Collection
    
    Dim r As New VBA.Collection
    
    Dim p As Variant
    For Each p In SYS≫Process≫GetProcesses
        If UI≫Window≫IsVisible(p) Then
            If GetWindow(p, 4) = 0 Then
                
                Dim c As String
                c = UI≫Window≫GetWindowClass(p)
                
                If _
                ( _
                    c <> "Windows.UI.Core.CoreWindow" And _
                    c <> "ApplicationFrameWindow" And _
                    c <> "Progman" _
                ) _
                Then
                    r.Add p
                End If
                
            End If
        End If
    Next
    
    Set UI≫Window≫GetWindows = r
    
End Function

Public Function UI≫Window≫GetWindowTitle(ByVal hWnd As Long) As String
    
    Dim l As Long
    l = GetWindowTextLengthA(hWnd)
    
    Dim buf As String
    buf = VBA.String(l + 1, VBA.Chr$(0))
    
    If GetWindowTextA(hWnd, buf, VBA.Len(buf)) <> 0 Then
        UI≫Window≫GetWindowTitle = VBA.Replace(buf, VBA.Chr$(0), VBA.vbNullString)
    End If
    
End Function

Public Function UI≫Window≫GetWindowClass(ByVal hWnd As Long) As String
    
    Dim c As String
    c = VBA.String(128, VBA.Chr$(0))
    
    GetClassNameA hWnd, c, VBA.Len(c)
    UI≫Window≫GetWindowClass = VBA.Replace(c, VBA.Chr$(0), VBA.vbNullString)
    
End Function

Public Sub UI≫Window≫MoveWindow(ByVal hWndOrWindowTitle, ByVal x As Long, ByVal Y As Long)
    SetWindowPos UI≫Window≫GetWindowHandle(hWndOrWindowTitle), 0, x, Y, 0, 0, &H1 Or &H10
End Sub

Public Sub UI≫Window≫ResizeWindow(ByVal hWndOrWindowTitle, ByVal Width As Long, ByVal Height As Long)
    SetWindowPos UI≫Window≫GetWindowHandle(hWndOrWindowTitle), 0, 0, 0, Width, Height, &H2 Or &H10
End Sub

Public Function UI≫Window≫ActivateWindow(ByVal hWndOrWindowTitle) As Long
    Dim hWnd As Long: hWnd = UI≫Window≫GetWindowHandle(hWndOrWindowTitle)
    SetForegroundWindow hWnd
    UI≫Window≫ActivateWindow = hWnd
End Function

Public Sub UI≫Window≫SetWindowState(ByVal hWndOrWindowTitle, ByVal State As XlWindowState)
    
    Dim s As Long
    If State = XlWindowState.xlMinimized Then
        s = 7
    ElseIf State = XlWindowState.xlMaximized Then
        s = 3
    Else
        s = 4
    End If
    
    ShowWindow UI≫Window≫GetWindowHandle(hWndOrWindowTitle), s
    
End Sub

Public Function UI≫Window≫GetWindowState(ByVal hWndOrWindowTitle) As XlWindowState
    
    Dim hWnd As Long
    hWnd = UI≫Window≫GetWindowHandle(hWndOrWindowTitle)
    
    If IsIconic(hWnd) Then
        UI≫Window≫GetWindowState = XlWindowState.xlMinimized
    ElseIf IsZoomed(hWnd) Then
        UI≫Window≫GetWindowState = XlWindowState.xlMaximized
    Else
        UI≫Window≫GetWindowState = XlWindowState.xlNormal
    End If
    
End Function

Public Function UI≫Window≫GetWindowRectangle(ByVal hWnd As Long) As VBA.Collection
    
    Dim result As VBA.Collection
    Set result = New VBA.Collection
    
    Dim r As Rectangle
    GetWindowRect hWnd, r
    
    result.Add r.Left, "Left"
    result.Add r.Top, "Top"
    result.Add r.Right, "Right"
    result.Add r.Bottom, "Bottom"
    
    Set UI≫Window≫GetWindowRectangle = result
    
End Function

Public Function UI≫Window≫GetWindowFromPoint(ByVal x As Long, ByVal Y As Long) As Long
    
    Dim w As Variant
    For Each w In UI≫Window≫GetWindows
        
        If Not UI≫Window≫GetWindowState(w) = XlWindowState.xlMinimized Then
            
            Dim r As Collection
            Set r = UI≫Window≫GetWindowRectangle(w)
            
            If (r("Left") <= x And x <= r("Right")) Then
                If (r("Top") <= Y And Y <= r("Bottom")) Then
                    UI≫Window≫GetWindowFromPoint = w
                    Exit For
                End If
            End If
            
        End If
        
    Next
    
End Function

Public Function UI≫Window≫GetWindowHandle(ByVal windowTitle) As Long
    
    Dim t As VBA.VbVarType
    t = VBA.VarType(windowTitle)
    
    If _
    ( _
        t = VBA.VbVarType.vbInteger Or _
        t = VBA.VbVarType.vbLong _
    ) _
    Then
        UI≫Window≫GetWindowHandle = windowTitle
    Else
        
        Dim hWnd As Long: hWnd = FindWindowA(VBA.vbNullString, windowTitle)
        
        If hWnd = 0 Then
            UI≫Window≫GetWindowHandle = GetWindowHandleFromTitleRe(windowTitle)
        Else
            UI≫Window≫GetWindowHandle = hWnd
        End If
        
    End If
    
End Function

Public Function SYS≫Process≫GetProcesses() As VBA.Collection
    Set TEMP_ARRAY = Nothing
    Set TEMP_ARRAY = New VBA.Collection
    Call EnumWindows(AddressOf CallbackFunction4EnumWindows, 0)
    Set SYS≫Process≫GetProcesses = TEMP_ARRAY
End Function

Public Sub SYS≫Process≫Run(ByVal Path As String, Optional ByVal intWindowStyle As Long = 0, Optional ByVal bWaitOnReturn As Boolean = True)
    VBA.CreateObject("Wscript.Shell").Run """" & Path & """", intWindowStyle, bWaitOnReturn
End Sub

Public Function EscapeStr(ByVal str As String) As String
    
    Dim symbls As Variant
    symbls = Array("\", "*", "+", ".", "?", "{", "}", "(", ")", "[", "]", "^", "$", "|")
    
    Dim s As Variant
    For Each s In symbls
        str = Replace(str, s, "\" & s)
    Next
    
    EscapeStr = "^" & Replace(str, """", """""") & "$"
    
End Function

Public Function SYS≫Process≫GetProcessId(ByVal hWnd As Long) As Long
    Dim p As Long
    GetWindowThreadProcessId hWnd, p
    GetProcessIdFromHwnd = p
End Function

Public Function SYS≫Process≫GetProcess(ByRef WordObject As Object, ByVal ProcessId As Long) As Object
    
    Dim task As Object
    For Each task In WordObject.Tasks
        If task.Visible Then

            Dim h As Long
            Dim p As Long
            
            h = UI≫Window≫GetWindowHandle(task.Name)
            Call GetWindowThreadProcessId(h, p)
            
            If p = ProcessId Then
                Set SYS≫Process≫GetProcess = task
                Exit Function
            End If

        End If
    Next
    
    Set SYS≫Process≫GetProcess = Nothing
    
End Function

Function GetWindowHandleFromTitleRe(ByVal windowTitlePattern As String) As Long

    Dim re As Variant
    Dim hWnd As Variant
    
    GetWindowHandleFromTitleRe = 0
    Set re = CreateObject("VBScript.RegExp")
    
    For Each hWnd In SYS≫Process≫GetProcesses()
        
        Dim title As String
        title = Base.UI≫Window≫GetWindowTitle(hWnd)
        
        If Len(title) > 0 Then
            With re
                .Pattern = windowTitlePattern
                If .test(title) Then
                    GetWindowHandleFromTitleRe = hWnd
                    Exit For
                End If
            End With
        End If
        
    Next
    
End Function

Public Function IsKeyPressed(ByVal Key As Long) As Boolean
    Const KEY_PRESSED = -32768
    IsKeyPressed = (GetAsyncKeyState(Key) And KEY_PRESSED) = KEY_PRESSED
End Function

Public Function GetKeyCodes() As Variant
    Static KEYCODES As Variant
    
    If IsEmpty(KEYCODES) Then
        Set KEYCODES = CreateObject("Scripting.Dictionary")
        With KEYCODES
            .Add vbKeyA, "A": .Add vbKeyB, "B": .Add vbKeyC, "C": .Add vbKeyD, "D": .Add vbKeyE, "E": .Add vbKeyF, "F":
            .Add vbKeyG, "G": .Add vbKeyH, "H": .Add vbKeyI, "I": .Add vbKeyJ, "J": .Add vbKeyK, "K": .Add vbKeyL, "L":
            .Add vbKeyM, "M": .Add vbKeyN, "N": .Add vbKeyO, "O": .Add vbKeyP, "P": .Add vbKeyQ, "Q": .Add vbKeyR, "R":
            .Add vbKeyS, "S": .Add vbKeyT, "T": .Add vbKeyU, "U": .Add vbKeyV, "V": .Add vbKeyW, "W": .Add vbKeyX, "X":
            .Add vbKeyY, "Y": .Add vbKeyZ, "Z": .Add vbKey0, "0": .Add vbKey1, "1": .Add vbKey2, "2": .Add vbKey3, "3":
            .Add vbKey4, "4": .Add vbKey5, "5": .Add vbKey6, "6": .Add vbKey7, "7": .Add vbKey8, "8": .Add vbKey9, "9":
            .Add vbKeyF1, "F1 ": .Add vbKeyF2, "F2 ": .Add vbKeyF3, "F3 ": .Add vbKeyF4, "F4 ":
            .Add vbKeyF5, "F5 ": .Add vbKeyF6, "F6 ": .Add vbKeyF7, "F7 ": .Add vbKeyF8, "F8 ":
            .Add vbKeyF9, "F9 ": .Add vbKeyF10, "F10": .Add vbKeyF11, "F11": .Add vbKeyF12, "F12":
            .Add vbKeyF13, "F13": .Add vbKeyF14, "F14": .Add vbKeyF15, "F15": .Add vbKeyF16, "F16":
            .Add vbKeyNumpad0, "Numpad0": .Add vbKeyNumpad1, "Numpad1": .Add vbKeyNumpad2, "Numpad2": .Add vbKeyNumpad3, "Numpad3": .Add vbKeyNumpad4, "Numpad4":
            .Add vbKeyNumpad5, "Numpad5": .Add vbKeyNumpad6, "Numpad6": .Add vbKeyNumpad7, "Numpad7": .Add vbKeyNumpad8, "Numpad8": .Add vbKeyNumpad9, "Numpad9":
            .Add vbKeyLButton, "LButton": .Add vbKeyCancel, "Cancel": .Add vbKeyClear, "Clear": .Add vbKeyHome, "Home": .Add vbKeyAdd, "Add":
            .Add vbKeyRButton, "RButton": .Add vbKeyReturn, "Return": .Add vbKeyShift, "Shift": .Add vbKeyDown, "Down": .Add vbKeyTab, "Tab":
            .Add vbKeyMButton, "MButton": .Add vbKeyDivide, "Divide": .Add vbKeyPause, "Pause": .Add vbKeyLeft, "Left": .Add vbKeyEnd, "End":
            .Add vbKeyCapital, "Capital": .Add vbKeyEscape, "Escape": .Add vbKeySpace, "Space": .Add vbKeyHelp, "Help": .Add vbKeyUp, "Up":
            .Add vbKeyDecimal, "Decimal": .Add vbKeySelect, "Select": .Add vbKeyPrint, "Print": .Add vbKeyBack, "Back": .Add vbKeySnapshot, "Snapshot":
            .Add vbKeyControl, "Control": .Add vbKeyInsert, "Insert": .Add vbKeyRight, "Right": .Add vbKeyMenu, "Menu": .Add vbKeyPageDown, "PageDown":
            .Add vbKeyExecute, "Execute": .Add vbKeyPageUp, "PageUp": .Add vbKeyMultiply, "Multiply": .Add vbKeySeparator, "Separator":
            .Add vbKeyNumlock, "Numlock": .Add vbKeyDelete, "Delete": .Add vbKeySubtract, "Subtract":
        End With
    End If
    
    Set GetKeyCodes = KEYCODES
End Function

Public Function GetPressedKey() As Variant
    
    Dim keys As Variant: Set keys = GetKeyCodes()
    
    Dim k As Variant
    Dim i As Integer: i = 0
    Dim result As Variant: ReDim result(0)
    
    For Each k In keys
        If IsKeyPressed(k) Then
            ReDim Preserve result(i)
            result(UBound(result)) = k
            i = i + 1
        End If
    Next
    
    GetPressedKey = result
    
End Function

Public Function GetDesktopSize() As VBA.Collection

    Dim result As VBA.Collection
    Set result = New VBA.Collection
    
    result.Add 0, "Left"
    result.Add 0, "Top"
    result.Add GetSystemMetrics(0), "Right"
    result.Add GetSystemMetrics(1), "Bottom"
    
    Set GetDesktopSize = result
    
End Function



Public Function Array≫GetLength(ByRef ary As Variant, Optional Dimension As Long = 1) As Long
On Error GoTo Exception
    Array≫GetLength = UBound(ary) + 1
    Exit Function
Exception:
    If Err.Number = IndexOutOfRangeException Then
        Array≫GetLength = 0
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function

Public Function Array≫GetDimension4MultiDimension(ByRef ary As Variant) As Long
    On Error Resume Next
    Dim i As Long
    Dim j As Long
    
    Do
        i = i + 1
        j = UBound(ary, i)
    Loop _
    While Err.Number = 0
    
    On Error GoTo 0
    Array≫GetDimension4MultiDimension = i - 1
    
End Function

Public Function Array≫GetDimension4JaggedArray(ByRef ary As Variant) As Long
    Dim i As Long
    Dim v1 As Variant: SetAnyTypeObject v1, ary
    Dim v2 As Variant
    Do
        i = i + 1
        SetAnyTypeObject v2, v1(0)
        SetAnyTypeObject v1, v2
    Loop _
    While (VarType(v2) And vbArray) = vbArray
    Array≫GetDimension4JaggedArray = i
End Function

Public Function Array≫AddItem(ByRef ary As Variant, ByRef item As Variant, Optional ByVal Index As Long = -1) As Long
    If (VBA.VarType(ary) And VBA.VbVarType.vbArray) = VBA.VbVarType.vbArray Then
        
        Dim i As Long
        i = Array≫GetLength(ary)
        
        ReDim Preserve ary(i)
        
        If Index >= 0 Then
            For i = i To Index + 1 Step -1
                SetAnyTypeObject ary(i), ary(i - 1)
            Next
            i = Index
        End If
        
        SetAnyTypeObject ary(i), item
        Array≫AddItem = i
        
    Else
        Error InvalidTypeException
    End If
End Function

Public Sub Array≫RemoveItem(ByRef ary As Variant, ByVal Index As Long)
    
    Dim i As Long
    Dim cnt As Long: cnt = Array≫GetLength(ary) - 1
    
    If cnt < 1 Then
        Erase ary
    Else
        For i = Index To cnt - 1
            SetAnyTypeObject ary(i), ary(i + 1)
        Next
        ReDim Preserve ary(cnt - 1)
    End If
    
End Sub

Public Sub Array≫ReplaceArrayElement(ByRef ary As Variant, ByVal idx1 As Long, ByVal idx2 As Long)
    If idx1 <> idx2 Then
        Dim tmp As Variant
        If Array≫IsCollection(ary) Then
            If idx1 < idx2 Then
                tmp = idx1
                idx1 = idx2
                idx2 = tmp
            End If
            ary.Add item:=ary(idx2), After:=idx1
            ary.Add item:=ary(idx1), After:=idx2
            ary.Remove idx2
            ary.Remove idx1
        Else
            SetAnyTypeObject tmp, ary(idx1)
            SetAnyTypeObject ary(idx1), ary(idx2)
            SetAnyTypeObject ary(idx2), tmp
        End If
    End If
End Sub

Public Function Array≫Copy2Array(ByRef ary As Variant) As Variant()
    Dim newAry() As Variant
    Dim v As Variant
    For Each v In ary
        Array≫AddItem newAry, v
    Next
    Array≫Copy2Array = newAry
End Function

Public Function Array≫Copy2Collection(ByRef ary As Variant) As VBA.Collection
    Dim c As New VBA.Collection
    Dim v As Variant
    For Each v In ary
        c.Add v
    Next
    Set Array≫Copy2Collection = c
End Function

Public Function Array≫SearchArray(ByRef ary As Variant, ByVal Filter As String, Optional ByVal Dimension As Long = 0, Optional ByVal FindFirst As Boolean = False) As Variant()
    
    Dim a As Variant
    Dim v As Variant
    Dim r() As Variant
    Dim t() As Variant
    
    t = ary
    If Not Not t Then
        For Each a In t
            
            If Dimension > 0 Then
                v = a(Dimension)
            Else
                v = a
            End If
            
            If IO≫Text≫Match(v, Filter, True).Count > 0 Then
                Array≫AddItem r, a
                If FindFirst Then Exit For
            End If
            
        Next
    End If
    
    Array≫SearchArray = r
    
End Function

Public Function Array≫Exists(ByRef ary As Variant, ByVal Filter As String, Optional ByVal Dimension As Long = 0) As Boolean
    If Array≫GetLength(ary) > 0 Then
        Array≫Exists = Array≫GetLength(Array≫SearchArray(ary, Filter, Dimension, True)) > 0
    Else
        Array≫Exists = False
    End If
End Function

Public Function Array≫IsCollection(ByRef ary As Variant) As Boolean
    Array≫IsCollection = (VBA.TypeName(ary) = "Collection")
End Function

Public Sub Array≫Sort≫SortArray(ByRef ary As Variant, Optional ByVal AscendingSort As Boolean = True, Optional ByVal SortOnValues As Boolean = True)
    SortArrayStep1 ary, False, False, AscendingSort, SortOnValues, Array(1)
End Sub

Public Sub Array≫Sort≫Sort2DArray(ByRef ary As Variant, ByVal AscendingSort As Boolean, ByVal SortOnValues As Boolean, ParamArray SortKey())
    SortArrayStep1 ary, False, True, AscendingSort, SortOnValues, GetParamArray(SortKey)
End Sub

Public Sub Array≫Sort≫SortArrayByElementProperty(ByRef ary As Variant, ByVal AscendingSort As Boolean, ByVal SortOnValues As Boolean, ParamArray SortKey())
    SortArrayStep1 ary, True, False, AscendingSort, SortOnValues, SortKey
End Sub

Public Sub Array≫Sort≫Sort2DCollection(ByRef ary As Collection, ByVal AscendingSort As Boolean, ByVal SortOnValues As Boolean, ParamArray SortKey())
    SortArrayStep2 ary, VBA.VbVarType.vbArray, keys, 1, ary.Count, AscendingSort, SortOnValues
End Sub

Public Sub SetAnyTypeObject(ByRef a As Variant, ByRef b As Variant)
    If VBA.IsObject(b) Then
        Set a = b
    Else
        a = b
    End If
End Sub

Private Sub SortArrayStep1(ByRef ary As Variant, ByVal flgIsObject As Boolean, ByVal flgIsArray As Boolean, ByVal AscendingSort As Boolean, ByVal SortOnValues As Boolean, ByVal keys As Variant)
    If UBound(keys) < 0 Then Exit Sub
    
    Dim lMin As Long
    Dim lMax As Long
    
    If (VBA.VarType(ary) And VBA.VbVarType.vbArray) = VBA.VbVarType.vbArray Then
        lMin = LBound(ary)
        lMax = UBound(ary)
    ElseIf Array≫IsCollection(ary) Then
        lMin = 1
        lMax = ary.Count
    Else
        Error InvalidTypeException
    End If
    
    If IsObject(ary(lMin)) = flgIsObject Then
        
        Dim ty As VbVarType
        If Array≫IsCollection(ary(lMin)) Then
            ty = VBA.VbVarType.vbArray
        Else
            ty = VBA.VarType(ary(lMin))
        End If
        
        If ((ty And VBA.VbVarType.vbArray) = VBA.VbVarType.vbArray) = flgIsArray Then
            SortArrayStep2 ary, ty, keys, lMin, lMax, AscendingSort, SortOnValues
            Exit Sub
        End If
        
    End If
    
    Error InvalidTypeException
    
End Sub

Private Sub SortArrayStep2(ByRef ary As Variant, ByVal ty As VbVarType, ByRef keys As Variant, ByVal lMin As Long, ByVal lMax As Long, ByVal AscendingSort As Boolean, ByVal SortOnValues As Boolean)
    
    Dim idx1 As Long: idx1 = lMin
    Dim idx2 As Long: idx2 = lMax
    
    Dim v1 As Variant: v1 = ary((lMin + lMax) \ 2)
    Dim val() As Variant: val = SortArrayStep3(v1, ty, keys, SortOnValues)
    
    Do
        Dim v2 As Variant: v2 = ary(idx1)
        Do While SortArrayStep4(SortArrayStep3(v2, ty, keys, SortOnValues), val, AscendingSort): idx1 = idx1 + 1: v2 = ary(idx1): Loop
        
        Dim v3 As Variant: v3 = ary(idx2)
        Do While SortArrayStep4(val, SortArrayStep3(v3, ty, keys, SortOnValues), AscendingSort): idx2 = idx2 - 1: v3 = ary(idx2): Loop
        
        If idx1 < idx2 Then
            Array≫ReplaceArrayElement ary, idx1, idx2
            idx1 = idx1 + 1
            idx2 = idx2 - 1
        Else
            Exit Do
        End If
        
    Loop
    
    If lMin < idx1 - 1 Then SortArrayStep2 ary, ty, keys, lMin, idx1 - 1, AscendingSort, SortOnValues
    If idx2 + 1 < lMax Then SortArrayStep2 ary, ty, keys, idx1, lMax, AscendingSort, SortOnValues
    
End Sub

Private Function SortArrayStep3(ByRef obj As Variant, ByVal ty As VBA.VbVarType, ByRef keys As Variant, ByVal SortOnValues As Boolean) As Variant()
    Dim result() As Variant
    Dim Key As Variant
    For Each Key In keys
        Array≫AddItem result, GetAnonymousTypeValue(obj, ty, Key, SortOnValues)
    Next
    SortArrayStep3 = result
End Function

Private Function SortArrayStep4(ByRef ary1() As Variant, ByRef ary2() As Variant, ByVal AscendingSort As Boolean) As Boolean
    Dim i As Long
    SortArrayStep4 = False
    For i = 0 To UBound(ary1)
        If ary1(i) = ary2(i) Then
        ElseIf (ary1(i) < ary2(i)) = AscendingSort Then
            SortArrayStep4 = True
            Exit Function
        ElseIf (ary1(i) > ary2(i)) <> AscendingSort Then
            Exit Function
        End If
    Next
End Function

Private Function GetAnonymousTypeValue(ByRef obj As Variant, ByVal ty As VbVarType, ByVal Key As Variant, ByVal SortOnValues As Boolean) As Variant
    Dim v As Variant
    If ty = VBA.VbVarType.vbObject Then
        v = VBA.CallByName(obj, Key, VbGet)
    ElseIf (ty And VBA.VbVarType.vbArray) = VBA.VbVarType.vbArray Then
        v = obj(Key)
    Else
        v = obj
    End If
    If VBA.IsNumeric(v) And SortOnValues Then
        GetAnonymousTypeValue = VBA.CDbl(v)
    Else
        GetAnonymousTypeValue = v
    End If
End Function

Public Function UI≫Accessible≫GetIAccessibleFromPoint(ByVal x As Long, ByVal Y As Long) As Variant
    
    Dim acc As Object
    Dim child As Variant
    
    #If Win64 Then
        Call AccessibleObjectFromPoint(y * &H100000000^ Or x, acc, child)
    #Else
        Call AccessibleObjectFromPoint(x, Y, acc, child)
    #End If
    
    Set UI≫Accessible≫GetIAccessibleFromPoint = VBA.IIf(VBA.IsObject(acc), acc, Nothing)
    
End Function

Public Function Accessible≫GetIAccessible(ByVal hWnd As Long) As Variant
    Dim acc As Variant
    Set acc = GetIAccessibleFromWindow(hWnd)
    Set Accessible≫GetIAccessible = VBA.IIf(VBA.IsObject(acc), acc, Nothing)
End Function

Private Function GetIAccessibleFromWindow(ByVal hWnd As Long) As Variant
    
    Static IID_IAccessible As UUID
    If Not Not IID_IAccessible Then
        With IID_IAccessible
            .Data1 = &H618736E0
            .Data2 = &H3C3D
            .Data3 = &H11CF
            .Data4(0) = &H81: .Data4(4) = &H0
            .Data4(1) = &HC:  .Data4(5) = &H38
            .Data4(2) = &H0:  .Data4(6) = &H9B
            .Data4(3) = &HAA: .Data4(7) = &H71
        End With
    End If

    Dim acc As Object
    Call AccessibleObjectFromWindow(hWnd, &HFFFFFFFC, IID_IAccessible, acc)
    Set GetIAccessibleFromWindow = VBA.IIf(VBA.IsObject(acc), acc, Nothing)
    
End Function

Public Function Accessible≫GetIAccessibleFromProcess(Optional ByVal windowTitle As String = VBA.vbNullString, Optional ByVal Filter As String = VBA.vbNullString, Optional ByVal FindFirst As Boolean = False) As Variant
On Error GoTo Finally
    
    Dim oWord As Object
    Set oWord = VBA.CreateObject("Word.Application")
    
    Dim taskObjects() As Variant
    taskObjects = GetProcessFromWindowTitle(oWord, windowTitle, Filter, FindFirst)

    If FindFirst Then
        Set Accessible≫GetIAccessibleFromProcess = VBA.IIf(Not Not taskObjects, GetIAccessibleFromWindow(UI≫Window≫GetWindowHandle(taskObjects(0).Name)), Nothing)
    Else
        Dim result() As Variant
        Dim task As Variant
        If Not Not taskObjects Then
            For Each task In taskObjects
                Array≫AddItem result, GetIAccessibleFromWindow(UI≫Window≫GetWindowHandle(task.Name))
            Next
        End If
        Accessible≫GetIAccessibleFromProcess = result
    End If

Finally:
    If Not oWord Is Nothing Then oWord.Quit
    If VBA.Err.Number <> 0 Then VBA.Err.Raise VBA.Err.Number, VBA.Err.Source, VBA.Err.Description, VBA.Err.HelpFile, VBA.Err.HelpContext
    
End Function

Private Function GetProcessFromWindowTitle(ByRef WordObject As Object, ByVal windowTitle As String, ByVal Filter As String, ByVal FindFirst As Boolean) As Variant()

    Dim result() As Variant
    Dim task As Object

    For Each task In WordObject.Tasks
        If task.Visible Then

            Dim filFlg As Boolean: filFlg = True
            If VBA.Len(windowTitle) > 0 Then filFlg = Not Application.IsError(Application.Search(windowTitle, task.Name))
            If VBA.Len(Filter) > 0 Then filFlg = filFlg And FilterValidation(task, Filter)

            If filFlg Then
                Array≫AddItem result, task
                If FindFirst Then Exit For
            End If

        End If
    Next

    GetProcessFromWindowTitle = result

End Function

Public Function Accessible≫GetChildObject(ByRef acc As Variant, Optional ByVal ObjectName As String = VBA.vbNullString, Optional ByVal Filter As String = VBA.vbNullString, Optional ByVal FindFirst As Boolean = False) As Variant

    Dim childObjects As Variant
    childObjects = GetChildObjectsFromIAccessible(acc, ObjectName, Filter, FindFirst)

    If FindFirst Then
        Set Accessible≫GetChildObject = VBA.IIf(Not Not childObjects, childObjects(0), Nothing)
    Else
        Dim result() As Variant
        Dim child As Variant
        If Not Not childObjects Then
            For Each child In childObjects
                Array≫AddItem result, child
            Next
        End If
        Accessible≫GetChildObject = result
    End If

End Function

Private Function GetChildObjectsFromIAccessible(ByRef acc As Variant, ByVal ObjectName As String, ByVal Filter As String, ByVal FindFirst As Boolean) As Variant()
On Error GoTo Finally
    
    Dim result() As Variant
    If Not VBA.IsNull(acc.accName) Then On Error GoTo 0

    Dim filFlg As Boolean: filFlg = True
    If VBA.Len(ObjectName) > 0 Then filFlg = Not Application.IsError(Application.Search(ObjectName, acc.accName))
    If VBA.Len(Filter) > 0 Then filFlg = filFlg And FilterValidation(acc, Filter)

    If filFlg Then
        Array≫AddItem result, acc
        If FindFirst Then GoTo Finally
    End If

    Dim childCount As Long
    childCount = acc.accChildCount

    If childCount > 0 Then

        Dim childObjects() As Variant
        ReDim childObjects(childCount - 1) As Variant
        Call AccessibleChildren(acc, 0, childCount, childObjects(0), childCount)

        Dim childObject As Variant
        For Each childObject In childObjects
            
            Dim grandChildObjects() As Variant
            Dim grandChildObject As Variant
            grandChildObjects = GetChildObjectsFromIAccessible(childObject, ObjectName, Filter, FindFirst)
            
            If Not Not grandChildObjects Then
                For Each grandChildObject In grandChildObjects
                    If VBA.IsObject(grandChildObject) Then
                        Array≫AddItem result, grandChildObject
                        If FindFirst Then GoTo Finally
                    End If
                Next
            End If
            
        Next

    End If

Finally:
    GetChildObjectsFromIAccessible = result

End Function

Public Function Accessible≫GetRoleSystemName(ByVal accRoleNo As Long) As String
    Accessible≫GetRoleSystemName = _
    VBA.Choose _
    ( _
        accRoleNo, "TITLEBAR", "MENUBAR", "SCROLLBAR", "GRIP", "SOUND", "CURSOR", "CARET", "ALERT", "WINDOW", "CLIENT", _
        "MENUPOPUP", "MENUITEM", "TOOLTIP", "APPLICATION", "DOCUMENT", "PANE", "CHART", "DIALOG", "BORDER", "GROUPING", _
        "SEPARATOR", "TOOLBAR", "STATUSBAR", "TABLE", "COLUMNHEADER", "ROWHEADER", "COLUMN", "ROW", "CELL", "LINK", _
        "HELPBALLOON", "CHARACTER", "LIST", "LISTITEM", "OUTLINE", "OUTLINEITEM", "PAGETAB", "PROPERTYPAGE", "INDICATOR", _
        "GRAPHIC", "STATICTEXT", "TEXT", "PUSHBUTTON", "CHECKBUTTON", "RADIOBUTTON", "COMBOBOX", "DROPLIST", "PROGRESSBAR", _
        "DIAL", "HOTKEYFIELD", "SLIDER", "SPINBUTTON", "DIAGRAM", "ANIMATION", "EQUATION", "BUTTONDROPDOWN", "BUTTONMENU", _
        "BUTTONDROPDOWNGRID", "WHITESPACE", "PAGETABLIST", "CLOCK", "SPLITBUTTON", "IPADDRESS" _
    )
End Function

Public Sub Accessible≫SetFocus(ByRef acc As Variant)
    Dim v As Variant
    acc.accSelect 0, v
End Sub

Public Function Accessible≫RunAndGetIAccessible(ByVal FilePath As String, Optional ByVal WindowStyle As VbAppWinStyle = VbAppWinStyle.vbMinimizedFocus) As Variant
On Error GoTo Finally
    
    Dim p As Long
    Dim h As Long
    Dim t As Variant
    
    Dim oWord As Object
    Set oWord = VBA.CreateObject("Word.Application")

    p = VBA.Shell(FilePath, WindowStyle)
    Set t = SYS≫Process≫GetProcess(oWord, p)
    h = Window≫GetWindowHandle(t.Name)

    Dim acc As Variant
    Set acc = GetIAccessibleFromWindow(h)

    Set Accessible≫RunAndGetIAccessible = VBA.IIf(acc Is Nothing, Nothing, acc)
    
Finally:
    If Not oWord Is Nothing Then oWord.Quit
    If VBA.Err.Number <> 0 Then VBA.Err.Raise VBA.Err.Number, VBA.Err.Source, VBA.Err.Description, VBA.Err.HelpFile, VBA.Err.HelpContext
    
End Function

Public Sub DB≫CreateAccessDatabaseFile(ByVal Path As String)
    VBA.CreateObject("ADOX.Catalog").Create GetConnectionString(Path)
End Sub

Public Function DB≫CreateDataColumn(ByVal ColumnName As String, ByVal DataType As VbVarType, Optional ByVal DataSize As Long = 0, Optional ByVal Nullable As Boolean = True) As Object
    Dim col As Object
    Set col = VBA.CreateObject("ADOX.Column")
    col.Name = ColumnName
    col.Type = DataType
    If DataSize > 0 Then col.DefinedSize = DataSize
    If Not Nullable Then col.Attributes = col.Attributes Or 2
    Set DB≫CreateDataColumn = col
End Function

Public Function DB≫CreateUniqueKey(ByVal KeyName As String, ByVal Primary As Boolean, ParamArray Columns()) As Object
    
    Dim uk As Object
    Set uk = VBA.CreateObject("ADOX.Key")
    uk.Name = KeyName
    uk.Type = VBA.IIf(Primary, 1, 3)
    
    Dim c As Variant
    For Each c In Columns
        uk.Columns.Append c
    Next
    
    Set DB≫CreateUniqueKey = uk
    
End Function

Public Function DB≫CreateForeignKey(ByVal KeyName As String, ByVal KeyColumnName As String, ByVal RelatedTableName As String, ByVal RelatedColumnName As String, ByVal Cascade As Boolean) As Object
    Dim fk As Object
    Set fk = VBA.CreateObject("ADOX.Key")
    fk.Name = KeyName
    fk.Type = 2
    fk.Columns.Append KeyColumnName
    fk.RelatedTable = RelatedTableName
    fk.Columns(KeyColumnName).RelatedColumn = RelatedColumnName
    fk.UpdateRule = IIf(Cascade, 1, 0)
    fk.DeleteRule = IIf(Cascade, 1, 0)
    Set DB≫CreateForeignKey = fk
End Function

Public Function DB≫CreateDataTable(ByVal TableName As String, Optional ByRef Columns As Variant = vbEmpty, Optional ByRef keys As Variant = vbEmpty) As Object
    
    Dim tbl As Object
    Set tbl = VBA.CreateObject("ADOX.Table")
    tbl.Name = TableName
    
    If VBA.IsArray(Columns) Then
        Dim c As Variant
        For Each c In Columns
            tbl.Columns.Append c
        Next
    End If
    
    If VBA.IsArray(keys) Then
        Dim k As Variant
        For Each k In keys
            tbl.keys.Append k
        Next
    End If
    
    Set DB≫CreateDataTable = tbl
    
End Function

Public Sub DB≫AddDataTable(ByRef Connection As Object, ByRef DataTable As Object, Optional ByVal Temporary As Boolean = False)
    DB≫Execute Connection, GetCreateTableSql(DataTable)
End Sub

Public Sub DB≫DeleteDataTable(ByRef Connection As Object, ByVal TableName As String)
    Const DROP_TABLE_STRING As String = "DROP TABLE "
    DB≫Execute Connection, DROP_TABLE_STRING & TableName
End Sub

Public Sub DB≫AddDataColumn(ByRef Connection As Object, ByVal TableName As String, ByRef DataColumn As Variant)
    Const ADD_COLUMN_STRING As String = "ALTER TABLE 'TABLE_NAME' ADD COLUMN 'NEW_COLUMN'"
    Dim params As New VBA.Collection
    params.Add TableName, "TABLE_NAME"
    params.Add GetQueryFromColumnObject(DataColumn), "NEW_COLUMN"
    DB≫Execute Connection, SetParameters(ADD_COLUMN_STRING, params)
End Sub

Public Sub DB≫AlterDataColumn(ByRef Connection As Object, ByVal TableName As String, ByRef DataColumn As Variant)
    Const ALTER_COLUMN_STRING As String = "ALTER TABLE 'TABLE_NAME' ALTER COLUMN 'NEW_COLUMN'"
    Dim params As New VBA.Collection
    params.Add TableName, "TABLE_NAME"
    params.Add GetQueryFromColumnObject(DataColumn), "NEW_COLUMN"
    DB≫Execute Connection, SetParameters(ALTER_COLUMN_STRING, params)
End Sub

Public Sub DB≫DeleteDataColumn(ByRef Connection As Object, ByVal TableName As String, ByVal ColumnName As String)
    Const DELETE_COLUMN_STRING As String = "ALTER TABLE 'TABLE_NAME' DROP COLUMN 'NEW_COLUMN'"
    Dim params As New VBA.Collection
    params.Add TableName, "TABLE_NAME"
    params.Add ColumnName, "NEW_COLUMN"
    DB≫Execute Connection, SetParameters(DELETE_COLUMN_STRING, params)
End Sub

Private Function GetCreateTableSql(ByRef DataTable As Object, Optional ByVal Temporary As Boolean = False) As String
    
    Const CREATE_TABLE_STRING As String = "CREATE 'TEMPORARY' TABLE 'TABLE_NAME'('FIELDS')"
    
    Dim i As Variant
    Dim f() As String
    
    For Each i In DataTable.Columns
        Array≫AddItem f, GetQueryFromColumnObject(i)
    Next
    For Each i In DataTable.keys
        Array≫AddItem f, GetQueryFromKeyObject(i)
    Next
    
    Dim params As New Collection
    params.Add VBA.IIf(Temporary, "TEMPORARY", VBA.vbNullString), "TEMPORARY"
    params.Add DataTable.Name, "TABLE_NAME"
    params.Add VBA.Join(f, ","), "FIELDS"
    
    GetCreateTableSql = SetParameters(CREATE_TABLE_STRING, params)
    
End Function

Private Function GetQueryFromColumnObject(ByRef DataColumn As Variant) As String
    
    Dim strDataType As String
    strDataType = VBA.Choose(DataColumn.Type, "COUNTER", "SMALLINT", "INTEGER", "REAL", "FLOAT", "MONEY", "DATETIME", "TEXT", "TEXT", "BIT")
    
    Dim strDataSize As String
    If _
    ( _
        DataColumn.Type = VBA.VbVarType.vbString And _
        DataColumn.DefinedSize > 0 And _
        DataColumn.DefinedSize < &H100 _
    ) _
    Then
        strDataSize = "(" & DataColumn.DefinedSize & ")"
    Else
        strDataSize = VBA.vbNullString
    End If
    
    Dim strNullable As String
    strNullable = VBA.IIf((DataColumn.Attributes And 2) = 2, "NOT NULL", VBA.vbNullString)
    
    GetQueryFromColumnObject = VBA.Join(Array(DataColumn.Name, strDataType & strDataSize, strNullable), VBA.Space(1))
    
End Function

Private Function GetQueryFromKeyObject(ByRef DataKey As Variant) As String
    
    Const DATA_KEY_STRING As String = "CONSTRAINT 'KEY_NAME' 'KEY_TYPE' ('COLUMNS')"
    Const REFERENCES_KEY_STRING As String = "REFERENCES 'RELATED_TABLE_NAME'('RELATED_COLUMN_NAME')"
    'Const REFERENCES_KEY_STRING As String = "REFERENCES 'RELATED_TABLE_NAME'('RELATED_COLUMN_NAME') ON UPDATE 'UPDATE_CASCADE' ON DELETE 'DELETE_CASCADE'"
    
    Dim dataKeyString As String
    dataKeyString = DATA_KEY_STRING
    
    Dim params As New VBA.Collection
    params.Add DataKey.Name, "KEY_NAME"
    params.Add VBA.Join(Array≫Copy2Array(DataKey.Columns), ","), "COLUMNS"
    
    If DataKey.Type = 1 Then
        params.Add "PRIMARY KEY", "KEY_TYPE"
    ElseIf DataKey.Type = 2 Then
        dataKeyString = dataKeyString & Space(1) & REFERENCES_KEY_STRING
        params.Add "FOREIGN KEY", "KEY_TYPE"
        params.Add DataKey.RelatedTable, "RELATED_TABLE_NAME"
        params.Add DataKey.Columns(DataKey.Columns(0).Name).RelatedColumn, "RELATED_COLUMN_NAME"
        'params.Add IIf(DataKey.UpdateRule = 1, "CASCADE", "SET NULL"), "UPDATE_CASCADE"
        'params.Add IIf(DataKey.DeleteRule = 1, "CASCADE", "SET NULL"), "DELETE_CASCADE"
    ElseIf DataKey.Type = 3 Then
        params.Add "UNIQUE", "KEY_TYPE"
    Else
        VBA.Error ArgumentOutOfRangeException
    End If
    
    GetQueryFromKeyObject = SetParameters(dataKeyString, params)
    
End Function

Public Function DB≫GetSession(ByVal Provider As Providers, ByVal DataSource As String, ByVal User As String, ByVal Password As String, Optional ByVal ReadOnly As Boolean = False) As Object
    
    Dim con As Object
    Set con = VBA.CreateObject("ADODB.Connection")
    
    Dim params As New VBA.Collection
    params.Add DataSource, "DATASOURCE"
    params.Add User, "USER"
    params.Add Password, "PASSWORD"
    params.Add ReadOnly, "READONLY"
    
    con.ConnectionString = SetParameters(GetConnectionString(Provider), params)
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataSource & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
    con.Open
    
    Set DB≫GetSession = con
    
End Function

Public Function DB≫ReleaseSession(ByRef Connection As Object)
    Connection.Close
    Set Connection = Nothing
End Function

Public Function DB≫BeginTrans(ByRef Connection As Object) As Boolean
On Error GoTo Exception
    Connection.BeginTrans
    DB≫BeginTrans = True
    Exit Function
Exception:
    If VBA.Err.Number = FeatureNotAvailableException Then
        DB≫BeginTrans = False
    Else
        VBA.Err.Raise VBA.Err.Number, VBA.Err.Source, VBA.Err.Description, VBA.Err.HelpFile, VBA.Err.HelpContext
    End If
End Function

Public Sub DB≫Commit(ByRef Connection As Object)
    Connection.CommitTrans
End Sub

Public Sub DB≫Rollback(ByRef Connection As Object)
    Connection.RollbackTrans
End Sub

Public Function DB≫CreateSqlCommand(ByRef Connection As Object, ByVal SQL As String, ByRef Parameters() As Variant, Optional ByVal Timeout As Long = 0) As Object
    
    Dim cmd As Object
    Set cmd = VBA.CreateObject("ADODB.Command")
    
    cmd.ActiveConnection = Connection
    cmd.CommandType = 1
    cmd.CommandText = SQL
    cmd.CommandTimeout = Timeout
    
    cmd.Parameters.Append cmd.CreateParameter("Param", 11, 2)
    cmd.Parameters.Delete 0
    
    Dim i As Long
    For i = 0 To UBound(Parameters)
        Dim strPara As String: strPara = VBA.CStr(Parameters(i))
        If VBA.LenB(strPara) = 0 Then
            cmd.Parameters.Append cmd.CreateParameter("Param" & i, 200, 2, 4000)
        Else
            cmd.Parameters.Append cmd.CreateParameter("Param" & i, 200, 1, VBA.LenB(strPara), strPara)
        End If
    Next
    
    Set DB≫CreateSqlCommand = cmd
    
End Function

Public Function DB≫Execute(ByRef Connection As Object, ByVal SQL As String, ParamArray Parameters()) As Variant
    Dim Args() As Variant: Args = Parameters
    Set DB≫Execute = DB≫ExecuteSqlCommand(DB≫CreateSqlCommand(Connection, SQL, Args))
End Function

Public Function DB≫ExecuteSqlCommand(ByRef Command As Object) As Object
    Dim rec As Variant
    Set rec = Command.Execute
    If rec.State = 0 Then
        Set DB≫ExecuteSqlCommand = Array≫Copy2Collection(Command.Parameters)
    Else
        Set DB≫ExecuteSqlCommand = rec
    End If
End Function

Private Function GetConnectionString(ByVal Provider As Providers) As String
    Select Case Provider
        Case Providers.Sqloledb
            GetConnectionString = "Provider=Sqloledb;Data Source='DATASOURCE';User ID='USER';Password='PASSWORD';ReadOnly='READONLY'"
        Case Providers.Msdasql
            'GetConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='DATASOURCE';User ID='USER';Password='PASSWORD';ReadOnly='READONLY'"
            GetConnectionString = "Provider=MSDASQL.1;Data Source=MS Access Database;DBQ='DATASOURCE';User ID='USER';Password='PASSWORD';ReadOnly='READONLY'"
        Case Providers.MicrosoftExcelDriver
            GetConnectionString = "Driver={Microsoft Excel Driver (*.xls)};DBQ='DATASOURCE';User ID='USER';Password='PASSWORD';ReadOnly='READONLY'"
        Case Providers.MicrosoftTextDriver
            GetConnectionString = "Driver={Microsoft Text Driver (*.txt; *.csv)};DBQ='DATASOURCE';User ID='USER';Password='PASSWORD';ReadOnly='READONLY'"
        Case Else
            GetConnectionString = "Provider=Sqloledb"
    End Select
End Function

Private Sub AddOleObject(ByRef ws As Worksheet, ByVal Path As String, ByVal OleName As String)
    With ws.OLEObjects.Add(FileName:=Path, Link:=False, DisplayAsIcon:=False, IconLabel:="")
        .Name = OleName
    End With
End Sub

Public Function DB≫GetInternalDatabase() As Object
    
    Const INTERNAL_DATABASE_NAME As String = "IDB"
    
    Dim O As OLEObject
    Set O = [Sheet1].OLEObjects(INTERNAL_DATABASE_NAME)
    
    O.Copy
    Application.CutCopyMode = False
    
    Set DB≫GetInternalDatabase = DB≫GetSession(Msdasql, BuildPath(IO≫File≫GetTempPath, INTERNAL_DATABASE_NAME), VBA.vbNullString, VBA.vbNullString)
    
End Function

Public Function SYS≫Process≫ThreadStart(ByVal FunctionName As String, ParamArray Args()) As Application
    Const EXECUTE_FUNCTION_NAME As String = "''ExecuteAndClose ""'FUNCTION'('PARAMS')""''"
    
    Dim p As String: p = vbNullString
    If Not Not Args Then _
        p = """""" & Join(Args, """"",""""") & """"""
    
    Dim params As New Collection
    params.Add FunctionName, "FUNCTION"
    params.Add p, "PARAMS"
    
    Dim app As Application
    Set app = VBA.CreateObject("Excel.Application")
    With app

        Set Thread≫ThreadStart = .Application
        .EnableEvents = False
        .Interactive = False
        .DisplayAlerts = False
        .ScreenUpdating = False
        .Visible = False
        .WindowState = xlMinimized

        With .Workbooks.Open(ThisWorkbook.FullName, ReadOnly:=True)
        End With

        Call .OnTime(VBA.Now(), SetParameters(EXECUTE_FUNCTION_NAME, params))
        DoEvents

    End With
    
End Function

Public Function SYS≫Environment≫GetDpi() As Long
    
    Static DPI As Long
    
    If DPI = 0 Then
        
        Dim w As Long
        w = GetDesktopWindow()
        
        Dim d As Long
        d = GetDC(w)
        
        DPI = GetDeviceCaps(d, 88)
        ReleaseDC w, d
       
    End If
    
    SYS≫Environment≫GetDpi = DPI
    
End Function

Public Sub ExecuteAndClose(ByVal FunctionName As String)
    Dim result As Variant
    result = Application.Evaluate(FunctionName)
    Application.Quit
End Sub

Public Function UI≫Android≫Tap(ByVal x As Long, ByVal Y As Long) As Boolean
    
    Const TAP_COMMAND = """'adbPath'"" shell input tap 'X' 'Y'"
    
    Dim params As New Collection
    params.Add ADB_ADB_PATH, "adbPath"
    params.Add x, "X"
    params.Add Y, "Y"
    
    UI≫Android≫Tap = (VBA.CreateObject("Wscript.Shell").Run(SetParameters(TAP_COMMAND, params), WSCRIPT_RUN_WINDOW_STATE, True) = 0)
    Set params = Nothing
    
End Function

Public Function UI≫Android≫Swipe(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    
    Const SWIPE_COMMAND = """'adbPath'"" shell input swipe 'X1' 'Y1' 'X2' 'Y2'"
    
    Dim params As New Collection
    params.Add ADB_ADB_PATH, "adbPath"
    params.Add x1, "X1"
    params.Add y1, "Y1"
    params.Add x2, "X2"
    params.Add y2, "Y2"
    
    UI≫Android≫Swipe = (VBA.CreateObject("Wscript.Shell").Run(SetParameters(SWIPE_COMMAND, params), WSCRIPT_RUN_WINDOW_STATE, True) = 0)
    Set params = Nothing
    
End Function

Public Function UI≫Android≫Wait(ByVal Sec) As Boolean
    
    Const WAIT_COMMAND As String = """'adbPath'"" shell sleep 'SECOND'"
    
    Dim params As New Collection
    params.Add ADB_ADB_PATH, "adbPath"
    params.Add Sec, "SECOND"
    
    UI≫Android≫Wait = (VBA.CreateObject("Wscript.Shell").Run(SetParameters(WAIT_COMMAND, params), WSCRIPT_RUN_WINDOW_STATE, True) = 0)
    Set params = Nothing
    
End Function

Public Function UI≫Android≫ScreenShot(Optional ByVal FileName As String = VBA.vbNullString)
    
    Const ANDROID_TEMP_PATH = "/storage/emulated/0/Download/"
    Const SCREENSHOT_COMMAND = "cmd /K 'adbPath1' shell screencap -p 'tmpPath1' & 'adbPath2' pull 'tmpPath2' ""'localPath'"" & 'adbPath3' shell rm -f 'tmpPath3' & Exit"
    
    Dim t As String
    t = VBA.Format(VBA.Now(), "yyyymmddhhnnss") & "_" & _
    Application.hWnd & ".png"
    
    If VBA.Len(FileName) = 0 Then
        FileName = IO≫Path≫GetTempPath & "\" & t
    End If
    
    Dim params As New VBA.Collection
    params.Add ADB_ADB_PATH, "adbPath1"
    params.Add ADB_ADB_PATH, "adbPath2"
    params.Add ADB_ADB_PATH, "adbPath3"
    params.Add ANDROID_TEMP_PATH & t, "tmpPath1"
    params.Add ANDROID_TEMP_PATH & t, "tmpPath2"
    params.Add ANDROID_TEMP_PATH & t, "tmpPath3"
    params.Add FileName, "localPath"
    
    VBA.CreateObject("Wscript.Shell").Run SetParameters(SCREENSHOT_COMMAND, params), WSCRIPT_RUN_WINDOW_STATE, True
    Set params = Nothing
    
    UI≫Android≫ScreenShot = FileName
    
End Function

Public Sub UI≫Android≫StartApplication(ByVal ClassName As String)
    
    Const START_COMMAND = """'adbPath'"" shell am start -n 'CLASS'"
    
    Dim params As New Collection
    params.Add ADB_ADB_PATH, "adbPath"
    params.Add ClassName, "CLASS"
    
    Call VBA.CreateObject("WScript.Shell").Run(SetParameters(START_COMMAND, params), WSCRIPT_RUN_WINDOW_STATE, True)
    Set params = Nothing
    
End Sub

Public Sub UI≫Android≫StopApplication(ByVal PackageName As String)
    
    Const STOP_COMMAND = """'adbPath'"" shell am force-stop 'PACKAGE'"
    
    Dim params As New Collection
    params.Add ADB_ADB_PATH, "adbPath"
    params.Add PackageName, "PACKAGE"
    
    Call VBA.CreateObject("WScript.Shell").Run(SetParameters(STOP_COMMAND, params), WSCRIPT_RUN_WINDOW_STATE, True)
    Set params = Nothing
    
End Sub

Public Function UI≫Android≫GetDeviceName() As String
    
    Const RESULT_HEADER = "List of devices attached" & vbNewLine
    Const STOP_COMMAND = """'adbPath'"" devices"
    
    Dim params As New Collection
    params.Add ADB_ADB_PATH, "adbPath"
    
    Dim result As String
    With (VBA.CreateObject("Wscript.Shell").Exec("C:\Users\M\Documents\Application\platform-tools\adb.exe devices"))
        Do While .Status = 0: DoEvents: Loop
        result = .StdOut.ReadAll
        result = VBA.Replace(result, RESULT_HEADER, vbNullString)
        result = VBA.Left(result, InStr(result, vbTab))
        result = VBA.Replace(result, vbTab, vbNullString)
    End With
    
    UI≫Android≫GetDeviceName = result
    Set params = Nothing
    
End Function

Public Function Math≫Random(Optional ByVal max As Long = 100, Optional ByVal min As Long = 0)
    Randomize
    Math≫Random = VBA.Int((max - min + 1) * Rnd + min)
End Function

Private Function GetParamArray(ByVal Args As Variant) As Variant
    If InternalLogic.Array≫GetLength(Args) = 1 Then
        If IsArray(Args(0)) Then
            GetParamArray = Args(0)
            Exit Function
        End If
    End If
    GetParamArray = Args
End Function

Public Function Init(obj As Object, Optional ByVal Args) As Object
    Dim ic As IConstructor
    Set ic = obj
    Set Init = ic.Init(Args)
End Function

Public Sub SaveImage(strFileName As String, rgbData() As Long)
    
    Dim bmpData() As Byte: bmpData = CreateBmpData(rgbData)
    
    If Len(Dir(strFileName)) Then
        Kill strFileName
    End If
    
    Dim f As Long: f = FreeFile()
    Open strFileName For Binary As f
        Put f, , bmpData
    Close
    
End Sub

Public Function CreateBmpData(ByRef rgbData() As Long) As Byte()
        
    Dim i As Long
    Dim j As Long
    
    Dim bmpFileHeader(13) As Byte
    Dim bmpInfoHeader(39) As Byte
    Dim bmpData() As Byte
    
    ' 画像サイズ
    Dim imgWidth  As Long: imgWidth = UBound(rgbData, 1) + 1
    Dim imgHeight As Long: imgHeight = UBound(rgbData, 2) + 1
    Dim bufPitch  As Long: bufPitch = (4 - (imgWidth * 3 Mod 4)) Mod 4
    
    ' ファイルサイズ
    Dim imgSize  As Long: imgSize = 3 * imgWidth * imgHeight
    Dim bfhSize  As Long: bfhSize = UBound(bmpFileHeader) + 1
    Dim bihSize  As Long: bihSize = UBound(bmpInfoHeader) + 1
    Dim fileSize As Long: fileSize = bfhSize + bihSize + imgSize + (bufPitch * imgWidth)
    
    ' Long -> Byte()
    Dim bfSize()    As Byte: bfSize = Long2Byte(fileSize)    ' サイズ[ファイル全体]
    Dim biSizeImg() As Byte: biSizeImg = Long2Byte(imgSize)  ' サイズ[画像]
    Dim biSize()    As Byte: biSize = Long2Byte(bihSize)     ' サイズ[BITMAPINFOHEADER]
    Dim biWidth()   As Byte: biWidth = Long2Byte(imgWidth)   ' 画像幅
    Dim biHeight()  As Byte: biHeight = Long2Byte(imgHeight) ' 画像高さ
    Dim biDPM()     As Byte: biDPM = Long2Byte(3780)         ' DPM
    
    ' BITMAPFILEHEADER
    bmpFileHeader(0) = &H42      ' bfType
    bmpFileHeader(1) = &H4D      '
    bmpFileHeader(2) = bfSize(0) ' bfSize
    bmpFileHeader(3) = bfSize(1) '
    bmpFileHeader(4) = bfSize(2) '
    bmpFileHeader(5) = bfSize(3) '
    bmpFileHeader(6) = 0         ' bfReserved1
    bmpFileHeader(7) = 0         '
    bmpFileHeader(8) = 0         ' bfReserved2
    bmpFileHeader(9) = 0         '
    bmpFileHeader(10) = &H36     ' bfOffBits
    bmpFileHeader(11) = 0        '
    bmpFileHeader(12) = 0        '
    bmpFileHeader(13) = 0        '
    
    ' BITMAPINFOHEADER
    bmpInfoHeader(0) = biSize(0)        ' biSize
    bmpInfoHeader(1) = biSize(1)        '
    bmpInfoHeader(2) = biSize(2)        '
    bmpInfoHeader(3) = biSize(3)        '
    bmpInfoHeader(4) = biWidth(0)       ' biWidth
    bmpInfoHeader(5) = biWidth(1)       '
    bmpInfoHeader(6) = biWidth(2)      '
    bmpInfoHeader(7) = biWidth(3)      '
    bmpInfoHeader(8) = biHeight(0)      ' biHeight
    bmpInfoHeader(9) = biHeight(1)      '
    bmpInfoHeader(10) = biHeight(2)      '
    bmpInfoHeader(11) = biHeight(3)      '
    bmpInfoHeader(12) = 1               ' biPlanes
    bmpInfoHeader(13) = 0               '
    bmpInfoHeader(14) = 24              ' biBitCount
    bmpInfoHeader(15) = 0               '
    bmpInfoHeader(16) = 0               ' biCompression
    bmpInfoHeader(17) = 0               '
    bmpInfoHeader(18) = 0               '
    bmpInfoHeader(19) = 0               '
    bmpInfoHeader(20) = biSizeImg(0)    ' biSizeImage
    bmpInfoHeader(21) = biSizeImg(1)    '
    bmpInfoHeader(22) = biSizeImg(2)    '
    bmpInfoHeader(23) = biSizeImg(3)    '
    bmpInfoHeader(24) = biDPM(0)        ' biXPixPerMeter
    bmpInfoHeader(25) = biDPM(1)        '
    bmpInfoHeader(26) = biDPM(2)        '
    bmpInfoHeader(27) = biDPM(3)        '
    bmpInfoHeader(28) = biDPM(0)        ' biYPixPerMeter
    bmpInfoHeader(29) = biDPM(1)        '
    bmpInfoHeader(30) = biDPM(2)        '
    bmpInfoHeader(31) = biDPM(3)        '
    bmpInfoHeader(32) = 0               ' biClrUsed
    bmpInfoHeader(33) = 0               '
    bmpInfoHeader(34) = 0               '
    bmpInfoHeader(35) = 0               '
    bmpInfoHeader(36) = 0               ' biClrImporant
    bmpInfoHeader(37) = 0               '
    bmpInfoHeader(38) = 0               '
    bmpInfoHeader(39) = 0               '
    
    Dim cnt As Long
    ReDim bmpData(fileSize - 1) As Byte
    
    For i = 0 To UBound(bmpFileHeader)
        bmpData(cnt) = bmpFileHeader(i)
        cnt = cnt + 1
    Next
    
    For i = 0 To UBound(bmpInfoHeader)
        bmpData(cnt) = bmpInfoHeader(i)
        cnt = cnt + 1
    Next
    
    For i = imgHeight - 1 To 0 Step -1
        For j = 0 To imgWidth - 1
            Dim bgr() As Byte: bgr = Long2Byte(rgbData(j, i))
            bmpData(cnt) = bgr(2): cnt = cnt + 1
            bmpData(cnt) = bgr(1): cnt = cnt + 1
            bmpData(cnt) = bgr(0): cnt = cnt + 1
        Next
        For j = 1 To bufPitch
            bmpData(cnt) = 0: cnt = cnt + 1
        Next
    Next
    
    CreateBmpData = bmpData
    
End Function

Public Function Long2Byte(ByVal lng As Long) As Byte()
    Dim l As LONG_TYPE: l.val = lng
    Dim b As BYTE_TYPE: LSet b = l
    Long2Byte = b.val
End Function

Public Function Gray2Binary(ByRef img() As Long, Optional ByVal Threshold As Byte = 0) As Long()
    
    Dim i As Integer
    Dim j As Integer
    Dim GrayScale() As Long: GrayScale = RGB2Gray(img)
    
    Dim result() As Long
    ReDim result(UBound(GrayScale, 1), UBound(GrayScale, 2))
    
    Dim t As Long
    If Threshold > 0 Then
        t = RGB(Threshold, Threshold, Threshold)
    Else
        t = AdaptiveThreshold(GrayScale)
    End If
    
    For i = 0 To UBound(GrayScale, 1) - 1
        For j = 0 To UBound(GrayScale, 2) - 1
            result(i, j) = IIf(GrayScale(i, j) > t, RGB(&HFF, &HFF, &HFF), RGB(0, 0, 0))
        Next
    Next
    
    Gray2Binary = result
    
End Function

Private Function AdaptiveThreshold(ByRef GrayScale() As Long) As Long
    
    Const MAXLONG As Long = (2 ^ 31) - 256
    Dim i As Integer
    Dim j As Integer
    Dim cnt As Long
    Dim sum As Long
    Dim avg As New Collection
    
    For i = 0 To UBound(GrayScale, 1) - 1
        For j = 0 To UBound(GrayScale, 2) - 1
            Dim b() As Byte: b = Long2Byte(GrayScale(i, j))
            sum = sum + b(0)
            cnt = cnt + 1
            If sum > MAXLONG Then
                avg.Add sum / cnt
                sum = 0
                cnt = 0
            End If
        Next
    Next
    
    avg.Add sum / cnt
    sum = 0
    cnt = 0
    Dim a As Variant
    
    For Each a In avg
        sum = sum + a
        cnt = cnt + 1
    Next
    
    sum = sum / cnt
    AdaptiveThreshold = RGB(sum, sum, sum)

End Function

Public Function RGB2Gray(ByRef img() As Long) As Long()
    
    Const Red   As Double = 0.2126
    Const Green As Double = 0.7152
    Const Blue  As Double = 0.0722
    
    Dim i As Long
    Dim j As Long
    Dim GrayScale() As Long
    ReDim GrayScale(UBound(img, 1), UBound(img, 2))
    
    For i = 0 To UBound(img, 1) - 1
        For j = 0 To UBound(img, 2) - 1
            Dim b() As Byte: b = Long2Byte(img(i, j))
            Dim g As Long: g = (b(0) * Red) + (b(1) * Green) + (b(2) * Blue) / 3
            GrayScale(i, j) = RGB(g, g, g)
        Next
    Next
    
    RGB2Gray = GrayScale
    
End Function

Public Function Byte2Base64(bytes() As Byte) As String
    
    Const vbEqual As String = "="
    Const BIN_FORMAT As String = "00000000"
    
    Static Base64Table()
    If Not Base64Table Then
        Base64Table = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "+", "/")
    End If
    
    Dim i As Long
    Dim j As Long
    Dim l As Long: l = UBound(bytes)
    Dim cnt As Long: cnt = Int(((l + 2) / 3) + 0.9) * 4 - 1
    ReDim result(cnt) As String
    
    For i = 2 To l Step 3

        Dim dec As String
        dec = Format(WorksheetFunction.Dec2Bin(bytes(i - 2)), BIN_FORMAT) & _
              Format(WorksheetFunction.Dec2Bin(bytes(i - 1)), BIN_FORMAT) & _
              Format(WorksheetFunction.Dec2Bin(bytes(i - 0)), BIN_FORMAT)
        
        result(j) = Base64Table(WorksheetFunction.Bin2Dec(CLng(Mid(dec, 1, 6)))): j = j + 1
        result(j) = Base64Table(WorksheetFunction.Bin2Dec(CLng(Mid(dec, 7, 6)))): j = j + 1
        result(j) = Base64Table(WorksheetFunction.Bin2Dec(CLng(Mid(dec, 13, 6)))): j = j + 1
        result(j) = Base64Table(WorksheetFunction.Bin2Dec(CLng(Mid(dec, 19, 6)))): j = j + 1
        
    Next
    
    Dim dec2 As String
    For i = i - 2 To l
        dec2 = dec2 & Format(WorksheetFunction.Dec2Bin(bytes(i)), BIN_FORMAT)
    Next
    
    dec2 = dec2 & String(6 - (Len(dec2) Mod 6), "0")
    For i = 0 To Len(dec2) / 6 - 1
        Dim idx As Long: idx = i * 6 + 1
        result(j) = Base64Table(WorksheetFunction.Bin2Dec(CLng(Mid(dec2, idx, 6)))): j = j + 1
    Next
    
    For j = j To cnt
        result(j) = vbEqual
    Next
    
    Byte2Base64 = Join(result, vbNullString)
    
End Function

Public Function Base64ToByte(ByVal Base64 As String) As Byte()
    
    Static Base64Table2
    If Not Not Base64Table2 Then
        Base64Table2 = New Collection
        With Base64Table2
            .Add "000000", "A"
            .Add "000001", "B"
            .Add "000010", "C"
            .Add "000011", "D"
            .Add "000100", "E"
            .Add "000101", "F"
            .Add "000110", "G"
            .Add "000111", "H"
            .Add "001000", "I"
            .Add "001001", "J"
            .Add "001010", "K"
            .Add "001011", "L"
            .Add "001100", "M"
            .Add "001101", "N"
            .Add "001110", "O"
            .Add "001111", "P"
            .Add "010000", "Q"
            .Add "010001", "R"
            .Add "010010", "S"
            .Add "010011", "T"
            .Add "010100", "U"
            .Add "010101", "V"
            .Add "010110", "W"
            .Add "010111", "X"
            .Add "011000", "Y"
            .Add "011001", "Z"
            .Add "011010", "a"
            .Add "011011", "b"
            .Add "011100", "c"
            .Add "011101", "d"
            .Add "011110", "e"
            .Add "011111", "f"
            .Add "100000", "g"
            .Add "100001", "h"
            .Add "100010", "i"
            .Add "100011", "j"
            .Add "100100", "k"
            .Add "100101", "l"
            .Add "100110", "m"
            .Add "100111", "n"
            .Add "101000", "o"
            .Add "101001", "p"
            .Add "101010", "q"
            .Add "101011", "r"
            .Add "101100", "s"
            .Add "101101", "t"
            .Add "101110", "u"
            .Add "101111", "v"
            .Add "110000", "w"
            .Add "110001", "x"
            .Add "110010", "y"
            .Add "110011", "z"
            .Add "110100", "0"
            .Add "110101", "1"
            .Add "110110", "2"
            .Add "110111", "3"
            .Add "111000", "4"
            .Add "111001", "5"
            .Add "111010", "6"
            .Add "111011", "7"
            .Add "111100", "8"
            .Add "111101", "9"
            .Add "111110", "+"
            .Add "111111", "/"
        End With
    End If
    
    Dim i As Long
    Dim l As Long: l = VBA.Len(Base64)
    
    For i = 4 To l Step 4
        Dim dec As String
        dec = WorksheetFunction.Dec2Bin(Base64Table2(Mid(Base64, i - 3, 1))) & _
              WorksheetFunction.Dec2Bin(Base64Table2(Mid(Base64, i - 2, 1))) & _
              WorksheetFunction.Dec2Bin(Base64Table2(Mid(Base64, i - 1, 1))) & _
              WorksheetFunction.Dec2Bin(Base64Table2(Mid(Base64, i - 0, 1)))
    Next
    
End Function

Public Function CreateBinaryTiffData(ByRef img() As Long) As Byte()
    
    Dim i As Long
    Dim j As Long
    Dim l1 As Long: l1 = UBound(img, 1)
    Dim l2 As Long: l2 = UBound(img, 2)
    Dim l As Long: l = (l1 + 1) * (l2 + 1)
    
    Dim cnt As Long
    ReDim tmp(l - 1)
    
    For i = 0 To l1
        For j = 0 To l2
            tmp(cnt) = IIf(img(i, j) = 0, 0, 1)
            cnt = cnt + 1
        Next
    Next
    
    cnt = l / 8 + l Mod 8
    ReDim result(cnt) As Byte
    j = 0
    
    For i = 7 To l Step 8
        result(j) = tmp(i - 7) * 1 + _
                    tmp(i - 6) * 2 + _
                    tmp(i - 5) * 4 + _
                    tmp(i - 4) * 8 + _
                    tmp(i - 3) * 16 + _
                    tmp(i - 2) * 32 + _
                    tmp(i - 1) * 64 + _
                    tmp(i - 0) * 128
        j = j + 1
    Next
    
    CreateBinaryTiffData = result
    
End Function

Public Sub DrawRectangle(ByVal x As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long)
    
    Dim dc As Long
    dc = GetDC(Application.hWnd)
    
    Dim hbr As Long
    hbr = CreateSolidBrush(Color)
    
    Dim hbrPrev As Long
    hbrPrev = SelectObject(dc, hbr)
    
    Rectangle dc, x, Y, x + Width - 1, Y + Height - 1
    
    DeleteObject hbr
    SelectObject dc, hbrPrev
    ReleaseDC Application.hWnd, dc
    
End Sub

Public Function Eor(ByVal rng As Range) As Long
    Eor = rng.Parent.Cells(Rows.Count, rng.Column).End(xlUp).Row
End Function

Public Function Eoc(ByVal rng As Range) As Long
    Eoc = rng.Parent.Cells(rng.Row, Columns.Count).End(xlToLeft).Column
End Function

Sub testa()
    Dim i As Long
    Wait 1000
    i = GetBitmapHandleFromWindow(GetForegroundWindow())
    Dim pic As Variant
    Set pic = CreateBitmap(i)
    SavePicture pic, "C:\Users\M\Desktop\test.bmp"
    DeleteObject pic
End Sub


