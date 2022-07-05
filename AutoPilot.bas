Attribute VB_Name = "AutoPilot"
Option Explicit

Private jumpRowNum As Long
Private errorTraps As Object
Private vals As Object
Private CurrentCell As Range

Private Sub Auto_Open()
    Application.OnKey "{F1}", "SetCommand_LeftClick"
    Application.OnKey "{F2}", "SetCommand_KeyPress"
    Application.OnKey "{F3}", "SetCommand_SendText"
    Application.OnKey "{F4}", "SetCommand_CaptureWindow"
    Application.OnKey "{F12}", "Exec"
End Sub

Public Sub Exec(Optional Address)

Dim i As Long
Dim r As Range

' 引数が省略された場合はデフォルト
If VBA.IsError(Address) Then
    If VBA.TypeName(Selection) <> "Range" Then Exit Sub
    Set errorTraps = CreateObject("Scripting.Dictionary")
    Set vals = CreateObject("Scripting.Dictionary")
    Set r = Selection
Else
    Set r = Range(Address)
End If

' 処理開始
For i = 1 To r.Rows.Count
    On Error GoTo Exception
    
    If Base.IsKeyPressed(vbKeyEscape) Then
        Base.UI≫Window≫ActivateWindow Application.hWnd
        If MsgBox("処理を中断しますか？", vbYesNo + vbInformation) = vbYes Then Exit Sub
    End If
    
    Set CurrentCell = r.Cells(i, 1)
    
    Dim c As String
    c = VBA.Trim(CurrentCell.Text)
    
    ' 空白行スキップ
    If VBA.Len(c) = 0 Then GoTo Continue
    
    ' コメント行スキップ
    If VBA.Left(c, 2) = "//" Then GoTo Continue
    
    ' If --> Switch
    ReplaceReservedWord2FunctionName c, "if", "Switch"
    
    ' Jump --> JumpEx
    ReplaceReservedWord2FunctionName c, "jump", "JumpEx"
    
    ' [A1]= --> SetValue
    If VBA.Left(c, 1) = "[" Then
        Dim idx1 As Integer: idx1 = InStr(c, "]")
        Dim idx2 As Integer: idx2 = InStr(idx1, c, "=")
        If idx1 <> 0 And idx2 <> 0 Then _
            c = "SetValue " & """" & VBA.Mid(c, 2, idx1 - 2) & """," & VBA.Trim(Mid(c, idx2 + 1))
    End If
    
    
    ' 処理実行
    Debug.Print "[DEBUG] " & CurrentCell.Address & ": " & c
    Application.Run "'" & c & "'"
    
    If Err.Number = 0 Then GoTo Finally
    
Exception:
    
    Dim e As Long
    Debug.Print "[ERROR] " & Err.Number & "：" & Err.Description
    
    ' 例外処理取得
    If errorTraps.Exists(Err.Number) Then
        e = Err.Number
    ElseIf errorTraps.Exists(0) Then
        e = 0
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
        Exit Sub
    End If
    
    ' 例外処理実行
    If Err.Number <> 0 And errorTraps.item(e) <> vbNullString Then
        Debug.Print errorTraps.item(e)
        Application.Run "'" & errorTraps.item(e) & "'"
        On Error GoTo 0
    End If
    
Finally:
    
    ' ジャンプ処理
    Dim j As Long: j = JumpRowNumber
    If j > 0 Then
        If j < r.Cells(1, 1).Row Or r.Cells(r.Rows.Count, 1).Row < j Then
            Exit Sub
        Else
            i = j - r.Cells(1, 1).Row
        End If
    End If
    
Continue:
Next

End Sub


Public Sub SetCommand_LeftClick()
    
    Base.UI≫Mouse≫MouseClick MouseButton.Left
    
    Dim p As New VBA.Collection
    Set p = Base.UI≫Mouse≫GetMousePoint
    
    Dim t As String
    t = Base.UI≫Window≫GetWindowTitle(Base.UI≫Window≫GetWindowFromPoint(p("X"), p("Y")))
    
    Dim hWnd As Long
    hWnd = Base.UI≫Window≫GetWindowHandle(t)
    
    Dim s As XlWindowState
    s = Base.UI≫Window≫GetWindowState(hWnd)
    
    Dim m As String
    If s = XlWindowState.xlMaximized Then
        m = ",[Max]"
    ElseIf s = XlWindowState.xlMinimized Then
        m = ",[Min]"
    Else
        
        Dim r As VBA.Collection
        Set r = Base.UI≫Window≫GetWindowRectangle(hWnd)
        
        If r("Left") > 0 Or r("Top") > 0 Then _
            m = m & "," & r("Left") & "," & r("Top")
            
        If r("Left") - r("Right") > 0 Or r("Top") - r("Bottom") > 0 Then _
            m = m & "," & r("Right") & "," & r("Bottom")
            
        If Len(m) > 0 Then _
            m = ",[Nomal]" & m
        
    End If
    
    Application.ActiveCell = "Click " & p("X") - r("Left") & "," & p("Y") - r("Top") & ",""" & EscapeStr(t) & """" & m
    Application.ActiveSheet.Cells(Application.ActiveCell.Row + 1, Application.ActiveCell.Column).Select
    
    Base.UI≫Window≫ActivateWindow Application.hWnd
    
End Sub

Public Sub Click(ByVal x As Long, ByVal Y As Long, Optional ByVal windowTitle, Optional ByVal State As WindowState = [Nomal], Optional ByVal Left As Long, Optional ByVal Top As Long, Optional ByVal Right As Long, Optional ByVal Bottom As Long)
    
    Dim hWnd As Long
    hWnd = Base.UI≫Window≫GetWindowHandle(windowTitle)
    
    If hWnd <> 0 Then
    
        ' 対象のウィンドウのステータスを設定
        Dim s1 As XlWindowState
        s1 = Base.UI≫Window≫GetWindowState(hWnd)
        
        Dim s2 As XlWindowState
        Select Case State
            Case WindowState.max
                s2 = XlWindowState.xlMaximized
            Case WindowState.min
                s2 = XlWindowState.xlMinimized
            Case Else
                s2 = XlWindowState.xlNormal
        End Select
        
        If s1 <> s2 Then
            If s1 = XlWindowState.xlMinimized Then
                Base.UI≫Window≫SetWindowState hWnd, XlWindowState.xlNormal
                Base.UI≫Window≫SetWindowState hWnd, s2
            Else
                Base.UI≫Window≫SetWindowState hWnd, s2
            End If
        End If
        
        ' 対象のウィンドウを移動
        If State = Nomal Then
            If Left > 0 And Top > 0 Then _
                Base.UI≫Window≫MoveWindow hWnd, Left, Top
            If Right - Left > 0 And Bottom - Top > 0 Then _
                Base.UI≫Window≫ResizeWindow hWnd, Right - Left, Bottom - Top
        End If
        
        ' 対象のウィンドウをアクティブ化
        Wait 200, Base.UI≫Window≫ActivateWindow(hWnd)
    
    End If
    
    ' カーソル移動
    Base.UI≫Mouse≫MouseMove Left + x, Top + Y
    
    ' クリック実行
    Base.UI≫Mouse≫MouseClick MouseButton.Left
    
End Sub

Public Sub SetCommand_KeyPress()
    
    Wait 200, MsgBox("入力するキーを押しながら、Enterキーを押してください。" & vbNewLine & "※Enterキーを押したい場合は、Enterキーを長押ししてください。", vbInformation)
    
    ' キーコード取得
    Dim codes As Variant
    Do
        codes = GetPressedKey()
        If Not IsEmpty(codes(0)) Then Exit Do
        Wait 200
    Loop
    
    Dim p As New Collection
    Set p = Base.UI≫Mouse≫GetMousePoint
    
    Dim h As Long
    h = UI≫Window≫GetWindowFromPoint(p("X"), p("Y"))
    
    Dim t As String
    t = Base.UI≫Window≫GetWindowTitle(h)
    
    Dim i As Integer
    Dim codestr As String
    Dim sorted As Variant
    ReDim sorted(UBound(codes))
    
    For i = 0 To UBound(codes)
        sorted(i) = WorksheetFunction.Small(codes, i + 1)
        codestr = codestr & "," & "vbKey" & GetKeyCodes()(sorted(i))
    Next
    
    ' 対象のウィンドウをアクティブ化
    Wait 200, Base.UI≫Window≫ActivateWindow(h)
    Base.KeyPress sorted
    
    Application.ActiveCell = "Key """ & EscapeStr(t) & """" & codestr
    Application.ActiveSheet.Cells(Application.ActiveCell.Row + 1, Application.ActiveCell.Column).Select
    
    Base.UI≫Window≫ActivateWindow Application.hWnd
    
End Sub

Public Sub Key(ByVal args1, ParamArray Vk())
    
    If Not GetKeyCodes().Exists(args1) Then
        
        Dim hWnd As Long
        hWnd = Base.UI≫Window≫GetWindowHandle(args1)
        
        ' 対象のウィンドウをアクティブ化
        Wait 200, Base.UI≫Window≫ActivateWindow(hWnd)
        
        Base.KeyPress Vk
        
    Else
        
        Dim i As Integer
        Dim ar As Variant
        ReDim ar(UBound(Vk) + 1)
        
        For i = UBound(ar) To 1 Step -1
            ar(i) = Vk(i - 1)
        Next
        
        ar(0) = args1
        Base.KeyPress ar
        
    End If
    
End Sub

Public Sub SetCommand_SendText()
    
    Dim str As String
    str = Application.InputBox("文字列を入力してください。")
    
    If VBA.Len(str) <> 0 Then
        
        Dim p As New Collection
        Set p = Base.UI≫Mouse≫GetMousePoint
        
        Dim h As Long
        h = UI≫Window≫GetWindowFromPoint(p("X"), p("Y"))
        
        Dim t As String
        t = Base.UI≫Window≫GetWindowTitle(h)
        
        SendText str, h
        
        Application.ActiveCell = "SendText """ & str & """,""" & EscapeStr(t) & """"
        Application.ActiveSheet.Cells(Application.ActiveCell.Row + 1, Application.ActiveCell.Column).Select
        
    End If
    
    Base.UI≫Window≫ActivateWindow Application.hWnd
    
End Sub

Public Sub SendText(ByVal Text As String, Optional ByVal windowTitle)
    
    Dim hWnd As Long
    hWnd = Base.UI≫Window≫GetWindowHandle(windowTitle)
    
    ' 文字列をクリップボードに格納
    Base.IO≫Clipboard≫SetClipboard Text
    
    ' 対象のウィンドウをアクティブ化
    Wait 200, Base.UI≫Window≫ActivateWindow(hWnd)
    
    ' 文字列貼り付け
    Base.UI≫Keyboard≫KeyPress vbKeyControl, vbKeyV
    
End Sub

Public Sub SetCommand_CaptureWindow()
    
    Dim p As VBA.Collection
    Set p = Base.UI≫Mouse≫GetMousePoint
    
    Dim h As Long
    h = UI≫Window≫GetWindowFromPoint(p("X"), p("Y"))
    
    Dim t As String
    t = Base.UI≫Window≫GetWindowTitle(h)
    
    Dim Path As String
    Path = Base.UI≫Window≫Dialog≫SaveFileDialog
    
    If Path <> "False" Then
        Capture Path, h
        Application.ActiveCell = "Capture """ & Path & """,""" & EscapeStr(t) & """"
        Application.ActiveSheet.Cells(Application.ActiveCell.Row + 1, Application.ActiveCell.Column).Select
    End If
    
    Base.UI≫Window≫ActivateWindow Application.hWnd
    
End Sub

Public Sub Capture(ByVal Path As String, Optional ByVal windowTitle)
    
    Dim hWnd As Long
    hWnd = Base.UI≫Window≫GetWindowHandle(windowTitle)
    
    If hWnd <> 0 Then
    
        ' 対象のウィンドウをアクティブ化
        Wait 500, Base.UI≫Window≫ActivateWindow(hWnd)
        
    Else
        ' PrintScreenキー押下
        Base.UI≫Keyboard≫KeyPress vbKeySnapshot
        Base.UI≫Keyboard≫KeyPress vbKeySnapshot
    End If
    
    ' 画像保存
    Dim p As StdPicture
    Set p = Base.IO≫Image≫GetBitmapFromWindow(hWnd)
    
    SavePicture p, Path
    Set p = Nothing
    
End Sub

Private Sub ReplaceReservedWord2FunctionName(ByRef code As String, ByVal ReservedWord As String, ByVal FunctionName As String)
    Dim l As Long: l = VBA.Len(ReservedWord)
    Dim c As String: c = VBA.LCase(VBA.Left(code, l + 1))
    If VBA.RTrim(c) = ReservedWord Then
        ' "if", "if ", "if ("
        code = FunctionName & Mid(code, l + 1)
    ElseIf c = ReservedWord & "(" Then
        ' "if("
        code = FunctionName & Mid(code, l + 1)
    End If
End Sub

Public Function Jump(Optional ByVal RowNumber As Long = 0) As String
    Jump = "JumpEx " & RowNumber
End Function
Public Sub JumpEx(Optional ByVal RowNumber As Long = 0)
    If RowNumber < 1 Then
        jumpRowNum = Rows.Count + 1
    Else
        jumpRowNum = RowNumber
    End If
End Sub
Private Function JumpRowNumber() As Long
    JumpRowNumber = jumpRowNum
    jumpRowNum = 0
End Function

Public Sub SetErrtrap(ByVal ErrorNumber As Long, ByVal Command As String)
    If errorTraps.Exists(ErrorNumber) Then
       errorTraps.item(ErrorNumber) = Command
    Else
        errorTraps.Add ErrorNumber, Command
    End If
End Sub

Public Sub RemoveErrtrap(ByVal ErrorNumber As Long)
    If errorTraps.Exists(ErrorNumber) Then errorTraps.Remove ErrorNumber
End Sub

Public Sub Error(ByVal Number As Long, Optional ByVal Description As String = vbNullString)
    On Error Resume Next
    Err.Raise Number, vbNullString, Description
End Sub

Public Sub Wait(ByVal Milliseconds As Long, Optional result = Null)
    If Not IsNull(result) Then _
        Base.Util≫Wait Milliseconds
End Sub

Public Sub Switch(ByVal flg As Boolean, ByVal TrueCommand As String, Optional ByVal FalseCommand As String = vbNullString)
    Dim c As String: c = VBA.Trim(VBA.IIf(flg, TrueCommand, FalseCommand))
    If Len(c) > 0 Then Application.Run "'" & c & "'"
End Sub

Public Sub Retry(ByVal Expr As Boolean, ByVal Jump2 As Long, Optional ByVal RetryCount As Long = &HFFF, Optional ByVal Interval As Long = 1000)
    
    Static RetryId As Object
    If RetryId Is Nothing Then Set RetryId = VBA.CreateObject("Scripting.Dictionary")
    If Not RetryId.Exists(CurrentCell.Address) Then RetryId.Add CurrentCell.Address, 0
    
    If Expr Then
        If RetryId.item(CurrentCell.Address) < RetryCount Then
            RetryId.item(CurrentCell.Address) = RetryId.item(CurrentCell.Address) + 1
            JumpEx Jump2
            If Interval > 0 Then Wait Interval
        Else
            RetryId.Remove CurrentCell.Address
            Call Error(408, "RetryCount > " & RetryCount)
        End If
    Else
        RetryId.Remove CurrentCell.Address
    End If
    
End Sub

Public Sub MessageBox(ByVal Message As String)
    MsgBox Message
End Sub

Public Sub DebugPrint(ByVal Message As String)
    Debug.Print Message
End Sub

