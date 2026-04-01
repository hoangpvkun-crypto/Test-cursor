Option Explicit

' ==========================================
' CAU HINH HE THONG
' ==========================================
Private Const R_START As Long = 14
Private Const HEADER_ROWS As Long = 13
Private Const KEY_COL_LETTER As String = "C"
Private Const FOLDER_COL_LETTER As String = "D"
Private Const LAST_COL As Long = 30 ' AD
Private Const CHUNK_ROWS As Long = 500

' ==========================================
' CHUONG TRINH CHINH
' ==========================================
Public Sub DauNoi()
    Dim wsTemplateDN As Worksheet, wsTemplateLB As Worksheet
    Dim wsEPLAN_EXPORT As Worksheet
    Dim OrgWorkbook As Workbook, ws As Worksheet
    Dim importWB As Workbook
    Dim NewWorkbook As Workbook
    Dim wsNewDN As Worksheet, wsNewLB As Worksheet
    
    Dim folder0Path As String, Folder0Name As String
    Dim folderPath As String, FolderName As String
    Dim SheetName As String, FileName As String
    Dim strPos As String, strType As String
    
    Dim DoNoCount As Long, FileCount As Long
    Dim colKey As Long, colFolder As Long
    Dim lastRowSrc As Long
    Dim srcTable As Range
    Dim fieldIdx As Long
    Dim i As Long, k As Long
    Dim fd As FileDialog
    
    Dim prevScreenUpdating As Boolean, prevDisplayAlerts As Boolean
    Dim prevCalc As XlCalculation
    Dim prevEnableEvents As Boolean
    Dim prevStatusBar As Variant
    Dim stepName As String
    
    Dim arrHeader As Variant
    Dim arrHeaderInfo As Variant
    Dim arrHeaderG As Variant
    Dim keyValues As Variant
    Dim keyList() As String
    Dim dictKeys As Object
    Dim keyValue As String
    Dim keyItem As Variant
    Dim keyIndex As Long

    stepName = "Khoi tao"
    On Error GoTo ErrHandler
    
    ' 1. Luu trang thai Excel
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    prevCalc = Application.Calculation
    prevEnableEvents = Application.EnableEvents
    prevStatusBar = Application.StatusBar
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = False
    Application.CutCopyMode = False
    
    ' 2. Lay template
    Set wsTemplateDN = Nothing
    Set wsTemplateLB = Nothing
    Set wsEPLAN_EXPORT = Nothing
    
    On Error Resume Next
    Set wsTemplateDN = ThisWorkbook.Sheets("DN_Template")
    Set wsTemplateLB = ThisWorkbook.Sheets("LB_Template")
    Set wsEPLAN_EXPORT = ThisWorkbook.Sheets("EPLAN_DN")
    On Error GoTo ErrHandler
    
    If wsTemplateDN Is Nothing Or wsTemplateLB Is Nothing Then
        MsgBox "Thieu sheet DN_Template hoac LB_Template.", vbCritical
        GoTo SafeExit
    End If
    
    ' 3. Chon file nguon EPLAN
    stepName = "Chon file EPLAN"
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Mo tep Excel xuat tu EPLAN"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb", 1
        If .Show <> -1 Then GoTo SafeExit
        Set importWB = Workbooks.Open(.SelectedItems(1), ReadOnly:=True)
    End With
    Set OrgWorkbook = importWB
    
    ' 4. Lay sheet Smart Wiring_HA
    Set ws = Nothing
    On Error Resume Next
    Set ws = OrgWorkbook.Sheets("Smart Wiring_HA")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        MsgBox "File nhap khong co sheet 'Smart Wiring_HA'.", vbCritical
        GoTo SafeExit
    End If
    
    ' 5. Tao folder goc output
    stepName = "Tao folder output"
    Folder0Name = Left(OrgWorkbook.Name, InStrRev(OrgWorkbook.Name, ".") - 1)
    Folder0Name = ReplaceInvalidCharacters(Folder0Name)
    
    If Len(OrgWorkbook.Path) = 0 Then
        MsgBox "File EPLAN chua duoc luu.", vbCritical
        GoTo SafeExit
    End If
    
    folder0Path = OrgWorkbook.Path & "\" & Folder0Name
    If Dir(folder0Path, vbDirectory) = "" Then MkDir folder0Path
    
    ' 6. Chuan bi du lieu
    colKey = ws.Columns(KEY_COL_LETTER).Column
    colFolder = ws.Columns(FOLDER_COL_LETTER).Column
    
    stepName = "Lay last row"
    lastRowSrc = GetLastRowByValues(ws, colKey, HEADER_ROWS + 1)
    If lastRowSrc < HEADER_ROWS + 1 Then
        MsgBox "Khong co du lieu.", vbCritical
        GoTo SafeExit
    End If
    
    Set srcTable = ws.Range(ws.Cells(HEADER_ROWS, 1), ws.Cells(lastRowSrc, LAST_COL))
    fieldIdx = colKey - srcTable.Column + 1
    
    ' Lay danh sach key duy nhat tu mang (khong tao cot tam)
    keyValues = ws.Range(ws.Cells(HEADER_ROWS + 1, colKey), ws.Cells(lastRowSrc, colKey)).Value2
    Set dictKeys = CreateObject("Scripting.Dictionary")
    dictKeys.CompareMode = vbTextCompare
    
    For i = 1 To UBound(keyValues, 1)
        keyValue = CStr(keyValues(i, 1))
        If Len(keyValue) > 0 Then
            If Not dictKeys.Exists(keyValue) Then dictKeys.Add keyValue, 1
        End If
    Next i
    
    DoNoCount = dictKeys.Count
    If DoNoCount <= 0 Then GoTo SafeExit
    
    ReDim keyList(1 To DoNoCount)
    keyIndex = 0
    For Each keyItem In dictKeys.Keys
        keyIndex = keyIndex + 1
        keyList(keyIndex) = CStr(keyItem)
    Next keyItem
    
    ' Cache header de giam doc sheet
    arrHeader = ws.Range(ws.Cells(HEADER_ROWS, 1), ws.Cells(HEADER_ROWS, LAST_COL)).Value2
    arrHeaderInfo = ws.Range("A2:AD4").Value2
    arrHeaderG = ws.Range("E6:E10").Value2
    
    ' ==========================================
    ' 7. VONG LAP XU LY
    ' ==========================================
    FileCount = 0
    
    For i = 1 To DoNoCount
        DoEvents
        
        keyValue = keyList(i)
        
        ' Tao workbook moi
        Set NewWorkbook = Workbooks.Add(xlWBATWorksheet)
        wsTemplateLB.Copy Before:=NewWorkbook.Sheets(1)
        wsTemplateDN.Copy Before:=NewWorkbook.Sheets(1)
        DeleteNonTemplateSheets NewWorkbook, "DN_Template", "LB_Template"
        
        Set wsNewDN = NewWorkbook.Sheets("DN_Template")
        Set wsNewLB = NewWorkbook.Sheets("LB_Template")
        
        ' Format Text
        wsNewDN.Columns("A:AD").NumberFormat = "@"
        
        ' Filter & Copy
        ws.AutoFilterMode = False
        srcTable.AutoFilter Field:=fieldIdx, Criteria1:=keyValue
        
        ' Copy Header Chinh (Dong 13)
        wsNewDN.Range("A" & HEADER_ROWS & ":AD" & HEADER_ROWS).Value2 = arrHeader
        
        ' Copy Data
        Dim copiedRows As Long
        copiedRows = CopyVisibleRowsByScan(ws, HEADER_ROWS + 1, lastRowSrc, LAST_COL, wsNewDN.Range("A" & (HEADER_ROWS + 1)))
        
        If copiedRows = 0 Then
            NewWorkbook.Close SaveChanges:=False
            GoTo NextLoop
        End If
        
        Dim lastRowDN As Long
        lastRowDN = HEADER_ROWS + copiedRows
        
        ' Copy Header Info (A2:AD4)
        Dim rSrc As Long
        Dim arrRow As Variant
        For rSrc = 2 To 4
            arrRow = Application.Index(arrHeaderInfo, rSrc - 1, 0)
            With wsNewDN.Range("A" & rSrc & ":AD" & rSrc)
                .UnMerge
                .Value2 = arrRow
                .Merge
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        Next rSrc
        
        ' Copy tam E6:E10 -> G6:G10 (G10 se duoc cap nhat sau)
        wsNewDN.Range("G6:G10").Value2 = arrHeaderG
        
        ' Sort E & K
        With wsNewDN.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsNewDN.Range("E" & HEADER_ROWS & ":E" & lastRowDN), SortOn:=xlSortOnValues, Order:=xlAscending
            .SortFields.Add Key:=wsNewDN.Range("K" & HEADER_ROWS & ":K" & lastRowDN), SortOn:=xlSortOnValues, Order:=xlAscending
            .SetRange wsNewDN.Range("A" & HEADER_ROWS & ":AD" & lastRowDN)
            .Header = xlYes
            .Apply
        End With
        
        ' Danh STT
        Dim arrSTT() As Variant
        ReDim arrSTT(1 To copiedRows, 1 To 1)
        For k = 1 To copiedRows
            arrSTT(k, 1) = k
        Next k
        wsNewDN.Range("A" & (HEADER_ROWS + 1)).Resize(copiedRows, 1).Value2 = arrSTT
        
        ' ==========================================
        ' FORMATTING & DIEM CHUM
        ' ==========================================
        TaoDiemChum_Run_OnSheet wsNewDN
        
        ' Xu ly Label & Check Size Day
        Dim maxWireSize As Double
        maxWireSize = 0
        
        Dim arrSrc As Variant
        arrSrc = wsNewDN.Range("A" & R_START & ":M" & lastRowDN).Value2
        
        Dim arrLbl As Variant
        ReDim arrLbl(1 To copiedRows * 2, 1 To 4)
        
        Dim rOut As Long: rOut = 0
        Dim rowIdx As Long
        Dim sE As String, sh As String, valM As Variant
        Dim tmpSize As Double
        
        For rowIdx = 1 To UBound(arrSrc, 1)
            valM = arrSrc(rowIdx, 13)
            tmpSize = Val(CStr(valM))
            If tmpSize > maxWireSize Then maxWireSize = tmpSize
            
            If IsError(arrSrc(rowIdx, 5)) Then sE = "" Else sE = Trim(CStr(arrSrc(rowIdx, 5)))
            If IsError(arrSrc(rowIdx, 8)) Then sh = "" Else sh = Trim(CStr(arrSrc(rowIdx, 8)))
            
            Dim sz As Double: sz = SizeOngLong(valM)
            
            rOut = rOut + 1
            arrLbl(rOut, 1) = rOut: arrLbl(rOut, 2) = "'" & sE: arrLbl(rOut, 3) = arrSrc(rowIdx, 12): arrLbl(rOut, 4) = sz
            rOut = rOut + 1
            arrLbl(rOut, 1) = rOut: arrLbl(rOut, 2) = "'" & sh: arrLbl(rOut, 3) = arrSrc(rowIdx, 12): arrLbl(rOut, 4) = sz
        Next rowIdx
        wsNewLB.Range("A" & (HEADER_ROWS + 1)).Resize(UBound(arrLbl, 1), 4).Value2 = arrLbl
        
        ' ==========================================
        ' [FIX] LAY THONG TIN TU COT AB (Vi tri) VA AD (Suffix loai)
        ' ==========================================
        Dim rngFoundKey As Range
        Dim strPosFromSource As String
        Dim strSuffixFromSource As String
        strPosFromSource = ""
        strSuffixFromSource = ""
        
        ' Tim key trong file nguon
        Set rngFoundKey = ws.Columns(colKey).Find(What:=keyValue, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not rngFoundKey Is Nothing Then
            ' Cot AB (28) -> Vi tri
            strPosFromSource = CStr(ws.Cells(rngFoundKey.Row, 28).Value)
            ' Cot AD (30) -> Suffix (Day dien / Cap dien)
            strSuffixFromSource = CStr(ws.Cells(rngFoundKey.Row, 30).Value)
        End If
        
        ' 1. Cap nhat G10 (Vi tri)
        If Len(Trim(strPosFromSource)) > 0 And strPosFromSource <> "0" Then
            wsNewDN.Range("G10").Value2 = strPosFromSource
        End If
        
        ' 2. Cap nhat G11 (Loai = Prefix + Suffix tu AD)
        Dim strPrefix As String
        If maxWireSize <= 1 Then
            strPrefix = ChrW(272) & "i" & ChrW(7873) & "u khi" & ChrW(7875) & "n" ' Dieu khien
        Else
            strPrefix = ChrW(272) & ChrW(7897) & "ng l" & ChrW(7921) & "c" ' Dong luc
        End If
        
        ' Neu AD rong -> Mac dinh la "Day dien"
        If Len(Trim(strSuffixFromSource)) = 0 Then
            strSuffixFromSource = "D" & ChrW(226) & "y " & ChrW(273) & "i" & ChrW(7879) & "n"
        End If
        
        wsNewDN.Range("G11").Value2 = strPrefix & " (" & strSuffixFromSource & ")"
        
        ' Cap nhat Header Label
        wsNewLB.Cells(10, "B").Value2 = wsNewDN.Cells(R_START, "D").Value2
        wsNewLB.Cells(11, "B").Value2 = "LB" & Mid$(CStr(wsNewDN.Cells(R_START, "C").Value2), 3)
        wsNewLB.Range("C6:C11").Value2 = wsNewDN.Range("G6:G11").Value2
        
        ' ==========================================
        ' LOGIC DAT TEN FILE & FOLDER
        ' ==========================================
        stepName = "Dat ten va Luu"
        
        SheetName = CStr(wsNewDN.Cells(HEADER_ROWS + 1, colKey).Value2)
        FolderName = CStr(wsNewDN.Cells(HEADER_ROWS + 1, colFolder).Value2)
        strPos = CStr(wsNewDN.Cells(10, 7).Value2) ' G10
        strType = CStr(wsNewDN.Cells(11, 7).Value2) ' G11
        
        ' Lam sach ky tu
        SheetName = ReplaceInvalidCharacters(SheetName)
        FolderName = ReplaceInvalidCharacters(FolderName)
        strPos = ReplaceInvalidCharacters(strPos)
        strType = ReplaceInvalidCharacters(strType)
        
        ' Ten file: Ma_ViTri_Loai (Giu nguyen dau cach)
        FileName = SheetName & "_" & strPos & "_" & strType
        
        If Len(FolderName) = 0 Or Len(SheetName) = 0 Or Len(strPos) = 0 Then
            folderPath = folder0Path & "\ERR\"
            FileName = FileName & "_" & i
        Else
            folderPath = folder0Path & "\" & FolderName & "\"
        End If
        
        If Dir(folderPath, vbDirectory) = "" Then
            On Error Resume Next
            MkDir folderPath
            On Error GoTo ErrHandler
        End If
        
        HideColumns wsNewDN, True
        
        On Error Resume Next
        wsNewDN.Name = Left(SheetName, 31)
        wsNewLB.Name = Left(SheetName & "_LB", 31)
        On Error GoTo ErrHandler
        
        ' Luu file
        NewWorkbook.SaveAs FileName:=folderPath & FileName & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        NewWorkbook.Close SaveChanges:=False
        
        FileCount = FileCount + 1
        
        ' Cap nhat tien trinh
        Dim pct As String
        pct = Format(i / DoNoCount, "0.00%")
        If Not wsEPLAN_EXPORT Is Nothing Then
            Application.ScreenUpdating = True
            wsEPLAN_EXPORT.Range("E2").Value = "Da xu ly " & pct & "..."
            Application.ScreenUpdating = False
        End If
        Application.StatusBar = "Da xu ly " & pct & "..."
        
NextLoop:
        Set wsNewDN = Nothing
        Set wsNewLB = Nothing
        Set NewWorkbook = Nothing
        ws.AutoFilterMode = False
    Next i
    
    ' Don dep
    ws.AutoFilterMode = False
    
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    
    If Not OrgWorkbook Is Nothing Then OrgWorkbook.Close SaveChanges:=False
    
    MsgBox "Da tao " & FileCount & " tep excel tai " & folder0Path, vbInformation
    If Not wsEPLAN_EXPORT Is Nothing Then
        wsEPLAN_EXPORT.Range("E2").Value = "Hoan thanh!"
    End If
    
SafeExit:
    On Error Resume Next
    Application.CutCopyMode = False
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEnableEvents
    Application.StatusBar = prevStatusBar
    If Not OrgWorkbook Is Nothing Then OrgWorkbook.Close SaveChanges:=False
    Exit Sub
    
ErrHandler:
    MsgBox "Loi: " & Err.Number & " - " & Err.Description & vbCrLf & "Buoc: " & stepName, vbCritical
    Resume SafeExit
End Sub

' ==========================================
' CAC HAM PHU TRO (QUAN TRONG)
' ==========================================
Private Function ReplaceInvalidCharacters(ByVal txt As String) As String
    Dim invalidChars As String, k As Long
    ' Them [] va ky tu xuong dong vao danh sach cam
    invalidChars = "\/:*?""<>|[]" & Chr(10) & Chr(13)
    For k = 1 To Len(invalidChars)
        txt = Replace(txt, Mid$(invalidChars, k, 1), "")
    Next k
    txt = Application.WorksheetFunction.Clean(txt)
    txt = Replace(txt, vbTab, " ")
    txt = Trim$(txt)
    ReplaceInvalidCharacters = txt
End Function

Private Function GetUniqueFilePath(ByVal folderPath As String, ByVal baseFileName As String) As String
    Dim p As Long, nameOnly As String, ext As String, fullPath As String, n As Long
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    p = InStrRev(baseFileName, ".")
    If p > 0 Then nameOnly = Left$(baseFileName, p - 1): ext = Mid$(baseFileName, p) Else nameOnly = baseFileName: ext = ".xlsx"
    fullPath = folderPath & nameOnly & ext
    n = 1
    Do While Len(Dir(fullPath)) > 0
        fullPath = folderPath & nameOnly & "_" & Format$(n, "000") & ext
        n = n + 1
    Loop
    GetUniqueFilePath = fullPath
End Function

Private Function CopyVisibleRowsByScan(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal c2 As Long, ByVal tgt As Range) As Long
    Dim out As Long, r As Long, rStart As Long, rEnd As Long, cur As Long, tk As Long, lc As Long
    out = 0: r = r1
    Do While r <= r2
        If Not ws.Rows(r).Hidden Then
            rStart = r: Do While r <= r2 And Not ws.Rows(r).Hidden: r = r + 1: Loop: rEnd = r - 1
            cur = rStart
            Do While cur <= rEnd
                tk = CHUNK_ROWS: If cur + tk - 1 > rEnd Then tk = rEnd - cur + 1
                On Error Resume Next
                tgt.Offset(out, 0).Resize(tk, c2).Value2 = ws.Range(ws.Cells(cur, 1), ws.Cells(cur + tk - 1, c2)).Value2
                If Err.Number = 7 Then Err.Clear: For lc = 0 To tk - 1: tgt.Offset(out + lc, 0).Resize(1, c2).Value2 = ws.Range(ws.Cells(cur + lc, 1), ws.Cells(cur + lc, c2)).Value2: Next
                On Error GoTo 0
                out = out + tk: cur = cur + tk
            Loop
        Else: r = r + 1: End If
    Loop
    CopyVisibleRowsByScan = out
End Function

Private Function SizeOngLong(ByVal valInput As Variant) As Double
    If IsError(valInput) Or Not IsNumeric(valInput) Then
        SizeOngLong = 0
        Exit Function
    End If
    Dim cd As Double
    cd = CDbl(valInput)
    If cd >= 16 Then
        SizeOngLong = 7
    ElseIf cd >= 10 Then
        SizeOngLong = 6
    ElseIf cd >= 6 Then
        SizeOngLong = 5.2
    ElseIf cd >= 4 Then
        SizeOngLong = 4.6
    ElseIf cd >= 2.5 Then
        SizeOngLong = 4.2
    ElseIf cd >= 0 Then
        SizeOngLong = 3.2
    Else
        SizeOngLong = 0
    End If
End Function

Private Sub DeleteNonTemplateSheets(ByVal wb As Workbook, ByVal k1 As String, ByVal k2 As String)
    Dim sh As Worksheet
    Application.DisplayAlerts = False
    For Each sh In wb.Worksheets
        If sh.Name <> k1 And sh.Name <> k2 Then sh.Delete
    Next sh
    Application.DisplayAlerts = True
End Sub

Private Function GetLastRowByValues(ByVal ws As Worksheet, ByVal col As Long, ByVal rMin As Long) As Long
    Dim f As Range: On Error Resume Next
    Set f = ws.Columns(col).Find("*", , xlValues, xlPart, xlByRows, xlPrevious)
    If f Is Nothing Then
        GetLastRowByValues = 0
    ElseIf f.Row < rMin Then
        GetLastRowByValues = 0
    Else
        GetLastRowByValues = f.Row
    End If
End Function

Private Sub HideColumns(ByVal ws As Worksheet, ByVal H As Boolean)
    ws.Columns("B:C").Hidden = H: ws.Columns("F").Hidden = H: ws.Columns("I").Hidden = H
    ws.Columns("M:Q").Hidden = H: ws.Columns("V:Z").Hidden = H: ws.Columns("AA:AD").Hidden = H
End Sub

Private Function NormAN(ByVal v As Variant) As String
    Dim s As String, i As Long, cH As String, c As Integer
    If IsError(v) Then Exit Function
    s = UCase$(CStr(v)): If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        cH = Mid$(s, i, 1): c = Asc(cH)
        If (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or cH = "+" Or cH = "-" Then NormAN = NormAN & cH
    Next i
End Function

Private Function SortLinesNumeric(ByVal s As String) As String
    Dim arr() As String, i As Long, j As Long, t As String
    If Len(s) = 0 Then SortLinesNumeric = "": Exit Function
    arr = Split(s, vbLf)
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If Val(arr(j)) < Val(arr(i)) Then t = arr(i): arr(i) = arr(j): arr(j) = t
        Next j
    Next i
    SortLinesNumeric = Join(arr, vbLf)
End Function

Public Sub TaoDiemChum_Run_OnSheet(ByVal ws As Worksheet)
    ' Xoa dinh dang cu
    ws.Cells.FormatConditions.Delete
    
    Dim rEnd As Long
    rEnd = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If rEnd < 14 Then Exit Sub
    
    ' [PHAN MOI] KE BANG & TO MAU XEN KE
    Dim rngTable As Range
    Dim colLast As Long
    ' Xac dinh cot cuoi cung (thuong la Cot AD = 30)
    colLast = LAST_COL
    Set rngTable = ws.Range(ws.Cells(14, 1), ws.Cells(rEnd, colLast))
    
    ' Ke vien den (All Borders)
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
    
    ' To mau xen ke (Xanh nhat / Trang)
    Dim r As Long
    For r = 14 To rEnd
        If (r - 14) Mod 2 = 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, colLast)).Interior.Color = RGB(220, 230, 241) ' Xanh nhat
        Else
            ws.Range(ws.Cells(r, 1), ws.Cells(r, colLast)).Interior.Color = xlNone ' Trang
        End If
    Next r
    
    ' --- LOGIC DIEM CHUM (GIU NGUYEN) ---
    Dim A, D, E, G, H
    A = ws.Range("A14:A" & rEnd).Value2
    D = ws.Range("D14:D" & rEnd).Value2
    E = ws.Range("E14:E" & rEnd).Value2
    G = ws.Range("G14:G" & rEnd).Value2
    H = ws.Range("H14:H" & rEnd).Value2
    
    Dim mGH As Object, mGE As Object, mDE As Object, sAll As Object
    Set mGH = CreateObject("Scripting.Dictionary")
    Set mGE = CreateObject("Scripting.Dictionary")
    Set mDE = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, k As String, j As Variant
    For i = 1 To UBound(A)
        k = NormAN(G(i, 1)) & "|" & NormAN(H(i, 1))
        If Not mGH.Exists(k) Then mGH.Add k, New Collection
        mGH(k).Add i
        
        k = NormAN(G(i, 1)) & "|" & NormAN(E(i, 1))
        If Not mGE.Exists(k) Then mGE.Add k, New Collection
        mGE(k).Add i
        
        k = NormAN(D(i, 1)) & "|" & NormAN(E(i, 1))
        If Not mDE.Exists(k) Then mDE.Add k, New Collection
        mDE(k).Add i
    Next
    
    ws.Range("J14:J" & rEnd).ClearContents
    Dim s1 As String, s2 As String, gN As String, hN As String, eN As String
    
    For i = 1 To UBound(A)
        s1 = "": s2 = ""
        Set sAll = CreateObject("Scripting.Dictionary")
        gN = NormAN(G(i, 1)): hN = NormAN(H(i, 1)): eN = NormAN(E(i, 1))
        
        k = gN & "|" & hN
        If mGH.Exists(k) Then
            For Each j In mGH(k)
                If CLng(j) <> i And Not sAll.Exists(CStr(A(j, 1))) Then
                    s1 = s1 & IIf(s1 <> "", vbLf, "") & A(j, 1): sAll(CStr(A(j, 1))) = 1
                End If
            Next j
        End If
        
        k = NormAN(D(i, 1)) & "|" & eN
        If mDE.Exists(k) Then
            For Each j In mDE(k)
                If CLng(j) <> i And Not sAll.Exists(CStr(A(j, 1))) Then
                    s1 = s1 & IIf(s1 <> "", vbLf, "") & A(j, 1): sAll(CStr(A(j, 1))) = 1
                End If
            Next j
        End If
        
        k = gN & "|" & hN
        If mGE.Exists(k) Then
            For Each j In mGE(k)
                If CLng(j) <> i And Not sAll.Exists(CStr(A(j, 1))) Then
                    s2 = s2 & IIf(s2 <> "", vbLf, "") & A(j, 1): sAll(CStr(A(j, 1))) = 1
                End If
            Next j
        End If
        
        If Len(s1) > 0 Then s1 = SortLinesNumeric(s1)
        If Len(s2) > 0 Then s2 = SortLinesNumeric(s2)
        
        ws.Cells(13 + i, "J").Value2 = IIf(s1 <> "" And s2 <> "", s1 & vbLf & s2, s1 & s2)
    Next
    
    ws.Range("J14:J" & rEnd).WrapText = True
    
    Dim cE As Object, cH As Object, k2 As String
    Set cE = CreateObject("Scripting.Dictionary")
    Set cH = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(A)
        If Len(NormAN(E(i, 1))) > 0 Then
            k2 = NormAN(D(i, 1)) & "|" & NormAN(E(i, 1))
            cE(k2) = cE(k2) + 1
        End If
        If Len(NormAN(H(i, 1))) > 0 Then
            k2 = NormAN(G(i, 1)) & "|" & NormAN(H(i, 1))
            cH(k2) = cH(k2) + 1
        End If
    Next
    
    For i = 1 To UBound(A)
        If cE(NormAN(D(i, 1)) & "|" & NormAN(E(i, 1))) >= 3 Then ws.Cells(13 + i, "E").Interior.Color = vbYellow
        If cH(NormAN(G(i, 1)) & "|" & NormAN(H(i, 1))) >= 3 Then ws.Cells(13 + i, "H").Interior.Color = vbYellow
    Next
End Sub
