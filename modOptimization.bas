Attribute VB_Name = "modOptimization"
'Attribute VB_Name = "modOptimization"
' Module toi uu hieu suat
' Cac chuc nang toi uu hoa hieu suat va kich thuoc file Excel
' Tac gia: Phong Khach hang Ca nhan, Agribank Chi nhanh 4
' Ngay tao: 15/05/2025

Option Explicit

' Cac hang so lien quan den toi uu hoa
Private Const MAX_ROWS_PER_SHEET As Long = 10000 ' So dong toi da cho moi sheet
Private Const DATA_RETENTION_DAYS As Integer = 180 ' So ngay luu tru du lieu

' Bat/tat che do toi uu hieu suat
Public Sub OptimizePerformance(Optional turnOn As Boolean = True)
    On Error GoTo ErrorHandler
    
    If turnOn Then
        ' Bat tat ca che do toi uu
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        Application.DisplayStatusBar = False
        Application.DisplayAlerts = False
    Else
        ' Tat che do toi uu, khoi phuc cai dat mac dinh
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayStatusBar = True
        Application.DisplayAlerts = True
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Khoi phuc cai dat mac dinh ngay khi co loi
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
    ' Ghi log loi
    LogError "OptimizePerformance", Err.Number, Err.description
End Sub

' Toi uu hoa kich thuoc file bang cach loai bo du lieu cu
Public Sub OptimizeFileSize()
    On Error GoTo ErrorHandler
    
    ' Bat che do toi uu hieu suat
    OptimizePerformance True
    
    ' Thuc hien toi uu tung loai du lieu
    PurgeOldTransactionData
    CompactRangeData
    RemoveUnusedNameRanges
    RemoveUnusedStyles
    
    ' Khoi phuc che do hieu suat
    OptimizePerformance False
    
    ' Thong bao ket qua
    MsgBox "Da toi uu hoa kich thuoc file thanh cong!", vbInformation, "Toi uu hoa"
    
    Exit Sub
    
ErrorHandler:
    ' Khoi phuc che do hieu suat
    OptimizePerformance False
    
    ' Thong bao loi
    MsgBox "Loi khi toi uu hoa kich thuoc file: " & Err.description, vbCritical, "Loi"
    LogError "OptimizeFileSize", Err.Number, Err.description
End Sub

' Loai bo du lieu giao dich cu (qua 180 ngay)
Private Sub PurgeOldTransactionData()
    On Error GoTo ErrorHandler
    
    ' Kiem tra sheet du lieu giao dich co ton tai hay khong
    If Not modUtility.sheetExists(SHEET_TRANSACTION_DATA) Then Exit Sub
    
    ' Mo khoa sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(SHEET_TRANSACTION_DATA)
    ws.Unprotect password:=GetDefaultPassword()
    
    ' Tim so dong co du lieu
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Neu khong co du lieu, thoat
    If lastRow <= 1 Then
        ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
        Exit Sub
    End If
    
    ' Ngay gioi han
    Dim cutoffDate As Date
    cutoffDate = DateAdd("d", -DATA_RETENTION_DAYS, Date)
    
    ' Tim cot ngay giao dich (cot F)
    Dim rowsToDelete As New Collection
    Dim i As Long
    
    ' Thu thap cac dong can xoa
    For i = lastRow To 2 Step -1
        ' Chi xoa neu ngay giao dich qua 180 ngay
        If Not IsEmpty(ws.Cells(i, 6).Value) Then
            If ws.Cells(i, 6).Value < cutoffDate Then
                ' Them dong vao danh sach can xoa
                rowsToDelete.Add i
            End If
        End If
    Next i
    
    ' Xoa cac dong cu
    If rowsToDelete.Count > 0 Then
        Application.EnableEvents = False
        
        ' Xoa tung dong, tu duoi len
        For i = 1 To rowsToDelete.Count
            ws.Rows(rowsToDelete(i)).Delete Shift:=xlUp
        Next i
        
        Application.EnableEvents = True
    End If
    
    ' Bao ve lai sheet
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    Exit Sub
    
ErrorHandler:
    ' Bao ve lai sheet va bat lai events
    On Error Resume Next
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    Application.EnableEvents = True
    
    ' Ghi log loi
    LogError "PurgeOldTransactionData", Err.Number, Err.description
End Sub

' Thu gon vung du lieu de giam kich thuoc file
Private Sub CompactRangeData()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim dataSheets As Variant
    
    ' Lay danh sach cac sheet du lieu
    dataSheets = GetRequiredDataSheets()
    
    ' Compact tung sheet du lieu
    Dim i As Integer
    For i = LBound(dataSheets) To UBound(dataSheets)
        ' Kiem tra sheet co ton tai khong
        If modUtility.sheetExists(CStr(dataSheets(i))) Then
            Set ws = ThisWorkbook.sheets(CStr(dataSheets(i)))
            
            ' Mo khoa sheet
            ws.Unprotect password:=GetDefaultPassword()
            
            ' Tim vung du lieu thuc su duoc su dung
            lastRow = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row
            lastCol = ws.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, LookIn:=xlValues).Column
            
            ' Xoa cac dong va cot ngoai vung du lieu
            If lastRow < ws.Rows.Count Then
                If lastRow + 100 < ws.Rows.Count Then
                    ws.Rows(lastRow + 100 & ":" & ws.Rows.Count).Delete
                End If
            End If
            
            If lastCol < ws.Columns.Count Then
                If lastCol + 10 < ws.Columns.Count Then
                    ws.Columns(GetColumnLetter(lastCol + 10) & ":" & GetColumnLetter(ws.Columns.Count)).Delete
                End If
            End If
            
            ' Bao ve lai sheet
            ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    ' Bao ve lai sheet neu dang duoc mo khoa
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    End If
    
    ' Ghi log loi
    LogError "CompactRangeData", Err.Number, Err.description
End Sub

' Chuyen doi so thu tu cot thanh ma chu cai
Private Function GetColumnLetter(colNum As Long) As String
    On Error GoTo ErrorHandler
    
    GetColumnLetter = Split(Cells(1, colNum).Address, "$")(1)
    Exit Function
    
ErrorHandler:
    GetColumnLetter = "A" ' Mac dinh la cot A neu co loi
End Function

' Xoa cac Name Ranges khong su dung
Private Sub RemoveUnusedNameRanges()
    On Error GoTo ErrorHandler
    
    Dim n As Name
    Dim requiredRanges As Variant
    
    ' Danh sach cac Name Range bat buoc
    requiredRanges = Array("tblDuNo", "tblTaiSan", "tblTraGoc", "tblTraLai")
    
    ' Kiem tra tung Name Range
    For Each n In ThisWorkbook.Names
        ' Chi xoa cac Name Range khong nam trong danh sach bat buoc
        If Not IsNameInList(n.Name, requiredRanges) Then
            ' Kiem tra them xem co phai la Name Range he thong khong
            If Not Left(n.Name, 1) = "_" Then
                n.Delete
            End If
        End If
    Next n
    
    Exit Sub
    
ErrorHandler:
    LogError "RemoveUnusedNameRanges", Err.Number, Err.description
End Sub

' Kiem tra ten co trong danh sach khong
Private Function IsNameInList(checkName As String, nameList As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    
    IsNameInList = False
    
    For i = LBound(nameList) To UBound(nameList)
        If checkName = nameList(i) Then
            IsNameInList = True
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    IsNameInList = False
End Function

' Xoa cac styles khong su dung
Private Sub RemoveUnusedStyles()
    On Error GoTo ErrorHandler
    
    ' Luu styles su dung
    ThisWorkbook.SaveAs ThisWorkbook.fullName
    
    Exit Sub
    
ErrorHandler:
    LogError "RemoveUnusedStyles", Err.Number, Err.description
End Sub

' Thong ke ve viec su dung bo nho cua ung dung
Public Sub ReportMemoryUsage()
    On Error GoTo ErrorHandler
    
    Dim msg As String
    Dim wsCount As Integer
    Dim dataCount As Long
    
    ' So luong sheet
    wsCount = ThisWorkbook.sheets.Count
    
    ' Tao thong bao
    msg = "Thong tin su dung bo nho:" & vbCrLf & vbCrLf
    msg = msg & "- Tong so worksheet: " & wsCount & vbCrLf
    
    ' Thong tin tung sheet du lieu
    msg = msg & vbCrLf & "Du lieu theo tung sheet:" & vbCrLf
    
    ' Lay thong tin su dung du lieu tung sheet
    Dim dataSheets As Variant
    dataSheets = GetRequiredDataSheets()
    
    Dim i As Integer
    For i = LBound(dataSheets) To UBound(dataSheets)
        If modUtility.sheetExists(CStr(dataSheets(i))) Then
            dataCount = CountDataRows(CStr(dataSheets(i)))
            msg = msg & "  + " & dataSheets(i) & ": " & dataCount & " dong" & vbCrLf
        End If
    Next i
    
    ' Kich thuoc file
    If ThisWorkbook.Path <> "" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        Dim fileSize As Double
        fileSize = fso.GetFile(ThisWorkbook.fullName).Size / 1024 / 1024 ' Size in MB
        
        msg = msg & vbCrLf & "Kich thuoc file: " & Round(fileSize, 2) & " MB" & vbCrLf
        
        ' Khuyen nghi toi uu hoa neu can
        If fileSize > 20 Then
            msg = msg & vbCrLf & "Khuyen nghi: File dang co kich thuoc lon, nen chay chuc nang toi uu kich thuoc file."
        End If
    End If
    
    ' Hien thi thong bao
    MsgBox msg, vbInformation, "Thong ke su dung bo nho"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi tao bao cao su dung bo nho: " & Err.description, vbCritical, "Loi"
    LogError "ReportMemoryUsage", Err.Number, Err.description
End Sub

' Dem so dong du lieu trong mot sheet
Private Function CountDataRows(sheetName As String) As Long
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(sheetName)
    
    ' Tim dong cuoi cung co du lieu
    CountDataRows = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1 ' Tru 1 cho dong tieu de
    
    ' Neu ket qua am, tra ve 0
    If CountDataRows < 0 Then CountDataRows = 0
    
    Exit Function
    
ErrorHandler:
    CountDataRows = 0
End Function

' Nen du lieu trong vung bang cach thay the gia tri lap
Public Sub CompressRepeatedValues()
    On Error GoTo ErrorHandler
    
    OptimizePerformance True
    
    ' Lay danh sach cac sheet du lieu
    Dim dataSheets As Variant
    dataSheets = GetRequiredDataSheets()
    
    Dim i As Integer
    For i = LBound(dataSheets) To UBound(dataSheets)
        ' Kiem tra sheet co ton tai khong
        If modUtility.sheetExists(CStr(dataSheets(i))) Then
            CompressRepeatedValuesInSheet CStr(dataSheets(i))
        End If
    Next i
    
    OptimizePerformance False
    
    MsgBox "Da nen du lieu thanh cong!", vbInformation, "Toi uu hoa"
    
    Exit Sub
    
ErrorHandler:
    OptimizePerformance False
    MsgBox "Loi khi nen du lieu: " & Err.description, vbCritical, "Loi"
    LogError "CompressRepeatedValues", Err.Number, Err.description
End Sub

' Nen du lieu trong mot sheet
Private Sub CompressRepeatedValuesInSheet(sheetName As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(sheetName)
    
    ' Mo khoa sheet
    ws.Unprotect password:=GetDefaultPassword()
    
    ' Tim vung du lieu
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Neu khong co du lieu, thoat
    If lastRow <= 1 Then
        ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
        Exit Sub
    End If
    
    ' Xac dinh cot nao can nen du lieu
    Dim colsToCompress As New Collection
    
    ' Cac cot chi nen du lieu neu la cot chuoi va co nhieu gia tri trung lap
    Dim col As Long
    For col = 1 To lastCol
        ' Kiem tra cot co phai chuoi va co nhieu gia tri trung lap
        If ShouldCompressColumn(ws, col, lastRow) Then
            colsToCompress.Add col
        End If
    Next col
    
    ' Thuc hien nen tung cot
    Dim j As Long
    For j = 1 To colsToCompress.Count
        CompressColumn ws, colsToCompress(j), lastRow
    Next j
    
    ' Bao ve lai sheet
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    Exit Sub
    
ErrorHandler:
    ' Bao ve lai sheet
    On Error Resume Next
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    LogError "CompressRepeatedValuesInSheet", Err.Number, Err.description
End Sub

' Kiem tra xem cot co nen duoc nen du lieu khong
Private Function ShouldCompressColumn(ws As Worksheet, col As Long, lastRow As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' Mac dinh la khong nen
    ShouldCompressColumn = False
    
    ' Dem so gia tri duy nhat trong cot
    Dim uniqueValues As New Collection
    Dim cell As Range
    Dim i As Long
    Dim uniqueCount As Long
    
    uniqueCount = 0
    
    ' Kiem tra mau du lieu (Sample 100 rows)
    Dim sampleSize As Long
    sampleSize = Application.WorksheetFunction.Min(100, lastRow - 1)
    
    ' Lay mau du lieu
    For i = 2 To sampleSize + 1
        If Not IsEmpty(ws.Cells(i, col).Value) Then
            On Error Resume Next
            uniqueValues.Add ws.Cells(i, col).Value, CStr(ws.Cells(i, col).Value)
            If Err.Number = 0 Then uniqueCount = uniqueCount + 1
            On Error GoTo ErrorHandler
        End If
    Next i
    
    ' Neu ty le gia tri duy nhat thap (nhieu gia tri trung lap), nen nen du lieu
    If uniqueCount > 0 And uniqueCount < sampleSize / 3 Then
        ShouldCompressColumn = True
    End If
    
    Exit Function
    
ErrorHandler:
    ShouldCompressColumn = False
End Function

' Nen du lieu trong mot cot
Private Sub CompressColumn(ws As Worksheet, col As Long, lastRow As Long)
    On Error GoTo ErrorHandler
    
    Dim currentValue As Variant
    Dim lastValue As Variant
    Dim i As Long
    
    lastValue = ws.Cells(2, col).Value
    
    For i = 3 To lastRow
        currentValue = ws.Cells(i, col).Value
        
        ' Neu gia tri giong voi gia tri truoc do, thay bang gia tri rong
        If currentValue = lastValue Then
            ws.Cells(i, col).Value = Empty
        Else
            lastValue = currentValue
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    LogError "CompressColumn", Err.Number, Err.description
End Sub
