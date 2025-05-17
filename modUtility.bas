Attribute VB_Name = "modUtility"
' Module tien ich
' Chua cac ham ho tro chung cho he thong
' Tac gia: Phong Khach hang Ca nhan, Agribank Chi nhanh 4
' Ngay tao: 11/05/2025

Option Explicit

' Tao ID duy nhat cho cac ban ghi
Public Function GenerateUniqueID() As String
    On Error GoTo ErrorHandler
    
    ' Tao ID duy nhat dua tren thoi gian hien tai va so ngau nhien
    Dim timeStamp As String
    timeStamp = Format(Now(), "yyyymmddhhnnss")
    
    Dim randomPart As String
    randomPart = Format(Int(Rnd() * 10000), "0000")
    
    GenerateUniqueID = timeStamp & randomPart
    Exit Function
    
ErrorHandler:
    MsgBox "Loi khi tao ID: " & Err.description, vbCritical, "Loi"
    LogError "GenerateUniqueID", Err.Number, Err.description
    GenerateUniqueID = "ERROR" & Format(Now(), "yyyymmddhhnnss")
End Function

' Lay gia tri cua mot cau hinh tu sheet Config
Public Function GetConfigValue(ByVal configName As String, Optional ByVal defaultValue As Variant = "") As Variant
    On Error GoTo ErrorHandler
    
    ' Kiem tra sheet Config co ton tai khong
    If Not sheetExists(SHEET_CONFIG) Then
        GetConfigValue = defaultValue
        Exit Function
    End If
    
    ' Tim kiem cau hinh trong sheet Config
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(SHEET_CONFIG)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow ' Bo qua dong header
        If ws.Cells(i, 1).Value = configName Then
            GetConfigValue = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    
    ' Neu khong tim thay, tra ve gia tri mac dinh
    GetConfigValue = defaultValue
    Exit Function
    
ErrorHandler:
    MsgBox "Loi khi doc cau hinh " & configName & ": " & Err.description, vbCritical, "Loi"
    LogError "GetConfigValue", Err.Number, Err.description
    GetConfigValue = defaultValue
End Function

' Luu gia tri cau hinh vao sheet Config
Public Sub SetConfigValue(ByVal configName As String, ByVal configValue As Variant, Optional ByVal description As String = "")
    On Error GoTo ErrorHandler
    
    ' Kiem tra sheet Config co ton tai khong
    If Not sheetExists(SHEET_CONFIG) Then
        MsgBox "Sheet Config khong ton tai!", vbExclamation, "Loi"
        Exit Sub
    End If
    
    ' Mo khoa sheet de cap nhat
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(SHEET_CONFIG)
    ws.Unprotect password:=GetDefaultPassword()
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Tim kiem cau hinh da ton tai
    Dim found As Boolean
    found = False
    
    Dim i As Long
    For i = 2 To lastRow ' Bo qua dong header
        If ws.Cells(i, 1).Value = configName Then
            ws.Cells(i, 2).Value = configValue
            If description <> "" Then ws.Cells(i, 3).Value = description
            found = True
            Exit For
        End If
    Next i
    
    ' Neu khong tim thay, them cau hinh moi
    If Not found Then
        ws.Cells(lastRow + 1, 1).Value = configName
        ws.Cells(lastRow + 1, 2).Value = configValue
        If description <> "" Then ws.Cells(lastRow + 1, 3).Value = description
    End If
    
    ' Cap nhat thoi gian cap nhat cuoi cung
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = "LAST_UPDATE" Then
            ws.Cells(i, 2).Value = Now()
            Exit For
        End If
    Next i
    
    ' Bao ve lai sheet
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi luu cau hinh " & configName & ": " & Err.description, vbCritical, "Loi"
    LogError "SetConfigValue", Err.Number, Err.description
    On Error Resume Next
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
End Sub

' Them mot dong vao sheet du lieu
Public Sub AddRowToDataSheet(ByVal sheetName As String, ByVal dataArray As Variant)
    On Error GoTo ErrorHandler
    
    ' Kiem tra sheet co ton tai khong
    If Not sheetExists(sheetName) Then
        MsgBox "Sheet " & sheetName & " khong ton tai!", vbExclamation, "Loi"
        Exit Sub
    End If
    
    ' Mo khoa sheet de cap nhat
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(sheetName)
    ws.Unprotect password:=GetDefaultPassword()
    
    ' Tim dong cuoi cung cua du lieu
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Them du lieu moi
    Dim i As Integer
    For i = LBound(dataArray) To UBound(dataArray)
        ws.Cells(lastRow + 1, i + 1).Value = dataArray(i)
    Next i
    
    ' Bao ve lai sheet
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi them dong moi vao sheet " & sheetName & ": " & Err.description, vbCritical, "Loi"
    LogError "AddRowToDataSheet", Err.Number, Err.description
    On Error Resume Next
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
End Sub

' Kiem tra xem sheet co ton tai khong (chuyen tu module DataStructure de su dung o cac noi khac)
Public Function sheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    sheetExists = Not (ThisWorkbook.sheets(sheetName) Is Nothing)
    On Error GoTo 0
End Function

' Ma hoa mat khau bang SHA-256 (day chi la mo phong, trong thuc te can su dung thu vien ma hoa tot hon)
Public Function HashPassword(ByVal password As String) As String
    On Error GoTo ErrorHandler
    
    ' Trong thuc te, can su dung thu vien ma hoa tot hon
    ' Ham nay chi la mo phong, khong nen dung trong san pham that
    Dim result As String
    Dim i As Long, char As Integer
    
    ' Mo phong ma hoa don gian
    For i = 1 To Len(password)
        char = Asc(Mid(password, i, 1))
        result = result & Hex(char * 2 + i)
    Next i
    
    ' Bo sung de dam bao co 64 ky tu
    While Len(result) < 64
        result = result & "0"
    Wend
    
    ' Cat bot neu vuot qua 64 ky tu
    If Len(result) > 64 Then
        result = Left(result, 64)
    End If
    
    HashPassword = result
    Exit Function
    
ErrorHandler:
    MsgBox "Loi khi ma hoa mat khau: " & Err.description, vbCritical, "Loi"
    LogError "HashPassword", Err.Number, Err.description
    HashPassword = ""
End Function

' Lay dong du lieu tu sheet dua tren gia tri cot tim kiem
Public Function GetRowByValue(ByVal sheetName As String, ByVal searchColumn As Integer, _
                             ByVal searchValue As Variant, ByVal returnColumn As Integer) As Variant
    On Error GoTo ErrorHandler
    
    ' Kiem tra sheet co ton tai khong
    If Not sheetExists(sheetName) Then
        GetRowByValue = Null
        Exit Function
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(sheetName)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, searchColumn).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow ' Bo qua dong header
        If ws.Cells(i, searchColumn).Value = searchValue Then
            If returnColumn = 0 Then
                ' Tra ve toan bo dong du lieu
                Dim rowData() As Variant
                Dim colCount As Integer, j As Integer
                colCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                ReDim rowData(1 To colCount)
                
                For j = 1 To colCount
                    rowData(j) = ws.Cells(i, j).Value
                Next j
                
                GetRowByValue = rowData
            Else
                ' Tra ve gia tri cua mot cot cu the
                GetRowByValue = ws.Cells(i, returnColumn).Value
            End If
            Exit Function
        End If
    Next i
    
    ' Neu khong tim thay
    GetRowByValue = Null
    Exit Function
    
ErrorHandler:
    MsgBox "Loi khi tim kiem du lieu trong sheet " & sheetName & ": " & Err.description, vbCritical, "Loi"
    LogError "GetRowByValue", Err.Number, Err.description
    GetRowByValue = Null
End Function

