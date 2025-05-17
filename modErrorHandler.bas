Attribute VB_Name = "modErrorHandler"
' Module xu ly loi
' Chua cac ham xu ly va ghi log loi
' Tac gia: Phong Khach hang Ca nhan, Agribank Chi nhanh 4
' Ngay tao: 11/05/2025

Option Explicit

' Ten file log
Private Const ERROR_LOG_FILE As String = "ErrorLog.txt"

' Ghi log loi vao file va sheet log
Public Sub LogError(ByVal source As String, ByVal errorNumber As Long, ByVal errorDescription As String)
    On Error Resume Next
    
    ' Tao chuoi thong tin loi
    Dim logMessage As String
    logMessage = Format(Now(), "yyyy-mm-dd hh:nn:ss") & " | " & _
                 "Source: " & source & " | " & _
                 "Error: " & errorNumber & " | " & _
                 "Description: " & errorDescription
    
    ' Ghi log vao file
    WriteToLogFile logMessage
    
    ' Ghi log vao sheet Error_Log neu co
    WriteToErrorLogSheet source, errorNumber, errorDescription
End Sub

' Ghi log vao file
Private Sub WriteToLogFile(ByVal logMessage As String)
    On Error Resume Next
    
    ' Xac dinh duong dan luu file log
    Dim logFilePath As String
    logFilePath = ThisWorkbook.Path & "\" & ERROR_LOG_FILE
    
    ' Kiem tra xem thu muc co ton tai khong
    If ThisWorkbook.Path = "" Then
        logFilePath = Environ("TEMP") & "\" & ERROR_LOG_FILE
    End If
    
    ' Mo file log de ghi
    Dim fileNum As Integer
    fileNum = FreeFile
    
    ' Mo file o che do Append (them vao cuoi file)
    Open logFilePath For Append As #fileNum
    
    ' Ghi thong tin loi vao file
    Print #fileNum, logMessage
    
    ' Dong file
    Close #fileNum
End Sub

' Ghi log vao sheet Error_Log
Private Sub WriteToErrorLogSheet(ByVal source As String, ByVal errorNumber As Long, ByVal errorDescription As String)
    On Error Resume Next
    
    ' Kiem tra xem sheet Error_Log co ton tai khong
    If Not sheetExists("Error_Log") Then
        ' Tao sheet Error_Log neu chua co
        CreateErrorLogSheet
    End If
    
    ' Mo khoa sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets("Error_Log")
    ws.Unprotect password:=GetDefaultPassword()
    
    ' Tim dong cuoi cung cua du lieu
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Ghi thong tin loi
    ws.Cells(lastRow + 1, 1).Value = Format(Now(), "yyyy-mm-dd hh:nn:ss")
    ws.Cells(lastRow + 1, 2).Value = source
    ws.Cells(lastRow + 1, 3).Value = errorNumber
    ws.Cells(lastRow + 1, 4).Value = errorDescription
    ws.Cells(lastRow + 1, 5).Value = Application.username
    
    ' Bao ve lai sheet
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
End Sub

' Tao sheet Error_Log neu chua co
Private Sub CreateErrorLogSheet()
    On Error Resume Next
    
    ' Them sheet moi
    ThisWorkbook.sheets.Add(After:=ThisWorkbook.sheets(ThisWorkbook.sheets.Count)).Name = "Error_Log"
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets("Error_Log")
    
    ' Thiet lap header
    With ws
        .Cells(1, 1).Value = "ThoiGian"
        .Cells(1, 2).Value = "Nguon"
        .Cells(1, 3).Value = "MaLoi"
        .Cells(1, 4).Value = "MoTaLoi"
        .Cells(1, 5).Value = "NguoiDung"
    End With
    
    ' Dinh dang header
    With ws.Range("A1:E1")
        .Font.Bold = True
        .Font.Size = HEADER_FONT_SIZE
        .Interior.Color = GetHeaderColor()
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
    End With
    
    ' Them dong vien
    ws.Range("A1:E1").Borders.LineStyle = xlContinuous
    
    ' Dong bang heading
    ws.Range("A1:E1").AutoFilter
    
    ' Tu dong dieu chinh do rong cot
    ws.Cells.EntireColumn.AutoFit
    
    ' An sheet va bao ve
    ws.Visible = xlSheetVeryHidden
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
End Sub

' Ham xu ly loi toan cuc cua ung dung
Public Sub GlobalErrorHandler(ByVal source As String, ByVal errorNumber As Long, ByVal errorDescription As String)
    On Error Resume Next
    
    ' Ghi log loi
    LogError source, errorNumber, errorDescription
    
    ' Hien thi thong bao loi cho nguoi dung
    MsgBox "Da xay ra loi trong qua trinh su dung ung dung:" & vbCrLf & _
           "- Nguon: " & source & vbCrLf & _
           "- Ma loi: " & errorNumber & vbCrLf & _
           "- Mo ta: " & errorDescription & vbCrLf & vbCrLf & _
           "Loi nay da duoc ghi lai trong log he thong.", _
           vbExclamation, "Loi ung dung"
End Sub

' Xu ly va hien thi thong tin loi
Public Sub DisplayErrorInfo(ByVal errorDescription As String, Optional ByVal errorTitle As String = "Loi")
    On Error Resume Next
    
    ' Ghi log loi
    LogError "DisplayErrorInfo", 0, errorDescription
    
    ' Hien thi thong bao loi cho nguoi dung
    MsgBox errorDescription, vbExclamation, errorTitle
End Sub

