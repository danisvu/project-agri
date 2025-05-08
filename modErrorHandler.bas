Attribute VB_Name = "modErrorHandler"
' Module xu ly loi
' Mo ta: Quan ly viec xu ly va ghi log loi trong he thong
' Tac gia: Phong KHCN, Agribank Chi nhanh 4
' Ngay tao: 08/05/2025

Option Explicit

' ===========================
' HANG SO
' ===========================

' Ten file log
Private Const ERROR_LOG_FILE As String = "ErrorLog.txt"

' Duong dan luu log
Private Const ERROR_LOG_PATH As String = "C:\Agribank\Logs\"

' Muc do loi
Public Enum ErrorSeverity
    ErrorSeverity_Low = 1       ' Loi it nghiem trong, khong anh huong den hoat dong chinh
    ErrorSeverity_Medium = 2    ' Loi co the gay ra van de nhung ung dung van hoat dong
    ErrorSeverity_High = 3      ' Loi nghiem trong, mot so chuc nang co the khong hoat dong
    ErrorSeverity_Critical = 4  ' Loi nguy hiem, co the gay crash ung dung
End Enum

' ===========================
' CAC HAM CHINH
' ===========================

' Ham ghi log loi co ban
Public Sub LogError(ByVal functionName As String, ByVal errorNumber As Long, ByVal errorDescription As String)
    LogErrorDetailed functionName, errorNumber, errorDescription, ErrorSeverity_Medium, ""
End Sub

' Ham ghi log loi chi tiet
Public Sub LogErrorDetailed(ByVal functionName As String, ByVal errorNumber As Long, _
                           ByVal errorDescription As String, ByVal severity As ErrorSeverity, _
                           Optional ByVal additionalInfo As String = "")
    On Error Resume Next
    
    Dim fileNum As Integer
    Dim logFilePath As String
    Dim logMessage As String
    Dim currentUser As String
    
    ' Kiem tra va tao thu muc log neu chua ton tai
    CreateLogFolder
    
    ' Xac dinh duong dan file log
    logFilePath = ERROR_LOG_PATH & Format(Date, "yyyy-mm-dd") & "_" & ERROR_LOG_FILE
    
    ' Lay thong tin nguoi dung hien tai
    currentUser = GetCurrentUser()
    
    ' Tao thong bao log
    logMessage = Format(Now, "yyyy-mm-dd hh:mm:ss") & vbTab & _
                 "User: " & currentUser & vbTab & _
                 "Function: " & functionName & vbTab & _
                 "Error: [" & errorNumber & "] " & errorDescription & vbTab & _
                 "Severity: " & severity
                 
    If additionalInfo <> "" Then
        logMessage = logMessage & vbTab & "Info: " & additionalInfo
    End If
    
    ' Ghi vao file log
    fileNum = FreeFile
    Open logFilePath For Append As #fileNum
    Print #fileNum, logMessage
    Close #fileNum
    
    ' Ghi vao immediate window de debug
    Debug.Print "ERROR: " & logMessage
    
    ' Neu loi nghiem trong, hien thi thong bao cho nguoi dung
    If severity >= ErrorSeverity_High Then
        ShowErrorMessage functionName, errorNumber, errorDescription
    End If
End Sub

' Ham tao thu muc log
Private Sub CreateLogFolder()
    On Error Resume Next
    
    ' Kiem tra va tao thu muc log neu chua ton tai
    If Dir(ERROR_LOG_PATH, vbDirectory) = "" Then
        MkDir ERROR_LOG_PATH
    End If
    
    On Error GoTo 0
End Sub

' Ham lay thong tin nguoi dung hien tai
Private Function GetCurrentUser() As String
    On Error Resume Next
    
    ' Neu co bien toan cuc luu thong tin nguoi dung
    If Not IsEmpty(gCurrentUser) And gCurrentUser <> "" Then
        GetCurrentUser = gCurrentUser
    Else
        ' Lay ten nguoi dung tu he thong
        GetCurrentUser = Application.UserName
    End If
    
    On Error GoTo 0
End Function

' Ham hien thi thong bao loi cho nguoi dung
Public Sub ShowErrorMessage(ByVal functionName As String, ByVal errorNumber As Long, ByVal errorDescription As String)
    On Error Resume Next
    
    Dim errorMsg As String
    
    ' Tao thong bao loi
    errorMsg = "Ðã x?y ra l?i trong quá trình th?c hi?n ch?c nang: " & functionName & vbCrLf & _
               "Mã l?i: " & errorNumber & vbCrLf & _
               "Mô t?: " & errorDescription & vbCrLf & vbCrLf & _
               "H? th?ng dã ghi nh?n l?i này vào log." & vbCrLf & _
               "Vui lòng liên h? b? ph?n IT d? du?c h? tr? n?u l?i v?n ti?p t?c x?y ra."
    
    ' Hien thi thong bao
    MsgBox errorMsg, vbExclamation, "L?i h? th?ng"
    
    On Error GoTo 0
End Sub

' Ham xu ly loi toan bo ung dung
Public Sub GlobalErrorHandler(ByVal errorNumber As Long, ByVal errorDescription As String, _
                             ByVal errorSource As String, ByVal errorLine As Long)
    On Error Resume Next
    
    Dim errorMsg As String
    Dim additionalInfo As String
    
    ' Tao thong tin bo sung
    additionalInfo = "Source: " & errorSource & ", Line: " & errorLine
    
    ' Ghi log loi
    LogErrorDetailed "GlobalErrorHandler", errorNumber, errorDescription, ErrorSeverity_Critical, additionalInfo
    
    ' Tao thong bao loi
    errorMsg = "Ðã x?y ra l?i nghiêm tr?ng trong h? th?ng:" & vbCrLf & _
               "Mã l?i: " & errorNumber & vbCrLf & _
               "Mô t?: " & errorDescription & vbCrLf & _
               "Ngu?n: " & errorSource & vbCrLf & vbCrLf & _
               "H? th?ng có th? không ho?t d?ng bình thu?ng." & vbCrLf & _
               "Vui lòng luu công vi?c c?a b?n (n?u có th?) và kh?i d?ng l?i ?ng d?ng." & vbCrLf & _
               "N?u l?i v?n ti?p t?c x?y ra, vui lòng liên h? b? ph?n IT."
    
    ' Hien thi thong bao
    MsgBox errorMsg, vbCritical, "L?i h? th?ng nghiêm tr?ng"
    
    On Error GoTo 0
End Sub

' Ham backup truoc khi thuc hien thao tac quan trong
Public Sub BackupBeforeAction(ByVal actionName As String)
    On Error GoTo ErrorHandler
    
    Dim backupPath As String
    Dim backupFileName As String
    
    ' Tao ten file backup
    backupFileName = "Backup_" & Format(Now, "yyyymmdd_hhmmss") & "_" & _
                     Replace(actionName, " ", "_") & ".xlsb"
    
    ' Duong dan day du
    backupPath = DEFAULT_BACKUP_PATH & backupFileName
    
    ' Kiem tra va tao thu muc backup neu chua ton tai
    If Dir(DEFAULT_BACKUP_PATH, vbDirectory) = "" Then
        MkDir DEFAULT_BACKUP_PATH
    End If
    
    ' Luu ban sao
    ThisWorkbook.SaveCopyAs backupPath
    
    ' Ghi log
    LogBackupAction actionName, backupPath, True, ""
    
    Exit Sub
    
ErrorHandler:
    ' Ghi log that bai
    LogBackupAction actionName, backupPath, False, Err.description
End Sub

' Ham ghi log backup
Private Sub LogBackupAction(ByVal actionName As String, ByVal backupPath As String, _
                           ByVal isSuccess As Boolean, ByVal errorMessage As String)
    On Error Resume Next
    
    Dim fileNum As Integer
    Dim logFilePath As String
    Dim logMessage As String
    Dim currentUser As String
    
    ' Kiem tra va tao thu muc log neu chua ton tai
    CreateLogFolder
    
    ' Xac dinh duong dan file log
    logFilePath = ERROR_LOG_PATH & "Backup_Log.txt"
    
    ' Lay thong tin nguoi dung hien tai
    currentUser = GetCurrentUser()
    
    ' Tao thong bao log
    logMessage = Format(Now, "yyyy-mm-dd hh:mm:ss") & vbTab & _
                 "User: " & currentUser & vbTab & _
                 "Action: " & actionName & vbTab & _
                 "Backup Path: " & backupPath & vbTab & _
                 "Status: " & IIf(isSuccess, "Success", "Failed")
                 
    If Not isSuccess Then
        logMessage = logMessage & vbTab & "Error: " & errorMessage
    End If
    
    ' Ghi vao file log
    fileNum = FreeFile
    Open logFilePath For Append As #fileNum
    Print #fileNum, logMessage
    Close #fileNum
End Sub

' Ham kiem tra va xu ly truong hop sheet bi mat
Public Function ValidateRequiredSheets() As Boolean
    On Error Resume Next
    
    Dim missingSheets As String
    Dim requiredSheets As Variant
    Dim i As Integer
    
    ' Danh sach cac sheet bat buoc
    requiredSheets = Array(SHEET_DU_NO, SHEET_TAI_SAN, SHEET_TRA_GOC, SHEET_TRA_LAI, _
                          SHEET_PROCESSED_DATA, SHEET_IMPORT_LOG, SHEET_TRANSACTION, _
                          SHEET_STAFF_ASSIGNMENT, SHEET_CONFIG, SHEET_USERS)
    
    ' Kiem tra tung sheet
    missingSheets = ""
    For i = LBound(requiredSheets) To UBound(requiredSheets)
        If sheetExists(CStr(requiredSheets(i))) = False Then
            If missingSheets <> "" Then missingSheets = missingSheets & ", "
            missingSheets = missingSheets & requiredSheets(i)
        End If
    Next i
    
    ' Xu ly ket qua
    If missingSheets <> "" Then
        ' Ghi log
        LogErrorDetailed "ValidateRequiredSheets", 0, "Missing required sheets: " & missingSheets, _
                        ErrorSeverity_High, "System integrity check"
        
        ' Hien thi thong bao
        MsgBox "C?u trúc d? li?u h? th?ng b? h?ng. Các sheet sau dây b? thi?u: " & vbCrLf & _
               missingSheets & vbCrLf & vbCrLf & _
               "H? th?ng s? c? g?ng khôi ph?c c?u trúc d? li?u. " & _
               "Vui lòng kh?i d?ng l?i ?ng d?ng sau khi quá trình khôi ph?c hoàn t?t.", _
               vbCritical, "L?i c?u trúc d? li?u"
        
        ' Thu hoi khoi phuc cau truc du lieu
        RecreateDataStructure
        
        ValidateRequiredSheets = False
    Else
        ValidateRequiredSheets = True
    End If
    
    On Error GoTo 0
End Function

' Ham kiem tra su ton tai cua sheet
Private Function sheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    sheetExists = Not ws Is Nothing
    
    On Error GoTo 0
End Function

' Ham khoi phuc cau truc du lieu
Private Sub RecreateDataStructure()
    On Error Resume Next
    
    ' Ham nay se goi lai ham khoi tao cau truc du lieu co ban trong modDataStructure
    ' Chu y: Day chi la giai phap tam thoi, du lieu co the bi mat
    
    ' Sao luu truoc khi khoi phuc
    BackupBeforeAction "RecreateDataStructure"
    
    ' Goi ham khoi tao lai cau truc
    Application.Run "InitializeDataStructure"
    
    ' Luu workbook
    ThisWorkbook.Save
    
    On Error GoTo 0
End Sub

' Ham ghi log su kien he thong
Public Sub LogSystemEvent(ByVal eventName As String, ByVal eventDetails As String, _
                         Optional ByVal isSuccess As Boolean = True)
    On Error Resume Next
    
    Dim fileNum As Integer
    Dim logFilePath As String
    Dim logMessage As String
    Dim currentUser As String
    
    ' Kiem tra va tao thu muc log neu chua ton tai
    CreateLogFolder
    
    ' Xac dinh duong dan file log
    logFilePath = ERROR_LOG_PATH & "System_Events.txt"
    
    ' Lay thong tin nguoi dung hien tai
    currentUser = GetCurrentUser()
    
    ' Tao thong bao log
    logMessage = Format(Now, "yyyy-mm-dd hh:mm:ss") & vbTab & _
                 "User: " & currentUser & vbTab & _
                 "Event: " & eventName & vbTab & _
                 "Status: " & IIf(isSuccess, "Success", "Failed") & vbTab & _
                 "Details: " & eventDetails
    
    ' Ghi vao file log
    fileNum = FreeFile
    Open logFilePath For Append As #fileNum
    Print #fileNum, logMessage
    Close #fileNum
End Sub

