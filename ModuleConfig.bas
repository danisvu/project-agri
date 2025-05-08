Attribute VB_Name = "ModuleConfig"
' Module cau hinh he thong
' Mo ta: Cung cap cac hang so va thiet lap cau hinh cho toan bo ung dung
' Tac gia: Phong KHCN, Agribank Chi nhanh 4
' Ngay tao: 08/05/2025

Option Explicit

' ===========================
' BIEN TOAN CUC
' ===========================

' Bien luu tru thong tin nguoi dung dang nhap hien tai
Public gCurrentUser As String
Public gCurrentUserName As String
Public gCurrentUserRole As String
Public gCurrentUserDept As String

' Bien theo doi trang thai du lieu
Public gDataLastImportDate As Date
Public gDataLastImportBy As String
Public gDataLastImportType As String

' Bien luu tru mau sac - Su dung bien thay vi hang so de tranh loi Overflow
Public COLOR_AGRIBANK_GREEN As Long   ' Mau xanh Agribank (#006E2A)
Public COLOR_BACKGROUND As Long       ' Mau trang
Public COLOR_WARNING As Long          ' Mau vang
Public COLOR_DANGER As Long           ' Mau do
Public COLOR_SUCCESS As Long          ' Mau xanh la
Public COLOR_INFO As Long             ' Mau xanh Agribank
Public COLOR_HEADER_BACKGROUND As Long ' Mau nen tieu de cot

' ===========================
' HANG SO HE THONG
' ===========================

' Thong tin phien ban
Public Const APP_NAME As String = "He thong Quan ly Thong tin Khach hang Vay"
Public Const APP_VERSION As String = "1.0"
Public Const APP_DATE As String = "08/05/2025"
Public Const APP_AUTHOR As String = "Agribank Chi nhanh 4"

' Cau hinh ung dung
Public Const ONE_KB As Long = 1024                ' 1 KB
Public Const ONE_MB As Long = 1048576            ' 1 MB (1024 * 1024)
Public Const MAX_IMPORT_FILE_SIZE As Long = 52428800 ' 50MB
Public Const MAX_RECORD_PROCESS As Long = 100000  ' So ban ghi toi da co the xu ly
Public Const DATA_WARNING_DAYS As Integer = 7     ' So ngay canh bao du lieu cu
Public Const LOAN_WARNING_DAYS As Integer = 30    ' So ngay canh bao truoc khi khoan vay dao han

' Cau hinh bao mat
Public Const PASSWORD_MIN_LENGTH As Integer = 8   ' Do dai toi thieu cua mat khau
Public Const PASSWORD_SALT As String = "Agribank2025" ' Chuoi salt cho ma hoa mat khau
Public Const LOGIN_MAX_ATTEMPTS As Integer = 3    ' So lan dang nhap toi da

' Cau hinh duong dan
Public Const DEFAULT_IMPORT_PATH As String = "C:\Agribank\Import\"
Public Const DEFAULT_EXPORT_PATH As String = "C:\Agribank\Export\"
Public Const DEFAULT_BACKUP_PATH As String = "C:\Agribank\Backup\"

' Ten cac sheet du lieu
Public Const SHEET_DU_NO As String = "Raw_DuNo"
Public Const SHEET_TAI_SAN As String = "Raw_TaiSan"
Public Const SHEET_TRA_GOC As String = "Raw_TraGoc"
Public Const SHEET_TRA_LAI As String = "Raw_TraLai"
Public Const SHEET_PROCESSED_DATA As String = "Processed_Data"
Public Const SHEET_IMPORT_LOG As String = "ImportLog"
Public Const SHEET_TRANSACTION As String = "TransactionHistory"
Public Const SHEET_STAFF_ASSIGNMENT As String = "StaffAssignment"
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_USERS As String = "Users"

' Ten cac sheet giao dien
Public Const SHEET_LOGIN As String = "Login"
Public Const SHEET_MAIN As String = "Main"
Public Const SHEET_CUSTOMER_VIEW As String = "CustomerView"
Public Const SHEET_LOAN_VIEW As String = "LoanView"
Public Const SHEET_ASSET_VIEW As String = "AssetView"
Public Const SHEET_TRANSACTION_VIEW As String = "TransactionHistory"
Public Const SHEET_STAFF_MANAGEMENT As String = "StaffManagement"
Public Const SHEET_REPORTS As String = "Reports"
Public Const SHEET_SETTINGS As String = "Settings"

' ===========================
' CAU HINH DINH DANG FILE
' ===========================

' Du no yyyy-mm-dd.xls
Public Const DU_NO_FILE_PREFIX As String = "Du no"
Public Const DU_NO_FILE_PATTERN As String = "Du no ????-??-??.xls"
Public Const DU_NO_DATE_PATTERN As String = "yyyy-mm-dd"

' Tai san yyyy-mm-dd.xls
Public Const TAI_SAN_FILE_PREFIX As String = "Tai san"
Public Const TAI_SAN_FILE_PATTERN As String = "Tai san ????-??-??.xls"
Public Const TAI_SAN_DATE_PATTERN As String = "yyyy-mm-dd"

' Tra goc mm-yyyy.xls
Public Const TRA_GOC_FILE_PREFIX As String = "Tra goc"
Public Const TRA_GOC_FILE_PATTERN As String = "Tra goc ??-????.xls"
Public Const TRA_GOC_DATE_PATTERN As String = "mm-yyyy"

' Tra lai mm-yyyy.xls
Public Const TRA_LAI_FILE_PREFIX As String = "Tra lai"
Public Const TRA_LAI_FILE_PATTERN As String = "Tra lai ??-????.xls"
Public Const TRA_LAI_DATE_PATTERN As String = "mm-yyyy"

' ===========================
' CAC PHAN LOAI VA MA
' ===========================

' Phan loai loai du lieu
Public Const DATA_TYPE_DU_NO As String = "DuNo"
Public Const DATA_TYPE_TAI_SAN As String = "TaiSan"
Public Const DATA_TYPE_TRA_GOC As String = "TraGoc"
Public Const DATA_TYPE_TRA_LAI As String = "TraLai"

' Phan loai giao dich
Public Const TRANS_TYPE_TRA_GOC As String = "TraGoc"
Public Const TRANS_TYPE_TRA_LAI As String = "TraLai"
Public Const TRANS_TYPE_GIAI_NGAN As String = "GiaiNgan"
Public Const TRANS_TYPE_TAT_TOAN As String = "TatToan"
Public Const TRANS_TYPE_GIA_HAN As String = "GiaHan"
Public Const TRANS_TYPE_CO_CAU_NO As String = "CoCauNo"
Public Const TRANS_TYPE_THAY_DOI_TS As String = "ThayDoiTaiSan"

' Trang thai cac loai du lieu
Public Const STATUS_ACTIVE As String = "Active"
Public Const STATUS_INACTIVE As String = "Inactive"
Public Const STATUS_PENDING As String = "Pending"
Public Const STATUS_PROCESSED As String = "Processed"
Public Const STATUS_ERROR As String = "Error"
Public Const STATUS_WARNING As String = "Warning"
Public Const STATUS_SUCCESS As String = "Success"

' Cac loai quyen nguoi dung
Public Const ROLE_ADMIN As String = "Admin"
Public Const ROLE_MANAGER As String = "Manager"
Public Const ROLE_SUPERVISOR As String = "Supervisor"
Public Const ROLE_USER As String = "User"

' Mat khau mac dinh cho workbook
Private Const DEFAULT_WORKBOOK_PASSWORD As String = "Agribank@2025"

' ===========================
' HAM CHINH
' ===========================

' Ham khoi tao thiet lap cau hinh
Public Sub InitializeConfig()
    On Error GoTo ErrorHandler
    
    ' Khoi tao cac bien mau sac
    COLOR_AGRIBANK_GREEN = RGB(0, 110, 42)     ' Mau xanh Agribank (#006E2A)
    COLOR_BACKGROUND = RGB(255, 255, 255)      ' Mau trang
    COLOR_WARNING = RGB(255, 192, 0)           ' Mau vang
    COLOR_DANGER = RGB(255, 0, 0)              ' Mau do
    COLOR_SUCCESS = RGB(80, 170, 0)            ' Mau xanh la
    COLOR_INFO = RGB(0, 110, 42)               ' Mau xanh Agribank
    COLOR_HEADER_BACKGROUND = RGB(0, 110, 42)  ' Mau nen tieu de cot
    
    ' Khoi tao cac bien toan cuc
    gCurrentUser = ""
    gCurrentUserName = ""
    gCurrentUserRole = ""
    gCurrentUserDept = ""
    
    ' Khoi tao thong tin du lieu
    gDataLastImportDate = DateSerial(1900, 1, 1) ' Gia tri mac dinh
    gDataLastImportBy = ""
    gDataLastImportType = ""
    
    ' Tao thu muc import/export/backup neu chua ton tai
    CreateFolderIfNotExists DEFAULT_IMPORT_PATH
    CreateFolderIfNotExists DEFAULT_EXPORT_PATH
    CreateFolderIfNotExists DEFAULT_BACKUP_PATH
    
    ' Doc thong tin cau hinh tu sheet Config
    LoadConfigFromSheet
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi khoi tao cau hinh: " & Err.description, vbExclamation, "Loi cau hinh"
    LogError "InitializeConfig", Err.Number, Err.description
End Sub

' Ham tao thu muc neu chua ton tai
Private Sub CreateFolderIfNotExists(ByVal folderPath As String)
    On Error Resume Next
    
    ' Kiem tra va tao thu muc neu chua ton tai
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    
    On Error GoTo 0
End Sub

' Ham doc cau hinh tu sheet Config
Private Sub LoadConfigFromSheet()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim configKey As String
    Dim configValue As String
    
    ' Tim sheet Config
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then Exit Sub
    
    ' Tim dong cuoi cung co du lieu
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Doc va xu ly cau hinh
    For i = 2 To lastRow ' Bat dau tu dong 2 (bo qua tieu de)
        configKey = ws.Cells(i, 1).Value
        configValue = ws.Cells(i, 2).Value
        
        ' Xu ly tung loai cau hinh
        Select Case configKey
            Case "DATA_WARNING_DAYS"
                If IsNumeric(configValue) Then
                    ' Cap nhat bien toan cuc tuong ung
                End If
                
            Case "LOAN_WARNING_DAYS"
                If IsNumeric(configValue) Then
                    ' Cap nhat bien toan cuc tuong ung
                End If
                
            ' Them cac truong hop khac tai day
        End Select
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi doc cau hinh tu sheet: " & Err.description, vbExclamation, "Loi cau hinh"
    LogError "LoadConfigFromSheet", Err.Number, Err.description
End Sub

' Ham ghi cau hinh vao sheet Config
Public Sub SaveConfigToSheet(ByVal configKey As String, ByVal configValue As String, Optional ByVal description As String = "")
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim foundRow As Long
    Dim found As Boolean
    Dim i As Long
    
    ' Tim sheet Config
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then Exit Sub
    
    ' Mo khoa bao ve sheet
    ws.Unprotect password:=DEFAULT_WORKBOOK_PASSWORD
    
    ' Tim dong cuoi cung co du lieu
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Tim xem cau hinh da ton tai chua
    found = False
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = configKey Then
            foundRow = i
            found = True
            Exit For
        End If
    Next i
    
    ' Cap nhat hoac them moi cau hinh
    If found Then
        ' Cap nhat cau hinh hien co
        ws.Cells(foundRow, 2).Value = configValue
        ws.Cells(foundRow, 4).Value = Now
        If description <> "" Then
            ws.Cells(foundRow, 3).Value = description
        End If
    Else
        ' Them cau hinh moi
        lastRow = lastRow + 1
        ws.Cells(lastRow, 1).Value = configKey
        ws.Cells(lastRow, 2).Value = configValue
        ws.Cells(lastRow, 3).Value = description
        ws.Cells(lastRow, 4).Value = Now
    End If
    
    ' Bao ve lai sheet
    ws.Protect password:=DEFAULT_WORKBOOK_PASSWORD, UserInterfaceOnly:=True
    
    Exit Sub
    
ErrorHandler:
    ' Bao ve lai sheet trong moi truong hop
    On Error Resume Next
    ws.Protect password:=DEFAULT_WORKBOOK_PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
    
    MsgBox "Loi khi luu cau hinh vao sheet: " & Err.description, vbExclamation, "Loi cau hinh"
    LogError "SaveConfigToSheet", Err.Number, Err.description
End Sub

' Ham doc gia tri cau hinh tu sheet
Public Function GetConfigValue(ByVal configKey As String, Optional ByVal defaultValue As String = "") As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Tim sheet Config
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        GetConfigValue = defaultValue
        Exit Function
    End If
    
    ' Tim dong cuoi cung co du lieu
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Tim gia tri cau hinh
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = configKey Then
            GetConfigValue = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    
    ' Neu khong tim thay, tra ve gia tri mac dinh
    GetConfigValue = defaultValue
    
    Exit Function
    
ErrorHandler:
    GetConfigValue = defaultValue
    LogError "GetConfigValue", Err.Number, Err.description
End Function

' Ham ghi log loi
Private Sub LogError(ByVal functionName As String, ByVal errorNumber As Long, ByVal errorDescription As String)
    ' Ham nay se duoc thuc hien day du trong modErrorHandler
    Debug.Print "Loi: " & functionName & " - " & errorNumber & " - " & errorDescription
End Sub

' Ham chuyen doi ngay tu ten file
Public Function ExtractDateFromFileName(ByVal fileName As String, ByVal fileType As String) As Date
    On Error GoTo ErrorHandler
    
    Dim datePart As String
    Dim dateFormat As String
    Dim filePrefix As String
    
    ' Xac dinh loai file va dinh dang ngay tuong ung
    Select Case fileType
        Case DATA_TYPE_DU_NO
            filePrefix = DU_NO_FILE_PREFIX
            dateFormat = DU_NO_DATE_PATTERN
        Case DATA_TYPE_TAI_SAN
            filePrefix = TAI_SAN_FILE_PREFIX
            dateFormat = TAI_SAN_DATE_PATTERN
        Case DATA_TYPE_TRA_GOC
            filePrefix = TRA_GOC_FILE_PREFIX
            dateFormat = TRA_GOC_DATE_PATTERN
        Case DATA_TYPE_TRA_LAI
            filePrefix = TRA_LAI_FILE_PREFIX
            dateFormat = TRA_LAI_DATE_PATTERN
        Case Else
            ' Loai file khong xac dinh, tra ve ngay mac dinh
            ExtractDateFromFileName = DateSerial(1900, 1, 1)
            Exit Function
    End Select
    
    ' Trich xuat phan ngay tu ten file
    fileName = Right(fileName, Len(fileName) - Len(filePrefix) - 1) ' Bo qua prefix va khoang trang
    datePart = left(fileName, InStr(1, fileName, ".") - 1) ' Lay phan truoc dau .
    
    ' Chuyen doi thanh ngay
    Select Case fileType
        Case DATA_TYPE_DU_NO, DATA_TYPE_TAI_SAN
            ' Dinh dang yyyy-mm-dd
            ExtractDateFromFileName = DateSerial(left(datePart, 4), Mid(datePart, 6, 2), Right(datePart, 2))
        Case DATA_TYPE_TRA_GOC, DATA_TYPE_TRA_LAI
            ' Dinh dang mm-yyyy
            ExtractDateFromFileName = DateSerial(Right(datePart, 4), left(datePart, 2), 1)
        Case Else
            ExtractDateFromFileName = DateSerial(1900, 1, 1)
    End Select
    
    Exit Function
    
ErrorHandler:
    ' Neu co loi, tra ve ngay mac dinh
    ExtractDateFromFileName = DateSerial(1900, 1, 1)
    LogError "ExtractDateFromFileName", Err.Number, Err.description
End Function

' Ham xac dinh loai du lieu tu ten file
Public Function DetermineFileType(ByVal fileName As String) As String
    On Error GoTo ErrorHandler
    
    Dim fileNameUpper As String
    
    ' Chuyen ten file thanh chu hoa de so sanh
    fileNameUpper = UCase(fileName)
    
    ' Xac dinh loai file
    If InStr(1, fileNameUpper, UCase(DU_NO_FILE_PREFIX)) = 1 Then
        DetermineFileType = DATA_TYPE_DU_NO
    ElseIf InStr(1, fileNameUpper, UCase(TAI_SAN_FILE_PREFIX)) = 1 Then
        DetermineFileType = DATA_TYPE_TAI_SAN
    ElseIf InStr(1, fileNameUpper, UCase(TRA_GOC_FILE_PREFIX)) = 1 Then
        DetermineFileType = DATA_TYPE_TRA_GOC
    ElseIf InStr(1, fileNameUpper, UCase(TRA_LAI_FILE_PREFIX)) = 1 Then
        DetermineFileType = DATA_TYPE_TRA_LAI
    Else
        ' Loai file khong xac dinh
        DetermineFileType = ""
    End If
    
    Exit Function
    
ErrorHandler:
    DetermineFileType = ""
    LogError "DetermineFileType", Err.Number, Err.description
End Function

' Ham lay ten sheet tuong ung voi loai du lieu
Public Function GetSheetNameForDataType(ByVal dataType As String) As String
    On Error GoTo ErrorHandler
    
    ' Xac dinh ten sheet tuong ung
    Select Case dataType
        Case DATA_TYPE_DU_NO
            GetSheetNameForDataType = SHEET_DU_NO
        Case DATA_TYPE_TAI_SAN
            GetSheetNameForDataType = SHEET_TAI_SAN
        Case DATA_TYPE_TRA_GOC
            GetSheetNameForDataType = SHEET_TRA_GOC
        Case DATA_TYPE_TRA_LAI
            GetSheetNameForDataType = SHEET_TRA_LAI
        Case Else
            ' Loai du lieu khong xac dinh
            GetSheetNameForDataType = ""
    End Select
    
    Exit Function
    
ErrorHandler:
    GetSheetNameForDataType = ""
    LogError "GetSheetNameForDataType", Err.Number, Err.description
End Function

' Ham lay gia tri mau tu ten
Public Function GetColorValue(ByVal colorName As String) As Long
    On Error GoTo ErrorHandler
    
    ' Lay gia tri mau tu ten
    Select Case colorName
        ' Hang so mau sac
        Case "COLOR_AGRIBANK_GREEN"
            GetColorValue = COLOR_AGRIBANK_GREEN
        Case "COLOR_BACKGROUND"
            GetColorValue = COLOR_BACKGROUND
        Case "COLOR_WARNING"
            GetColorValue = COLOR_WARNING
        Case "COLOR_DANGER"
            GetColorValue = COLOR_DANGER
        Case "COLOR_SUCCESS"
            GetColorValue = COLOR_SUCCESS
        Case "COLOR_INFO"
            GetColorValue = COLOR_INFO
        Case "COLOR_HEADER_BACKGROUND"
            GetColorValue = COLOR_HEADER_BACKGROUND
        Case Else
            ' Mau khong xac dinh, tra ve mau den
            GetColorValue = RGB(0, 0, 0)
    End Select
    
    Exit Function
    
ErrorHandler:
    ' Tra ve mau den neu co loi
    GetColorValue = RGB(0, 0, 0)
    LogError "GetColorValue", Err.Number, Err.description
End Function

' Ham lay ma khoa bao ve mac dinh
Public Function GetDefaultPassword() As String
    ' Trong thuc te, khong nen de mot hang so password truc tiep trong code
    ' Ma nen su dung mot co che bao mat tot hon
    GetDefaultPassword = DEFAULT_WORKBOOK_PASSWORD
End Function

' Ham chuyen doi mau tu he RGB sang Long
Public Function RGB2Long(ByVal red As Integer, ByVal green As Integer, ByVal blue As Integer) As Long
    RGB2Long = RGB(red, green, blue)
End Function

' Ham chuyen doi mau tu he Hex sang Long
Public Function Hex2Long(ByVal hexColor As String) As Long
    On Error GoTo ErrorHandler
    
    ' Loai bo ky tu # neu co
    If left(hexColor, 1) = "#" Then
        hexColor = Right(hexColor, Len(hexColor) - 1)
    End If
    
    Dim red As Integer, green As Integer, blue As Integer
    
    ' Mau hex phai co do dai 6 ky tu
    If Len(hexColor) = 6 Then
        red = CInt("&H" & Mid(hexColor, 1, 2))
        green = CInt("&H" & Mid(hexColor, 3, 2))
        blue = CInt("&H" & Mid(hexColor, 5, 2))
        
        Hex2Long = RGB(red, green, blue)
    Else
        ' Neu khong dung dinh dang, tra ve mau den
        Hex2Long = RGB(0, 0, 0)
    End If
    
    Exit Function
    
ErrorHandler:
    ' Tra ve mau den neu co loi
    Hex2Long = RGB(0, 0, 0)
    LogError "Hex2Long", Err.Number, Err.description
End Function
