Attribute VB_Name = "ModuleDataStructure"

' Module cau truc du lieu
' Mo ta: Module quan ly cau truc du lieu cua he thong
' Tac gia: Phong KHCN, Agribank Chi nhanh 4
' Ngay tao: 08/05/2025

Option Explicit

' ===========================
' HANG SO VA KHAI BAO
' ===========================

' Tham chieu cac hang so tu ModuleConfig

' Muc do nghiem trong cua loi
Private Const ErrorSeverity_Low As Integer = 1     ' Loi nhe, khong anh huong den hoat dong
Private Const ErrorSeverity_Medium As Integer = 2  ' Loi trung binh, anh huong mot phan chuc nang
Private Const ErrorSeverity_High As Integer = 3    ' Loi nghiem trong, anh huong den he thong
Private Const ErrorSeverity_Critical As Integer = 4 ' Loi tham hoa, he thong khong the hoat dong
Private Const ErrorSeverity_Info As Integer = 0    ' Thong tin, khong phai loi

' ===========================
' HAM CHINH
' ===========================

' Ham khoi tao cau truc du lieu
Public Sub InitializeDataStructure()
    On Error GoTo ErrorHandler
    
    ' Kiem tra cac sheet can thiet
    If Not ValidateRequiredSheets() Then
        ' Neu thieu sheet, tien hanh tao lai cau truc
        RecreateDataStructure
    End If
    
    ' Khoi tao cau truc du lieu cho cac sheet
    InitializeRawSheets
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi khoi tao cau truc du lieu: " & Err.description, vbCritical, "Loi he thong"
    LogErrorDetailed "InitializeDataStructure", Err.Number, Err.description, ErrorSeverity_High, "System initialization"
End Sub

' Ham kiem tra va xu ly truong hop sheet bi mat
Public Function ValidateRequiredSheets() As Boolean
    On Error Resume Next
    
    Dim missingSheets As String
    Dim requiredSheets As Variant
    Dim i As Integer
    
    ' Danh sach cac sheet bat buoc - Su dung tham chieu module cu the
    requiredSheets = Array(ModuleConfig.SHEET_DU_NO, ModuleConfig.SHEET_TAI_SAN, _
                          ModuleConfig.SHEET_TRA_GOC, ModuleConfig.SHEET_TRA_LAI, _
                          ModuleConfig.SHEET_PROCESSED_DATA, ModuleConfig.SHEET_IMPORT_LOG, _
                          ModuleConfig.SHEET_TRANSACTION, _
                          ModuleConfig.SHEET_STAFF_ASSIGNMENT, ModuleConfig.SHEET_CONFIG, ModuleConfig.SHEET_USERS)
    
    ' Kiem tra tung sheet
    missingSheets = ""
    For i = LBound(requiredSheets) To UBound(requiredSheets)
        If SheetExists(CStr(requiredSheets(i))) = False Then
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
        MsgBox "Cau truc du lieu he thong bi hong. Cac sheet sau day bi thieu: " & vbCrLf & _
               missingSheets & vbCrLf & vbCrLf & _
               "He thong se co gang khoi phuc cau truc du lieu. " & _
               "Vui long khoi dong lai ung dung sau khi qua trinh khoi phuc hoan tat.", _
               vbCritical, "Loi cau truc du lieu"
        
        ValidateRequiredSheets = False
    Else
        ValidateRequiredSheets = True
    End If
    
    On Error GoTo 0
End Function

' ===========================
' HAM HO TRO CAU TRUC DU LIEU
' ===========================

' Ham tao lai cau truc du lieu
Private Sub RecreateDataStructure()
    On Error GoTo ErrorHandler
    
    ' Khai bao bien
    Dim ws As Worksheet
    Dim requiredSheets As Variant
    Dim i As Integer
    
    ' Danh sach cac sheet can thiet
    requiredSheets = Array(ModuleConfig.SHEET_DU_NO, ModuleConfig.SHEET_TAI_SAN, _
                          ModuleConfig.SHEET_TRA_GOC, ModuleConfig.SHEET_TRA_LAI, _
                          ModuleConfig.SHEET_PROCESSED_DATA, ModuleConfig.SHEET_IMPORT_LOG, _
                          ModuleConfig.SHEET_TRANSACTION, _
                          ModuleConfig.SHEET_STAFF_ASSIGNMENT, ModuleConfig.SHEET_CONFIG, ModuleConfig.SHEET_USERS)
    
    ' Toi uu hoa hieu suat
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    
    ' Tao cac sheet neu chua ton tai
    For i = LBound(requiredSheets) To UBound(requiredSheets)
        If Not SheetExists(CStr(requiredSheets(i))) Then
            ' Cap nhat thanh trang thai
            Application.StatusBar = "Dang tao sheet " & requiredSheets(i) & "..."
            
            ' Them sheet moi
            Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            ws.Name = requiredSheets(i)
            
            ' Thiet lap cau truc mac dinh cho sheet
            InitializeSheetStructure ws.Name
            
            ' An sheet du lieu
            ws.Visible = xlSheetVeryHidden
        End If
    Next i
    
    ' Ghi log
    LogErrorDetailed "RecreateDataStructure", 0, "Data structure recreated successfully", _
                     ErrorSeverity_Info, "System initialization"
    
    ' Khoi phuc cai dat
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ' Khoi phuc cai dat trong moi truong hop
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    ' Thong bao loi
    MsgBox "Loi khi tao lai cau truc du lieu: " & Err.description, vbCritical, "Loi he thong"
    LogErrorDetailed "RecreateDataStructure", Err.Number, Err.description, ErrorSeverity_High, "System initialization"
End Sub

' Ham khoi tao cau truc cho cac sheet
Private Sub InitializeSheetStructure(ByVal sheetName As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' Thiet lap cau truc tuy thuoc vao loai sheet
    Select Case sheetName
        Case ModuleConfig.SHEET_DU_NO
            ' Tao tieu de cho sheet Du No
            ws.Cells(1, 1).Value = "MaKhoanVay"
            ws.Cells(1, 2).Value = "MaKhachHang"
            ws.Cells(1, 3).Value = "TenKhachHang"
            ws.Cells(1, 4).Value = "NgayPheDuyet"
            ws.Cells(1, 5).Value = "NgayDaoHan"
            ws.Cells(1, 6).Value = "SoTienPheDuyet"
            ws.Cells(1, 7).Value = "SoTienGiaiNgan"
            ws.Cells(1, 8).Value = "NgayGiaiNgan"
            ws.Cells(1, 9).Value = "LaiSuat"
            ws.Cells(1, 10).Value = "DuNoHienTai"
            ws.Cells(1, 11).Value = "TongSoTienDaTra"
            ws.Cells(1, 12).Value = "TongLaiDaTra"
            ws.Cells(1, 13).Value = "MaCanBoTinDung"
            ws.Cells(1, 14).Value = "TenCanBoTinDung"
            ws.Cells(1, 15).Value = "LoaiKhoanVay"
            ws.Cells(1, 16).Value = "PhanLoaiNo"
            ws.Cells(1, 17).Value = "NgayCapNhat"
            
        Case ModuleConfig.SHEET_TAI_SAN
            ' Tao tieu de cho sheet Tai San
            ws.Cells(1, 1).Value = "MaTaiSan"
            ws.Cells(1, 2).Value = "MaKhachHang"
            ws.Cells(1, 3).Value = "TenKhachHang"
            ws.Cells(1, 4).Value = "NgayCongChung"
            ws.Cells(1, 5).Value = "LoaiTaiSan"
            ws.Cells(1, 6).Value = "ChiTietTaiSan"
            ws.Cells(1, 7).Value = "SoLuong"
            ws.Cells(1, 8).Value = "DonViTinh"
            ws.Cells(1, 9).Value = "ViTriTaiSan"
            ws.Cells(1, 10).Value = "GiaTriTaiSan"
            ws.Cells(1, 11).Value = "TyLeChapNhan"
            ws.Cells(1, 12).Value = "GiaTriChapNhan"
            ws.Cells(1, 13).Value = "MaKhoanVay"
            ws.Cells(1, 14).Value = "NgayDinhGia"
            ws.Cells(1, 15).Value = "GhiChu"
            
        Case ModuleConfig.SHEET_TRA_GOC
            ' Tao tieu de cho sheet Tra Goc
            ws.Cells(1, 1).Value = "MaLichTraGoc"
            ws.Cells(1, 2).Value = "MaKhachHang"
            ws.Cells(1, 3).Value = "TenKhachHang"
            ws.Cells(1, 4).Value = "MaKhoanVay"
            ws.Cells(1, 5).Value = "NgayDenHan"
            ws.Cells(1, 6).Value = "SoTienGoc"
            ws.Cells(1, 7).Value = "DuNoConLai"
            ws.Cells(1, 8).Value = "TrangThai"
            ws.Cells(1, 9).Value = "NgayThanhToan"
            ws.Cells(1, 10).Value = "SoTienThanhToan"
            ws.Cells(1, 11).Value = "GhiChu"
            ws.Cells(1, 12).Value = "MaCanBoTinDung"
            ws.Cells(1, 13).Value = "TenCanBoTinDung"
            
        Case ModuleConfig.SHEET_TRA_LAI
            ' Tao tieu de cho sheet Tra Lai
            ws.Cells(1, 1).Value = "MaLichTraLai"
            ws.Cells(1, 2).Value = "MaKhachHang"
            ws.Cells(1, 3).Value = "TenKhachHang"
            ws.Cells(1, 4).Value = "MaKhoanVay"
            ws.Cells(1, 5).Value = "NgayDenHan"
            ws.Cells(1, 6).Value = "SoTienLai"
            ws.Cells(1, 7).Value = "TrangThai"
            ws.Cells(1, 8).Value = "NgayThanhToan"
            ws.Cells(1, 9).Value = "SoTienThanhToan"
            ws.Cells(1, 10).Value = "GhiChu"
            ws.Cells(1, 11).Value = "MaCanBoTinDung"
            ws.Cells(1, 12).Value = "TenCanBoTinDung"
            
        Case ModuleConfig.SHEET_PROCESSED_DATA
            ' Tao tieu de cho sheet du lieu tong hop
            ws.Cells(1, 1).Value = "MaKhachHang"
            ws.Cells(1, 2).Value = "TenKhachHang"
            ws.Cells(1, 3).Value = "TongSoKhoanVay"
            ws.Cells(1, 4).Value = "TongDuNo"
            ws.Cells(1, 5).Value = "TongTaiSan"
            ws.Cells(1, 6).Value = "TyLeTaiSanTrenDuNo"
            ws.Cells(1, 7).Value = "PhanLoaiRuiRo"
            ws.Cells(1, 8).Value = "SoNgayQuaHanLonNhat"
            ws.Cells(1, 9).Value = "MaCanBoTinDung"
            ws.Cells(1, 10).Value = "TenCanBoTinDung"
            ws.Cells(1, 11).Value = "ThoiGianCapNhat"
            
        Case ModuleConfig.SHEET_IMPORT_LOG
            ' Tao tieu de cho sheet Import Log
            ws.Cells(1, 1).Value = "ImportID"
            ws.Cells(1, 2).Value = "TenFile"
            ws.Cells(1, 3).Value = "LoaiDuLieu"
            ws.Cells(1, 4).Value = "NgayTaoFile"
            ws.Cells(1, 5).Value = "ThoiGianImport"
            ws.Cells(1, 6).Value = "NguoiThucHien"
            ws.Cells(1, 7).Value = "TongSoBanGhi"
            ws.Cells(1, 8).Value = "SoBanGhiThemMoi"
            ws.Cells(1, 9).Value = "SoBanGhiCapNhat"
            ws.Cells(1, 10).Value = "SoBanGhiXoa"
            ws.Cells(1, 11).Value = "TrangThai"
            ws.Cells(1, 12).Value = "GhiChu"
            
        Case ModuleConfig.SHEET_TRANSACTION
            ' Tao tieu de cho sheet Transaction History
            ws.Cells(1, 1).Value = "MaGiaoDich"
            ws.Cells(1, 2).Value = "MaKhachHang"
            ws.Cells(1, 3).Value = "TenKhachHang"
            ws.Cells(1, 4).Value = "MaKhoanVay"
            ws.Cells(1, 5).Value = "LoaiGiaoDich"
            ws.Cells(1, 6).Value = "NgayGiaoDich"
            ws.Cells(1, 7).Value = "SoTien"
            ws.Cells(1, 8).Value = "DuNoSauGiaoDich"
            ws.Cells(1, 9).Value = "GhiChu"
            ws.Cells(1, 10).Value = "NguoiThucHien"
            ws.Cells(1, 11).Value = "ThoiGianCapNhat"
            
        Case ModuleConfig.SHEET_STAFF_ASSIGNMENT
            ' Tao tieu de cho sheet Staff Assignment
            ws.Cells(1, 1).Value = "MaKhachHang"
            ws.Cells(1, 2).Value = "MaCanBo"
            ws.Cells(1, 3).Value = "NgayHieuLuc"
            ws.Cells(1, 4).Value = "NgayPhanCong"
            ws.Cells(1, 5).Value = "NguoiPhanCong"
            ws.Cells(1, 6).Value = "GhiChu"
            ws.Cells(1, 7).Value = "MaCanBoTruoc"
            ws.Cells(1, 8).Value = "TrangThai"
            
        Case ModuleConfig.SHEET_CONFIG
            ' Tao tieu de cho sheet Config
            ws.Cells(1, 1).Value = "ConfigKey"
            ws.Cells(1, 2).Value = "ConfigValue"
            ws.Cells(1, 3).Value = "Description"
            ws.Cells(1, 4).Value = "LastUpdated"
            
            ' Them cac cau hinh mac dinh
            ws.Cells(2, 1).Value = "DATA_WARNING_DAYS"
            ws.Cells(2, 2).Value = "7"
            ws.Cells(2, 3).Value = "So ngay canh bao du lieu cu"
            ws.Cells(2, 4).Value = Now
            
            ws.Cells(3, 1).Value = "LOAN_WARNING_DAYS"
            ws.Cells(3, 2).Value = "30"
            ws.Cells(3, 3).Value = "So ngay canh bao truoc khi khoan vay dao han"
            ws.Cells(3, 4).Value = Now
            
            ws.Cells(4, 1).Value = "DEFAULT_IMPORT_PATH"
            ws.Cells(4, 2).Value = "C:\Agribank\Import\"
            ws.Cells(4, 3).Value = "Duong dan mac dinh cho thu muc import"
            ws.Cells(4, 4).Value = Now
            
            ws.Cells(5, 1).Value = "DEFAULT_EXPORT_PATH"
            ws.Cells(5, 2).Value = "C:\Agribank\Export\"
            ws.Cells(5, 3).Value = "Duong dan mac dinh cho thu muc export"
            ws.Cells(5, 4).Value = Now
            
        Case ModuleConfig.SHEET_USERS
            ' Tao tieu de cho sheet Users
            ws.Cells(1, 1).Value = "UserID"
            ws.Cells(1, 2).Value = "UserName"
            ws.Cells(1, 3).Value = "Password"
            ws.Cells(1, 4).Value = "Role"
            ws.Cells(1, 5).Value = "Department"
            ws.Cells(1, 6).Value = "LastLogin"
            ws.Cells(1, 7).Value = "Status"
            ws.Cells(1, 8).Value = "CreatedBy"
            ws.Cells(1, 9).Value = "CreatedDate"
            
            ' Them tai khoan admin mac dinh
            ws.Cells(2, 1).Value = "admin"
            ws.Cells(2, 2).Value = "Administrator"
            ws.Cells(2, 3).Value = HashPassword("admin123") ' Ham nay can dinh nghia trong modSecurity
            ws.Cells(2, 4).Value = "Admin"
            ws.Cells(2, 5).Value = "IT"
            ws.Cells(2, 6).Value = Now
            ws.Cells(2, 7).Value = "Active"
            ws.Cells(2, 8).Value = "System"
            ws.Cells(2, 9).Value = Now
    End Select
    
    ' Dinh dang dong tieu de
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 20))
        .Font.Bold = True
        .Interior.Color = ModuleConfig.COLOR_HEADER_BACKGROUND
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
    End With
    
    ' Tu dong dieu chinh chieu rong cot
    ws.Cells.EntireColumn.AutoFit
    
    Exit Sub
    
ErrorHandler:
    LogErrorDetailed "InitializeSheetStructure", Err.Number, Err.description, ErrorSeverity_Medium, _
                    "Sheet: " & sheetName
End Sub

' Ham khoi tao cau truc cho cac sheet Raw
Private Sub InitializeRawSheets()
    On Error GoTo ErrorHandler
    
    ' Khoi tao cau truc cho cac sheet raw
    ' Can tuy chinh tuy theo yeu cau cu the
    
    Exit Sub
    
ErrorHandler:
    LogErrorDetailed "InitializeRawSheets", Err.Number, Err.description, ErrorSeverity_Medium, ""
End Sub

' ===========================
' CAC HAM TIEN ICH
' ===========================

' Ham kiem tra su ton tai cua sheet
Private Function SheetExists(ByVal sheetName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    SheetExists = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
    
    Exit Function
    
ErrorHandler:
    SheetExists = False
End Function

' Ham ma hoa mat khau
Private Function HashPassword(ByVal password As String) As String
    ' Day la ham don gian, can thay the bang phuong phap ma hoa tot hon
    Dim result As String
    Dim i As Integer
    Dim char As String
    
    result = ""
    For i = 1 To Len(password)
        char = Mid(password, i, 1)
        result = result & Chr(Asc(char) + 1)
    Next i
    
    HashPassword = result
End Function

' Ham ghi log loi
' Ham ghi log loi
Private Sub LogErrorDetailed(ByVal functionName As String, ByVal errorNumber As Long, _
                           ByVal errorDescription As String, ByVal severity As Integer, _
                           ByVal additionalInfo As String)
    ' Goi ham LogError tu modErrorHandler neu co
    On Error Resume Next
    
    ' Thu goi ham LogError tu modErrorHandler
    Call modErrorHandler.LogError(functionName, errorNumber, errorDescription)
    
    ' Kiem tra xem co loi khi goi hay khong
    If Err.Number <> 0 Then
        ' Xuat log ra Debug
        Debug.Print "Error in " & functionName & ": " & errorNumber & " - " & errorDescription & _
                   " [Severity: " & severity & "] " & additionalInfo
    End If
    
    On Error GoTo 0
End Sub

