Attribute VB_Name = "modConfig"
'Attribute VB_Name = "modConfig"
' Module cau hinh he thong
' Chua cac hang so va cai dat cau hinh
' Tac gia: Phong Khach hang Ca nhan, Agribank Chi nhanh 4
' Ngay tao: 11/05/2025
' Cap nhat: 15/05/2025 - Bo sung cac thong bao va hang so cho chuc nang sao luu va phuc hoi

Option Explicit

' ===== HANG SO CAC SHEET (TIENG VIET KHONG DAU) =====

' Sheet hien thi
Public Const SHEET_LOGIN As String = "Login"
Public Const SHEET_MAIN As String = "Main"
Public Const SHEET_CUSTOMER_VIEW As String = "CustomerView"
Public Const SHEET_LOAN_VIEW As String = "LoanView"
Public Const SHEET_ASSET_VIEW As String = "AssetView"
Public Const SHEET_TRANSACTION_HISTORY As String = "TransactionHistory"
Public Const SHEET_STAFF_MANAGEMENT As String = "StaffManagement"
Public Const SHEET_REPORTS As String = "Reports"
Public Const SHEET_SETTINGS As String = "Settings"

' Sheet du lieu (Hidden - Very Hidden)
Public Const SHEET_RAW_DU_NO As String = "Raw_DuNo"
Public Const SHEET_RAW_TAI_SAN As String = "Raw_TaiSan"
Public Const SHEET_RAW_TRA_GOC As String = "Raw_TraGoc"
Public Const SHEET_RAW_TRA_LAI As String = "Raw_TraLai"
Public Const SHEET_PROCESSED_DATA As String = "Processed_Data"
Public Const SHEET_IMPORT_LOG As String = "ImportLog"
Public Const SHEET_TRANSACTION_DATA As String = "TransactionData"
Public Const SHEET_STAFF_ASSIGNMENT As String = "StaffAssignment"

' Sheet bao mat (Very Hidden)
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_USERS As String = "Users"

' ===== CAU TRUC SHEET DU LIEU =====

' Cau truc sheet Raw_DuNo
Public Const COL_DU_NO_MA_KHOAN_VAY As String = "MaKhoanVay"
Public Const COL_DU_NO_MA_KHACH_HANG As String = "MaKhachHang"
Public Const COL_DU_NO_TEN_KHACH_HANG As String = "TenKhachHang"
' ... them cac cot khac

' Cau truc sheet Raw_TaiSan
Public Const COL_TAI_SAN_MA_TAI_SAN As String = "MaTaiSan"
Public Const COL_TAI_SAN_MA_KHACH_HANG As String = "MaKhachHang"
Public Const COL_TAI_SAN_TEN_KHACH_HANG As String = "TenKhachHang"
' ... them cac cot khac

' Cau truc sheet Raw_TraGoc
Public Const COL_TRA_GOC_MA_LICH_TRA_GOC As String = "MaLichTraGoc"
Public Const COL_TRA_GOC_MA_KHACH_HANG As String = "MaKhachHang"
Public Const COL_TRA_GOC_TEN_KHACH_HANG As String = "TenKhachHang"
' ... them cac cot khac

' Cau truc sheet Raw_TraLai
Public Const COL_TRA_LAI_MA_LICH_TRA_LAI As String = "MaLichTraLai"
Public Const COL_TRA_LAI_MA_KHACH_HANG As String = "MaKhachHang"
Public Const COL_TRA_LAI_TEN_KHACH_HANG As String = "TenKhachHang"
' ... them cac cot khac

' Cau truc sheet ImportLog
Public Const COL_IMPORT_LOG_ID As String = "ImportID"
Public Const COL_IMPORT_LOG_TEN_FILE As String = "TenFile"
Public Const COL_IMPORT_LOG_LOAI_DU_LIEU As String = "LoaiDuLieu"
Public Const COL_IMPORT_LOG_NGAY_TAO_FILE As String = "NgayTaoFile"
Public Const COL_IMPORT_LOG_THOI_GIAN_IMPORT As String = "ThoiGianImport"
Public Const COL_IMPORT_LOG_NGUOI_THUC_HIEN As String = "NguoiThucHien"
Public Const COL_IMPORT_LOG_TONG_BAN_GHI As String = "TongSoBanGhi"
Public Const COL_IMPORT_LOG_BAN_GHI_THEM_MOI As String = "SoBanGhiThemMoi"
Public Const COL_IMPORT_LOG_BAN_GHI_CAP_NHAT As String = "SoBanGhiCapNhat"
Public Const COL_IMPORT_LOG_BAN_GHI_XOA As String = "SoBanGhiXoa"
Public Const COL_IMPORT_LOG_TRANG_THAI As String = "TrangThai"
Public Const COL_IMPORT_LOG_GHI_CHU As String = "GhiChu"

' Cau truc sheet StaffAssignment
Public Const COL_STAFF_ASSIGNMENT_MA_KHACH_HANG As String = "MaKhachHang"
Public Const COL_STAFF_ASSIGNMENT_MA_CAN_BO As String = "MaCanBo"
Public Const COL_STAFF_ASSIGNMENT_NGAY_HIEU_LUC As String = "NgayHieuLuc"
Public Const COL_STAFF_ASSIGNMENT_NGAY_PHAN_CONG As String = "NgayPhanCong"
Public Const COL_STAFF_ASSIGNMENT_NGUOI_PHAN_CONG As String = "NguoiPhanCong"
Public Const COL_STAFF_ASSIGNMENT_GHI_CHU As String = "GhiChu"
Public Const COL_STAFF_ASSIGNMENT_MA_CAN_BO_TRUOC As String = "MaCanBoTruoc"
Public Const COL_STAFF_ASSIGNMENT_TRANG_THAI As String = "TrangThai"

' ===== CAU HINH BAO MAT =====
Public Const DEFAULT_PASSWORD As String = "agribank4" ' Mat khau mac dinh de bao ve sheet

' ===== CAU HINH MAU SAC =====
Public Const COLOR_HEADER As Long = 15773696 ' Mau do Agribank (RGB: 150, 28, 63)
Public Const COLOR_ALTERNATE_ROW As Long = 15921906 ' Mau xam nhat cho dong chan

' ===== CAU HINH FONT =====
Public Const DEFAULT_FONT_NAME As String = "Calibri"
Public Const DEFAULT_FONT_SIZE As Integer = 11
Public Const HEADER_FONT_SIZE As Integer = 12

' ===== THONG BAO HE THONG =====
Public Const MSG_ERROR_SHEET_NOT_FOUND As String = "Khong tim thay sheet: "
Public Const MSG_ERROR_SHEET_CREATE_FAILED As String = "Khong the tao sheet: "
Public Const MSG_INFO_STRUCTURE_INITIALIZED As String = "Cau truc du lieu da duoc khoi tao thanh cong!"
Public Const MSG_ERROR_STRUCTURE_FAILED As String = "Khong the khoi tao cau truc du lieu: "
Public Const MSG_CONFIRM_RECREATE_STRUCTURE As String = "Cau truc du lieu da ton tai. Ban co muon tao lai khong? Chu y: Du lieu hien tai se bi mat!"

' ===== THONG BAO SAO LUU VA PHUC HOI - THEM MOI GIAI DOAN 2 =====
Public Const MSG_BACKUP_SUCCESS As String = "Da sao luu thanh cong vao file: "
Public Const MSG_BACKUP_FAILED As String = "Khong the thuc hien sao luu. Loi: "
Public Const MSG_BACKUP_CONFIRM As String = "Ban co chac chan muon sao luu du lieu va cau truc hien tai khong?"
Public Const MSG_RESTORE_CONFIRM As String = "Ban co chac chan muon phuc hoi tu file sao luu nay? Luu y: Du lieu hien tai se bi mat!"
Public Const MSG_RESTORE_SUCCESS As String = "Da phuc hoi thanh cong tu file sao luu!"
Public Const MSG_RESTORE_FAILED As String = "Khong the phuc hoi du lieu. Loi: "
Public Const MSG_VALIDATE_STRUCTURE_ERROR As String = "Phat hien cau truc du lieu khong hop le. Can khoi phuc cau truc."
Public Const MSG_VALIDATE_MISSING_SHEETS As String = "Phat hien cac sheet bi thieu: "
Public Const MSG_REPAIR_SUCCESS As String = "Cau truc du lieu da duoc khoi phuc thanh cong."
Public Const MSG_REPAIR_BACKUP_FAILED As String = "Khong the sao luu du lieu hien tai. Qua trinh khoi phuc cau truc se khong duoc thuc hien."

' ===== HANG SO BAO MAT - THEM MOI GIAI DOAN 3 =====
Public Const ENCRYPTION_KEY As String = "AgribankSecure" ' Khoa ma hoa co ban
Public Const MIN_PASSWORD_LENGTH As Integer = 8 ' Do dai mat khau toi thieu
Public Const DEFAULT_USER_ROLE As String = "User" ' Quyen mac dinh cho nguoi dung moi

' ===== HAM TRUY XUAT CAU HINH =====

' Lay mat khau mac dinh cho bao ve sheet
Public Function GetDefaultPassword() As String
    On Error GoTo ErrorHandler
    
    ' Trong thuc te, mat khau nen duoc bao ve tot hon, vi du luu trong sheet Config voi dang ma hoa
    GetDefaultPassword = DEFAULT_PASSWORD
    Exit Function
    
ErrorHandler:
    MsgBox "Loi khi lay mat khau mac dinh: " & Err.description, vbCritical, "Loi"
    GetDefaultPassword = ""
End Function

' Lay danh sach cac sheet du lieu can thiet
Public Function GetRequiredDataSheets() As String()
    On Error GoTo ErrorHandler
    
    Dim sheets(0 To 7) As String
    sheets(0) = SHEET_RAW_DU_NO
    sheets(1) = SHEET_RAW_TAI_SAN
    sheets(2) = SHEET_RAW_TRA_GOC
    sheets(3) = SHEET_RAW_TRA_LAI
    sheets(4) = SHEET_PROCESSED_DATA
    sheets(5) = SHEET_IMPORT_LOG
    sheets(6) = SHEET_TRANSACTION_DATA
    sheets(7) = SHEET_STAFF_ASSIGNMENT
    
    GetRequiredDataSheets = sheets
    Exit Function
    
ErrorHandler:
    MsgBox "Loi khi lay danh sach sheet: " & Err.description, vbCritical, "Loi"
    ' Trong truong hop loi, tra ve mang rong
    Dim emptyArray(0 To 0) As String
    emptyArray(0) = ""
    GetRequiredDataSheets = emptyArray
End Function

' Lay danh sach cac sheet cau hinh can thiet
Public Function GetRequiredConfigSheets() As Variant
    On Error GoTo ErrorHandler
    
    Dim sheets As Variant
    sheets = Array(SHEET_CONFIG, SHEET_USERS)
    
    GetRequiredConfigSheets = sheets
    Exit Function
    
ErrorHandler:
    MsgBox "Loi khi lay danh sach sheet cau hinh: " & Err.description, vbCritical, "Loi"
    GetRequiredConfigSheets = Array()
End Function

' Lay mau header cho sheet
Public Function GetHeaderColor() As Long
    On Error GoTo ErrorHandler
    
    GetHeaderColor = COLOR_HEADER
    Exit Function
    
ErrorHandler:
    MsgBox "Loi khi lay mau header: " & Err.description, vbCritical, "Loi"
    GetHeaderColor = 5296274 ' Mau xam mac dinh
End Function
