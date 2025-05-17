Attribute VB_Name = "modDataStructure"
'Attribute VB_Name = "modDataStructure"
' Module quan ly cau truc du lieu
' Chiu trach nhiem tao va quan ly cau truc du lieu cua he thong
' Tac gia: Phong Khach hang Ca nhan, Agribank Chi nhanh 4
' Ngay tao: 11/05/2025

Option Explicit

' Khoi tao cau truc du lieu cua toan bo he thong
Public Sub InitializeDataStructure()
    On Error GoTo ErrorHandler
    
    ' Toi uu hoa hieu suat
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    
    ' Kiem tra xem cau truc co ton tai chua
    If CheckDataStructureExists Then
        Dim response As VbMsgBoxResult
        response = MsgBox(MSG_CONFIRM_RECREATE_STRUCTURE, vbYesNo + vbQuestion, "Xac nhan")
        
        If response = vbNo Then
            GoTo CleanUp
        End If
    End If
    
    ' Tao cac sheet du lieu
    Dim requiredSheets As Variant
    requiredSheets = GetRequiredDataSheets()
    
    Dim i As Integer
    For i = LBound(requiredSheets) To UBound(requiredSheets)
        InitializeSheetStructure CStr(requiredSheets(i))
    Next i
    
    ' Tao cac sheet cau hinh
    Dim configSheets As Variant
    configSheets = GetRequiredConfigSheets()
    
    For i = LBound(configSheets) To UBound(configSheets)
        InitializeConfigSheetStructure CStr(configSheets(i))
    Next i
    
    ' Tao cac Name Ranges
    CreateNameRanges
    
    ' Thong bao thanh cong
    MsgBox MSG_INFO_STRUCTURE_INITIALIZED, vbInformation, "Thanh cong"
    
CleanUp:
    ' Khoi phuc cai dat
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Exit Sub
    
ErrorHandler:
    MsgBox MSG_ERROR_STRUCTURE_FAILED & Err.description, vbCritical, "Loi"
    Resume CleanUp
End Sub

' Kiem tra xem cau truc du lieu co ton tai khong
Public Function CheckDataStructureExists() As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim allSheetsExist As Boolean   ' ??i t?n bi?n n?y t? sheetExists th?nh allSheetsExist
    allSheetsExist = True
    
    ' Kiem tra tung sheet theo ten co dinh thay vi dung mang
    If Not sheetExists(SHEET_RAW_DU_NO) Then allSheetsExist = False
    If Not sheetExists(SHEET_RAW_TAI_SAN) Then allSheetsExist = False
    If Not sheetExists(SHEET_RAW_TRA_GOC) Then allSheetsExist = False
    If Not sheetExists(SHEET_RAW_TRA_LAI) Then allSheetsExist = False
    If Not sheetExists(SHEET_PROCESSED_DATA) Then allSheetsExist = False
    If Not sheetExists(SHEET_IMPORT_LOG) Then allSheetsExist = False
    If Not sheetExists(SHEET_TRANSACTION_DATA) Then allSheetsExist = False
    If Not sheetExists(SHEET_STAFF_ASSIGNMENT) Then allSheetsExist = False
    
    CheckDataStructureExists = allSheetsExist
    Exit Function
    
ErrorHandler:
    MsgBox "Loi khi kiem tra cau truc du lieu: " & Err.description, vbCritical, "Loi"
    CheckDataStructureExists = False
End Function

' Kiem tra xem sheet co ton tai khong
Private Function sheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    sheetExists = Not (ThisWorkbook.sheets(sheetName) Is Nothing)
    On Error GoTo 0
End Function

' Khoi tao cau truc cho mot sheet du lieu
Public Sub InitializeSheetStructure(ByVal sheetName As String)
    On Error GoTo ErrorHandler
    
    ' Tao moi sheet neu chua ton tai, hoac xoa du lieu cu neu da ton tai
    If sheetExists(sheetName) Then
        ' Unprotect va xoa du lieu cu
        Dim ws As Worksheet
        Set ws = ThisWorkbook.sheets(sheetName)
        
        ' Unprotect truoc khi lam bat ky thu gi
        On Error Resume Next
        ws.Unprotect password:=GetDefaultPassword()
        On Error GoTo ErrorHandler
        
        ' Xoa du lieu cu nhung giu lai sheet
        ws.Cells.Clear
    Else
        ' Tao moi sheet
        ThisWorkbook.sheets.Add(After:=ThisWorkbook.sheets(ThisWorkbook.sheets.Count)).Name = sheetName
    End If
    
    ' Thiet lap cau truc cac cot cho sheet, tuy thuoc vao loai sheet
    Select Case sheetName
        Case SHEET_RAW_DU_NO
            SetupRawDuNoStructure
        
        Case SHEET_RAW_TAI_SAN
            SetupRawTaiSanStructure
        
        Case SHEET_RAW_TRA_GOC
            SetupRawTraGocStructure
        
        Case SHEET_RAW_TRA_LAI
            SetupRawTraLaiStructure
        
        Case SHEET_IMPORT_LOG
            SetupImportLogStructure
        
        Case SHEET_STAFF_ASSIGNMENT
            SetupStaffAssignmentStructure
        
        Case SHEET_PROCESSED_DATA
            SetupProcessedDataStructure
        
        Case SHEET_TRANSACTION_DATA
            SetupTransactionDataStructure
    End Select
    
    ' XU LY AN VA BAO VE SHEET - FIXED ORDER
    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.sheets(sheetName)
    
    ' DAM BAO SHEET MAIN TON TAI VA VISIBLE TRUOC KHI AN CAC SHEET KHAC
    EnsureMainSheetVisible
    
    ' AN SHEET TRUOC (khi chua protect)
    On Error Resume Next
    targetSheet.Visible = xlSheetVeryHidden
    If Err.Number <> 0 Then
        ' Neu khong the VeryHidden, thu Hidden
        Err.Clear
        targetSheet.Visible = xlSheetHidden
        If Err.Number <> 0 Then
            ' Neu van khong duoc, ghi log va de lai visible
            Debug.Print "Cannot hide sheet " & sheetName & " - leaving visible"
            Err.Clear
        End If
    End If
    On Error GoTo ErrorHandler
    
    ' BAO VE SHEET SAU KHI DA AN
    On Error Resume Next
    targetSheet.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi khoi tao cau truc sheet " & sheetName & ": " & Err.description, vbCritical, "Loi"
End Sub

' Khoi tao cau truc cho mot sheet cau hinh
Private Sub InitializeConfigSheetStructure(ByVal sheetName As String)
    On Error GoTo ErrorHandler
    
    ' Tao moi sheet neu chua ton tai, hoac xoa du lieu cu neu da ton tai
    If sheetExists(sheetName) Then
        ' Unprotect va xoa du lieu cu
        Dim ws As Worksheet
        Set ws = ThisWorkbook.sheets(sheetName)
        
        ' Unprotect truoc khi lam bat ky thu gi
        On Error Resume Next
        ws.Unprotect password:=GetDefaultPassword()
        On Error GoTo ErrorHandler
        
        ' Xoa du lieu cu nhung giu lai sheet
        ws.Cells.Clear
    Else
        ' Tao moi sheet
        ThisWorkbook.sheets.Add(After:=ThisWorkbook.sheets(ThisWorkbook.sheets.Count)).Name = sheetName
    End If
    
    ' Thiet lap cau truc cac cot cho sheet cau hinh
    Select Case sheetName
        Case SHEET_CONFIG
            SetupConfigStructure
        
        Case SHEET_USERS
            SetupUsersStructure
    End Select
    
    ' XU LY AN VA BAO VE SHEET - FIXED ORDER
    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.sheets(sheetName)
    
    ' DAM BAO SHEET MAIN TON TAI VA VISIBLE TRUOC KHI AN CAC SHEET KHAC
    EnsureMainSheetVisible
    
    ' AN SHEET TRUOC (khi chua protect)
    On Error Resume Next
    targetSheet.Visible = xlSheetVeryHidden
    If Err.Number <> 0 Then
        ' Neu khong the VeryHidden, thu Hidden
        Err.Clear
        targetSheet.Visible = xlSheetHidden
        If Err.Number <> 0 Then
            ' Neu van khong duoc, ghi log va de lai visible
            Debug.Print "Cannot hide config sheet " & sheetName & " - leaving visible"
            Err.Clear
        End If
    End If
    On Error GoTo ErrorHandler
    
    ' BAO VE SHEET SAU KHI DA AN (chat che hon cho config sheets)
    On Error Resume Next
    targetSheet.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi khoi tao cau truc sheet cau hinh " & sheetName & ": " & Err.description, vbCritical, "Loi"
End Sub

' Thiet lap cau truc cho sheet Raw_DuNo
Private Sub SetupRawDuNoStructure()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(SHEET_RAW_DU_NO)
    
    ' Thiet lap header
    With ws
        .Cells(1, 1).Value = COL_DU_NO_MA_KHOAN_VAY
        .Cells(1, 2).Value = COL_DU_NO_MA_KHACH_HANG
        .Cells(1, 3).Value = COL_DU_NO_TEN_KHACH_HANG
        .Cells(1, 4).Value = "NgayPheDuyet"
        .Cells(1, 5).Value = "NgayDaoHan"
        .Cells(1, 6).Value = "SoTienPheDuyet"
        .Cells(1, 7).Value = "SoTienGiaiNgan"
        .Cells(1, 8).Value = "LaiSuat"
        .Cells(1, 9).Value = "SoDuHienTai"
        .Cells(1, 10).Value = "NgayGiaiNgan"
        .Cells(1, 11).Value = "LoaiKhoanVay"
        .Cells(1, 12).Value = "TrangThai"
        .Cells(1, 13).Value = "MaCanBoTinDung"
        .Cells(1, 14).Value = "TenCanBoTinDung"
        .Cells(1, 15).Value = "MucDichVay"
        .Cells(1, 16).Value = "NguonVon"
        .Cells(1, 17).Value = "PhanLoaiNo"
        .Cells(1, 18).Value = "NgayTraGocGanNhat"
        .Cells(1, 19).Value = "NgayTraLaiGanNhat"
        .Cells(1, 20).Value = "NgayTraGocTiepTheo"
        .Cells(1, 21).Value = "NgayTraLaiTiepTheo"
        .Cells(1, 22).Value = "DiaChiKhachHang"
        .Cells(1, 23).Value = "SoDienThoai"
        .Cells(1, 24).Value = "GhiChu"
    End With
    
    ' Dinh dang header
    FormatSheetHeader ws
    
    ' Thiet lap cac dinh dang du lieu
    With ws
        ' Dinh dang cot ngay thang
        .Range("D:E,J:J,R:U").NumberFormat = "dd/mm/yyyy"
        
        ' Dinh dang cot tien te
        .Range("F:G,I:I").NumberFormat = "#,##0"
        
        ' Dinh dang cot lai suat
        .Range("H:H").NumberFormat = "0.00%"
        
        ' Tao bo loc
        .Range("A1:X1").AutoFilter
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi thiet lap cau truc Raw_DuNo: " & Err.description, vbCritical, "Loi"
End Sub

' Thiet lap cau truc cho sheet Raw_TaiSan
Private Sub SetupRawTaiSanStructure()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(SHEET_RAW_TAI_SAN)
    
    ' Thiet lap header
    With ws
        .Cells(1, 1).Value = COL_TAI_SAN_MA_TAI_SAN
        .Cells(1, 2).Value = COL_TAI_SAN_MA_KHACH_HANG
        .Cells(1, 3).Value = COL_TAI_SAN_TEN_KHACH_HANG
        .Cells(1, 4).Value = "NgayCongChung"
        .Cells(1, 5).Value = "NgayQuanLy"
        .Cells(1, 6).Value = "LoaiTaiSan"
        .Cells(1, 7).Value = "LoaiChiTietTaiSan"
        .Cells(1, 8).Value = "SoLuong"
        .Cells(1, 9).Value = "DonViTinh"
        .Cells(1, 10).Value = "ViTriTaiSan"
        .Cells(1, 11).Value = "GiaTriTaiSan"
        .Cells(1, 12).Value = "LoaiTheChan"
        .Cells(1, 13).Value = "NgayTheChan"
        .Cells(1, 14).Value = "NgayHetHan"
        .Cells(1, 15).Value = "TyLeGiaTriKhaDung"
        .Cells(1, 16).Value = "GiaTriKhaDung"
        .Cells(1, 17).Value = "GiaTriTheChan"
        .Cells(1, 18).Value = "MaKhoanVay"
        .Cells(1, 19).Value = "TrangThai"
        .Cells(1, 20).Value = "GhiChu"
    End With
    
    ' Dinh dang header
    FormatSheetHeader ws
    
    ' Thiet lap cac dinh dang du lieu
    With ws
        ' Dinh dang cot ngay thang
        .Range("D:E,M:N").NumberFormat = "dd/mm/yyyy"
        
        ' Dinh dang cot tien te
        .Range("K:K,P:Q").NumberFormat = "#,##0"
        
        ' Dinh dang cot ty le
        .Range("O:O").NumberFormat = "0.00%"
        
        ' Tao bo loc
        .Range("A1:T1").AutoFilter
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi thiet lap cau truc Raw_TaiSan: " & Err.description, vbCritical, "Loi"
End Sub

' Thiet lap cau truc cho sheet Raw_TraGoc
Private Sub SetupRawTraGocStructure()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(SHEET_RAW_TRA_GOC)
    
    ' Thiet lap header
    With ws
        .Cells(1, 1).Value = COL_TRA_GOC_MA_LICH_TRA_GOC
        .Cells(1, 2).Value = COL_TRA_GOC_MA_KHACH_HANG
        .Cells(1, 3).Value = COL_TRA_GOC_TEN_KHACH_HANG
        .Cells(1, 4).Value = "MaKhoanVay"
        .Cells(1, 5).Value = "NgayDenHan"
        .Cells(1, 6).Value = "SoTienPhaiTra"
        .Cells(1, 7).Value = "SoDuHienTai"
        .Cells(1, 8).Value = "TaiKhoan"
        .Cells(1, 9).Value = "MaGiaoDich"
        .Cells(1, 10).Value = "NgayGiaoDich"
        .Cells(1, 11).Value = "NgayCapNhat"
        .Cells(1, 12).Value = "TrangThai"
        .Cells(1, 13).Value = "NguoiXuLy"
        .Cells(1, 14).Value = "NguoiPheDuyet"
        .Cells(1, 15).Value = "GhiChu"
        .Cells(1, 16).Value = "DaXuLy"
    End With
    
    ' Dinh dang header
    FormatSheetHeader ws
    
    ' Thiet lap cac dinh dang du lieu
    With ws
        ' Dinh dang cot ngay thang
        .Range("E:E,J:K").NumberFormat = "dd/mm/yyyy"
        
        ' Dinh dang cot tien te
        .Range("F:G").NumberFormat = "#,##0"
        
        ' Tao bo loc
        .Range("A1:P1").AutoFilter
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi thiet lap cau truc Raw_TraGoc: " & Err.description, vbCritical, "Loi"
End Sub

' Thiet lap cau truc cho sheet Raw_TraLai
Private Sub SetupRawTraLaiStructure()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(SHEET_RAW_TRA_LAI)
    
    ' Thiet lap header
    With ws
        .Cells(1, 1).Value = COL_TRA_LAI_MA_LICH_TRA_LAI
        .Cells(1, 2).Value = COL_TRA_LAI_MA_KHACH_HANG
        .Cells(1, 3).Value = COL_TRA_LAI_TEN_KHACH_HANG
        .Cells(1, 4).Value = "MaKhoanVay"
        .Cells(1, 5).Value = "NgayDenHan"
        .Cells(1, 6).Value = "SoTienPhaiTra"
        .Cells(1, 7).Value = "SoDuHienTai"
        .Cells(1, 8).Value = "TaiKhoan"
        .Cells(1, 9).Value = "MaGiaoDich"
        .Cells(1, 10).Value = "NgayGiaoDich"
        .Cells(1, 11).Value = "NgayCapNhat"
        .Cells(1, 12).Value = "TrangThai"
        .Cells(1, 13).Value = "NguoiXuLy"
        .Cells(1, 14).Value = "NguoiPheDuyet"
        .Cells(1, 15).Value = "GhiChu"
        .Cells(1, 16).Value = "DaXuLy"
    End With
    
    ' Dinh dang header
    FormatSheetHeader ws
    
    ' Thiet lap cac dinh dang du lieu
    With ws
        ' Dinh dang cot ngay thang
        .Range("E:E,J:K").NumberFormat = "dd/mm/yyyy"
        
        ' Dinh dang cot tien te
        .Range("F:G").NumberFormat = "#,##0"
        
        ' Tao bo loc
        .Range("A1:P1").AutoFilter
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi thiet lap cau truc Raw_TraLai: " & Err.description, vbCritical, "Loi"
End Sub

' Thiet lap cau truc cho sheet ImportLog
Private Sub SetupImportLogStructure()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(SHEET_IMPORT_LOG)
    
    ' Thiet lap header
    With ws
        .Cells(1, 1).Value = COL_IMPORT_LOG_ID
        .Cells(1, 2).Value = COL_IMPORT_LOG_TEN_FILE
        .Cells(1, 3).Value = COL_IMPORT_LOG_LOAI_DU_LIEU
        .Cells(1, 4).Value = COL_IMPORT_LOG_NGAY_TAO_FILE
        .Cells(1, 5).Value = COL_IMPORT_LOG_THOI_GIAN_IMPORT
        .Cells(1, 6).Value = COL_IMPORT_LOG_NGUOI_THUC_HIEN
        .Cells(1, 7).Value = COL_IMPORT_LOG_TONG_BAN_GHI
        .Cells(1, 8).Value = COL_IMPORT_LOG_BAN_GHI_THEM_MOI
        .Cells(1, 9).Value = COL_IMPORT_LOG_BAN_GHI_CAP_NHAT
        .Cells(1, 10).Value = COL_IMPORT_LOG_BAN_GHI_XOA
        .Cells(1, 11).Value = COL_IMPORT_LOG_TRANG_THAI
        .Cells(1, 12).Value = COL_IMPORT_LOG_GHI_CHU
    End With
    
    ' Dinh dang header
    FormatSheetHeader ws
    
    ' Thiet lap cac dinh dang du lieu
    With ws
        ' Dinh dang cot ngay thang
        .Range("D:E").NumberFormat = "dd/mm/yyyy hh:mm:ss"
        
        ' Tao bo loc
        .Range("A1:L1").AutoFilter
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi thiet lap cau truc ImportLog: " & Err.description, vbCritical, "Loi"
End Sub

' Thiet lap cau truc cho sheet StaffAssignment
Private Sub SetupStaffAssignmentStructure()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(SHEET_STAFF_ASSIGNMENT)
    
    ' Thiet lap header
    With ws
        .Cells(1, 1).Value = COL_STAFF_ASSIGNMENT_MA_KHACH_HANG
        .Cells(1, 2).Value = COL_STAFF_ASSIGNMENT_MA_CAN_BO
        .Cells(1, 3).Value = COL_STAFF_ASSIGNMENT_NGAY_HIEU_LUC
        .Cells(1, 4).Value = COL_STAFF_ASSIGNMENT_NGAY_PHAN_CONG
        .Cells(1, 5).Value = COL_STAFF_ASSIGNMENT_NGUOI_PHAN_CONG
        .Cells(1, 6).Value = COL_STAFF_ASSIGNMENT_GHI_CHU
        .Cells(1, 7).Value = COL_STAFF_ASSIGNMENT_MA_CAN_BO_TRUOC
        .Cells(1, 8).Value = COL_STAFF_ASSIGNMENT_TRANG_THAI
    End With
    
    ' Dinh dang header
    FormatSheetHeader ws
    
    ' Thiet lap cac dinh dang du lieu
    With ws
        ' Dinh dang cot ngay thang
        .Range("C:D").NumberFormat = "dd/mm/yyyy hh:mm:ss"
        
        ' Tao bo loc
        .Range("A1:H1").AutoFilter
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi thiet lap cau truc StaffAssignment: " & Err.description, vbCritical, "Loi"
End Sub

' Thiet lap cau truc cho sheet Processed_Data
Private Sub SetupProcessedDataStructure()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(SHEET_PROCESSED_DATA)
    
    ' Thiet lap header
    With ws
        .Cells(1, 1).Value = "MaKhachHang"
        .Cells(1, 2).Value = "TenKhachHang"
        .Cells(1, 3).Value = "LoaiKhachHang"
        .Cells(1, 4).Value = "DiaChiKhachHang"
        .Cells(1, 5).Value = "SoDienThoai"
        .Cells(1, 6).Value = "TongSoKhoanVay"
        .Cells(1, 7).Value = "TongDuNo"
        .Cells(1, 8).Value = "TongGiaTriTaiSan"
        .Cells(1, 9).Value = "MaCanBoTinDung"
        .Cells(1, 10).Value = "TenCanBoTinDung"
        .Cells(1, 11).Value = "PhanLoaiNoXauNhat"
        .Cells(1, 12).Value = "NgayCapNhat"
        .Cells(1, 13).Value = "GhiChu"
    End With
    
    ' Dinh dang header
    FormatSheetHeader ws
    
    ' Thiet lap cac dinh dang du lieu
    With ws
        ' Dinh dang cot ngay thang
        .Range("L:L").NumberFormat = "dd/mm/yyyy hh:mm:ss"
        
        ' Dinh dang cot tien te
        .Range("G:H").NumberFormat = "#,##0"
        
        ' Tao bo loc
        .Range("A1:M1").AutoFilter
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi thiet lap cau truc Processed_Data: " & Err.description, vbCritical, "Loi"
End Sub

' Thiet lap cau truc cho sheet TransactionData
Private Sub SetupTransactionDataStructure()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(SHEET_TRANSACTION_DATA)
    
    ' Thiet lap header
    With ws
        .Cells(1, 1).Value = "MaGiaoDich"
        .Cells(1, 2).Value = "MaKhachHang"
        .Cells(1, 3).Value = "TenKhachHang"
        .Cells(1, 4).Value = "MaKhoanVay"
        .Cells(1, 5).Value = "LoaiGiaoDich"
        .Cells(1, 6).Value = "NgayGiaoDich"
        .Cells(1, 7).Value = "SoTien"
        .Cells(1, 8).Value = "NguoiThucHien"
        .Cells(1, 9).Value = "GhiChu"
        .Cells(1, 10).Value = "TrangThai"
    End With
    
    ' Dinh dang header
    FormatSheetHeader ws
    
    ' Thiet lap cac dinh dang du lieu
    With ws
        ' Dinh dang cot ngay thang
        .Range("F:F").NumberFormat = "dd/mm/yyyy hh:mm:ss"
        
        ' Dinh dang cot tien te
        .Range("G:G").NumberFormat = "#,##0"
        
        ' Tao bo loc
        .Range("A1:J1").AutoFilter
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi thiet lap cau truc TransactionData: " & Err.description, vbCritical, "Loi"
End Sub

' Thiet lap cau truc cho sheet Config
Private Sub SetupConfigStructure()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(SHEET_CONFIG)
    
    ' Thiet lap header
    With ws
        .Cells(1, 1).Value = "TenCauHinh"
        .Cells(1, 2).Value = "GiaTri"
        .Cells(1, 3).Value = "MoTa"
    End With
    
    ' Dinh dang header
    FormatSheetHeader ws
    
    ' Them mot so cau hinh mac dinh
    With ws
        .Cells(2, 1).Value = "VERSION"
        .Cells(2, 2).Value = "1.0"
        .Cells(2, 3).Value = "Phien ban he thong"
        
        .Cells(3, 1).Value = "ORGANIZATION"
        .Cells(3, 2).Value = "Agribank Chi nhanh 4"
        .Cells(3, 3).Value = "To chuc su dung"
        
        .Cells(4, 1).Value = "LAST_UPDATE"
        .Cells(4, 2).Value = Now()
        .Cells(4, 3).Value = "Lan cap nhat cuoi cung"
        
        .Range("B4").NumberFormat = "dd/mm/yyyy hh:mm:ss"
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi thiet lap cau truc Config: " & Err.description, vbCritical, "Loi"
End Sub

' Thiet lap cau truc cho sheet Users
Private Sub SetupUsersStructure()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(SHEET_USERS)
    
    ' Thiet lap header
    With ws
        .Cells(1, 1).Value = "ID"
        .Cells(1, 2).Value = "TenDangNhap"
        .Cells(1, 3).Value = "MatKhau"
        .Cells(1, 4).Value = "HoTen"
        .Cells(1, 5).Value = "ChucVu"
        .Cells(1, 6).Value = "PhongBan"
        .Cells(1, 7).Value = "QuyenHan"
        .Cells(1, 8).Value = "TrangThai"
        .Cells(1, 9).Value = "NgayTao"
        .Cells(1, 10).Value = "NguoiTao"
        .Cells(1, 11).Value = "LanDangNhapCuoi"
    End With
    
    ' Dinh dang header
    FormatSheetHeader ws
    
    ' Them tai khoan admin mac dinh
    With ws
        .Cells(2, 1).Value = 1
        .Cells(2, 2).Value = "admin"
        .Cells(2, 3).Value = "8c6976e5b5410415bde908bd4dee15dfb167a9c873fc4bb8a81f6f2ab448a918" ' Hash cua "admin"
        .Cells(2, 4).Value = "Administrator"
        .Cells(2, 5).Value = "Quan tri vien"
        .Cells(2, 6).Value = "IT"
        .Cells(2, 7).Value = "Admin"
        .Cells(2, 8).Value = "Active"
        .Cells(2, 9).Value = Now()
        .Cells(2, 10).Value = "System"
        .Cells(2, 11).Value = Now()
    End With
    
    ' Dinh dang cot ngay thang
    ws.Range("I:I,K:K").NumberFormat = "dd/mm/yyyy hh:mm:ss"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi thiet lap cau truc Users: " & Err.description, vbCritical, "Loi"
End Sub

' Dinh dang header cho sheet
Private Sub FormatSheetHeader(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Dinh dang header
    With ws.Range("A1").CurrentRegion.Rows(1)
        .Font.Bold = True
        .Font.Size = HEADER_FONT_SIZE
        .Interior.Color = GetHeaderColor()
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
    End With
    
    ' Them dong vien
    ws.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
    
    ' Dong bang heading
    ws.Range("A1").CurrentRegion.Rows(1).AutoFilter
    
    ' Tu dong dieu chinh do rong cot
    ws.Cells.EntireColumn.AutoFit
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi dinh dang header: " & Err.description, vbCritical, "Loi"
End Sub

' Kiem tra va khoi phuc cau truc du lieu
Public Sub ValidateDataStructure()
    On Error GoTo ErrorHandler
    
    ' Kiem tra cac sheet du lieu can thiet
    Dim requiredSheets As Variant
    requiredSheets = GetRequiredDataSheets()
    
    Dim i As Integer
    Dim missingSheets As String
    missingSheets = ""
    
    For i = LBound(requiredSheets) To UBound(requiredSheets)
        If Not sheetExists(CStr(requiredSheets(i))) Then
            missingSheets = missingSheets & CStr(requiredSheets(i)) & ", "
        End If
    Next i
    
    ' Kiem tra cac sheet cau hinh can thiet
    Dim configSheets As Variant
    configSheets = GetRequiredConfigSheets()
    
    For i = LBound(configSheets) To UBound(configSheets)
        If Not sheetExists(CStr(configSheets(i))) Then
            missingSheets = missingSheets & CStr(configSheets(i)) & ", "
        End If
    Next i
    
    ' Neu co sheet bi thieu
    If missingSheets <> "" Then
        missingSheets = Left(missingSheets, Len(missingSheets) - 2) ' Loai bo dau phay cuoi cung
        
        Dim response As VbMsgBoxResult
        response = MsgBox("Phat hien sheet bi thieu: " & missingSheets & ". Ban co muon khoi phuc cau truc du lieu khong?", _
                         vbYesNo + vbQuestion, "Canh bao")
        
        If response = vbYes Then
            InitializeDataStructure
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi kiem tra cau truc du lieu: " & Err.description, vbCritical, "Loi"
End Sub

' Tao cac Name Ranges
Private Sub CreateNameRanges()
    On Error GoTo ErrorHandler
    
    ' Tao Name Range cho Raw_DuNo
    ThisWorkbook.Names.Add Name:="tblDuNo", _
        RefersTo:="=" & SHEET_RAW_DU_NO & "!$A$1:$X$1000"
    
    ' Tao Name Range cho Raw_TaiSan
    ThisWorkbook.Names.Add Name:="tblTaiSan", _
        RefersTo:="=" & SHEET_RAW_TAI_SAN & "!$A$1:$T$1000"
    
    ' Tao Name Range cho Raw_TraGoc
    ThisWorkbook.Names.Add Name:="tblTraGoc", _
        RefersTo:="=" & SHEET_RAW_TRA_GOC & "!$A$1:$P$1000"
    
    ' Tao Name Range cho Raw_TraLai
    ThisWorkbook.Names.Add Name:="tblTraLai", _
        RefersTo:="=" & SHEET_RAW_TRA_LAI & "!$A$1:$P$1000"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi tao Name Ranges: " & Err.description, vbCritical, "Loi"
End Sub

' Ham dam bao sheet Main luon ton tai va visible - NEW FUNCTION
Private Sub EnsureMainSheetVisible()
    On Error Resume Next
    
    ' Kiem tra sheet Main co ton tai khong
    If Not sheetExists(SHEET_MAIN) Then
        ' Tao sheet Main moi
        Dim newSheet As Worksheet
        Set newSheet = ThisWorkbook.sheets.Add(After:=ThisWorkbook.sheets(1))
        newSheet.Name = SHEET_MAIN
        
        ' Them noi dung co ban cho sheet Main
        With newSheet
            .Cells(1, 1).Value = "HE THONG QUAN LY THONG TIN KHACH HANG VAY"
            .Cells(1, 1).Font.Size = 16
            .Cells(1, 1).Font.Bold = True
            .Cells(3, 1).Value = "Dashboard Chinh - Giai doan 1"
            .Cells(5, 1).Value = "Cac chuc nang dang phat trien..."
        End With
    End If
    
    ' Dam bao sheet Main visible
    ThisWorkbook.sheets(SHEET_MAIN).Visible = xlSheetVisible
    On Error GoTo 0
End Sub
