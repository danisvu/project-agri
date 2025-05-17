Attribute VB_Name = "modDataStructureValidator"
'Attribute VB_Name = "modDataStructureValidator"
' Module kiem tra cau truc du lieu
' Kiem tra tinh toan ven va kha nang khoi phuc cau truc du lieu
' Tac gia: Phong Khach hang Ca nhan, Agribank Chi nhanh 4
' Ngay tao: 11/05/2025

Option Explicit

' Kiem tra chi tiet cau truc cua mot sheet cu the
Public Function ValidateSheetStructure(ByVal sheetName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Kiem tra sheet ton tai
    If Not modUtility.sheetExists(sheetName) Then
        Debug.Print "Sheet " & sheetName & " khong ton tai"
        ValidateSheetStructure = False
        Exit Function
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(sheetName)
    
    ' Kiem tra tung loai sheet
    Select Case sheetName
        Case SHEET_RAW_DU_NO
            ValidateSheetStructure = ValidateDuNoStructure(ws)
        
        Case SHEET_RAW_TAI_SAN
            ValidateSheetStructure = ValidateTaiSanStructure(ws)
        
        Case SHEET_RAW_TRA_GOC
            ValidateSheetStructure = ValidateTraGocStructure(ws)
        
        Case SHEET_RAW_TRA_LAI
            ValidateSheetStructure = ValidateTraLaiStructure(ws)
        
        Case SHEET_IMPORT_LOG
            ValidateSheetStructure = ValidateImportLogStructure(ws)
        
        Case SHEET_STAFF_ASSIGNMENT
            ValidateSheetStructure = ValidateStaffAssignmentStructure(ws)
        
        Case SHEET_PROCESSED_DATA
            ValidateSheetStructure = ValidateProcessedDataStructure(ws)
        
        Case SHEET_TRANSACTION_DATA
            ValidateSheetStructure = ValidateTransactionDataStructure(ws)
        
        Case SHEET_CONFIG
            ValidateSheetStructure = ValidateConfigStructure(ws)
        
        Case SHEET_USERS
            ValidateSheetStructure = ValidateUsersStructure(ws)
        
        Case Else
            ' Sheet khong can kiem tra
            ValidateSheetStructure = True
    End Select
    
    Exit Function
    
ErrorHandler:
    LogError "ValidateSheetStructure", Err.Number, Err.description
    ValidateSheetStructure = False
End Function

' Kiem tra cau truc sheet DuNo
Private Function ValidateDuNoStructure(ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    ValidateDuNoStructure = False ' Mac dinh la that bai
    
    ' Kiem tra cac cot tieu de chu yeu
    If ws.Cells(1, 1).Value <> COL_DU_NO_MA_KHOAN_VAY Then GoTo StructureInvalid
    If ws.Cells(1, 2).Value <> COL_DU_NO_MA_KHACH_HANG Then GoTo StructureInvalid
    If ws.Cells(1, 3).Value <> COL_DU_NO_TEN_KHACH_HANG Then GoTo StructureInvalid
    
    ' Kiem tra dinh dang header
    If Not ws.Cells(1, 1).Font.Bold Then GoTo StructureInvalid
    
    ' Neu qua tat ca kiem tra
    ValidateDuNoStructure = True
    Exit Function
    
StructureInvalid:
    Debug.Print "Cau truc sheet " & SHEET_RAW_DU_NO & " khong hop le"
    ValidateDuNoStructure = False
    Exit Function
    
ErrorHandler:
    Debug.Print "Loi khi kiem tra cau truc " & SHEET_RAW_DU_NO & ": " & Err.Number & " - " & Err.description
    ValidateDuNoStructure = False
End Function

' Kiem tra cau truc sheet TaiSan
Private Function ValidateTaiSanStructure(ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    ValidateTaiSanStructure = False ' Mac dinh la that bai
    
    ' Kiem tra cac cot tieu de chu yeu
    If ws.Cells(1, 1).Value <> COL_TAI_SAN_MA_TAI_SAN Then GoTo StructureInvalid
    If ws.Cells(1, 2).Value <> COL_TAI_SAN_MA_KHACH_HANG Then GoTo StructureInvalid
    If ws.Cells(1, 3).Value <> COL_TAI_SAN_TEN_KHACH_HANG Then GoTo StructureInvalid
    
    ' Kiem tra dinh dang header
    If Not ws.Cells(1, 1).Font.Bold Then GoTo StructureInvalid
    
    ' Neu qua tat ca kiem tra
    ValidateTaiSanStructure = True
    Exit Function
    
StructureInvalid:
    Debug.Print "Cau truc sheet " & SHEET_RAW_TAI_SAN & " khong hop le"
    ValidateTaiSanStructure = False
    Exit Function
    
ErrorHandler:
    Debug.Print "Loi khi kiem tra cau truc " & SHEET_RAW_TAI_SAN & ": " & Err.Number & " - " & Err.description
    ValidateTaiSanStructure = False
End Function

' Kiem tra cau truc sheet TraGoc
Private Function ValidateTraGocStructure(ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    ValidateTraGocStructure = False ' Mac dinh la that bai
    
    ' Kiem tra cac cot tieu de chu yeu
    If ws.Cells(1, 1).Value <> COL_TRA_GOC_MA_LICH_TRA_GOC Then GoTo StructureInvalid
    If ws.Cells(1, 2).Value <> COL_TRA_GOC_MA_KHACH_HANG Then GoTo StructureInvalid
    If ws.Cells(1, 3).Value <> COL_TRA_GOC_TEN_KHACH_HANG Then GoTo StructureInvalid
    
    ' Kiem tra dinh dang header
    If Not ws.Cells(1, 1).Font.Bold Then GoTo StructureInvalid
    
    ' Neu qua tat ca kiem tra
    ValidateTraGocStructure = True
    Exit Function
    
StructureInvalid:
    Debug.Print "Cau truc sheet " & SHEET_RAW_TRA_GOC & " khong hop le"
    ValidateTraGocStructure = False
    Exit Function
    
ErrorHandler:
    Debug.Print "Loi khi kiem tra cau truc " & SHEET_RAW_TRA_GOC & ": " & Err.Number & " - " & Err.description
    ValidateTraGocStructure = False
End Function

' Kiem tra cau truc sheet TraLai
Private Function ValidateTraLaiStructure(ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    ValidateTraLaiStructure = False ' Mac dinh la that bai
    
    ' Kiem tra cac cot tieu de chu yeu
    If ws.Cells(1, 1).Value <> COL_TRA_LAI_MA_LICH_TRA_LAI Then GoTo StructureInvalid
    If ws.Cells(1, 2).Value <> COL_TRA_LAI_MA_KHACH_HANG Then GoTo StructureInvalid
    If ws.Cells(1, 3).Value <> COL_TRA_LAI_TEN_KHACH_HANG Then GoTo StructureInvalid
    
    ' Kiem tra dinh dang header
    If Not ws.Cells(1, 1).Font.Bold Then GoTo StructureInvalid
    
    ' Neu qua tat ca kiem tra
    ValidateTraLaiStructure = True
    Exit Function
    
StructureInvalid:
    Debug.Print "Cau truc sheet " & SHEET_RAW_TRA_LAI & " khong hop le"
    ValidateTraLaiStructure = False
    Exit Function
    
ErrorHandler:
    Debug.Print "Loi khi kiem tra cau truc " & SHEET_RAW_TRA_LAI & ": " & Err.Number & " - " & Err.description
    ValidateTraLaiStructure = False
End Function

' Kiem tra cau truc sheet ImportLog
Private Function ValidateImportLogStructure(ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    ValidateImportLogStructure = False ' Mac dinh la that bai
    
    ' Kiem tra cac cot tieu de chu yeu
    If ws.Cells(1, 1).Value <> COL_IMPORT_LOG_ID Then GoTo StructureInvalid
    If ws.Cells(1, 2).Value <> COL_IMPORT_LOG_TEN_FILE Then GoTo StructureInvalid
    If ws.Cells(1, 3).Value <> COL_IMPORT_LOG_LOAI_DU_LIEU Then GoTo StructureInvalid
    
    ' Kiem tra dinh dang header
    If Not ws.Cells(1, 1).Font.Bold Then GoTo StructureInvalid
    
    ' Neu qua tat ca kiem tra
    ValidateImportLogStructure = True
    Exit Function
    
StructureInvalid:
    Debug.Print "Cau truc sheet " & SHEET_IMPORT_LOG & " khong hop le"
    ValidateImportLogStructure = False
    Exit Function
    
ErrorHandler:
    Debug.Print "Loi khi kiem tra cau truc " & SHEET_IMPORT_LOG & ": " & Err.Number & " - " & Err.description
    ValidateImportLogStructure = False
End Function

' Kiem tra cau truc sheet StaffAssignment
Private Function ValidateStaffAssignmentStructure(ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    ValidateStaffAssignmentStructure = False ' Mac dinh la that bai
    
    ' Kiem tra cac cot tieu de chu yeu
    If ws.Cells(1, 1).Value <> COL_STAFF_ASSIGNMENT_MA_KHACH_HANG Then GoTo StructureInvalid
    If ws.Cells(1, 2).Value <> COL_STAFF_ASSIGNMENT_MA_CAN_BO Then GoTo StructureInvalid
    
    ' Kiem tra dinh dang header
    If Not ws.Cells(1, 1).Font.Bold Then GoTo StructureInvalid
    
    ' Neu qua tat ca kiem tra
    ValidateStaffAssignmentStructure = True
    Exit Function
    
StructureInvalid:
    Debug.Print "Cau truc sheet " & SHEET_STAFF_ASSIGNMENT & " khong hop le"
    ValidateStaffAssignmentStructure = False
    Exit Function
    
ErrorHandler:
    Debug.Print "Loi khi kiem tra cau truc " & SHEET_STAFF_ASSIGNMENT & ": " & Err.Number & " - " & Err.description
    ValidateStaffAssignmentStructure = False
End Function

' Kiem tra cau truc sheet Processed_Data
Private Function ValidateProcessedDataStructure(ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    ValidateProcessedDataStructure = False ' Mac dinh la that bai
    
    ' Kiem tra cac cot tieu de chu yeu
    If ws.Cells(1, 1).Value <> "MaKhachHang" Then GoTo StructureInvalid
    If ws.Cells(1, 2).Value <> "TenKhachHang" Then GoTo StructureInvalid
    
    ' Kiem tra dinh dang header
    If Not ws.Cells(1, 1).Font.Bold Then GoTo StructureInvalid
    
    ' Neu qua tat ca kiem tra
    ValidateProcessedDataStructure = True
    Exit Function
    
StructureInvalid:
    Debug.Print "Cau truc sheet " & SHEET_PROCESSED_DATA & " khong hop le"
    ValidateProcessedDataStructure = False
    Exit Function
    
ErrorHandler:
    Debug.Print "Loi khi kiem tra cau truc " & SHEET_PROCESSED_DATA & ": " & Err.Number & " - " & Err.description
    ValidateProcessedDataStructure = False
End Function

' Kiem tra cau truc sheet TransactionData
Private Function ValidateTransactionDataStructure(ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    ValidateTransactionDataStructure = False ' Mac dinh la that bai
    
    ' Kiem tra cac cot tieu de chu yeu
    If ws.Cells(1, 1).Value <> "MaGiaoDich" Then GoTo StructureInvalid
    If ws.Cells(1, 2).Value <> "MaKhachHang" Then GoTo StructureInvalid
    
    ' Kiem tra dinh dang header
    If Not ws.Cells(1, 1).Font.Bold Then GoTo StructureInvalid
    
    ' Neu qua tat ca kiem tra
    ValidateTransactionDataStructure = True
    Exit Function
    
StructureInvalid:
    Debug.Print "Cau truc sheet " & SHEET_TRANSACTION_DATA & " khong hop le"
    ValidateTransactionDataStructure = False
    Exit Function
    
ErrorHandler:
    Debug.Print "Loi khi kiem tra cau truc " & SHEET_TRANSACTION_DATA & ": " & Err.Number & " - " & Err.description
    ValidateTransactionDataStructure = False
End Function

' Kiem tra cau truc sheet Config
Private Function ValidateConfigStructure(ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    ValidateConfigStructure = False ' Mac dinh la that bai
    
    ' Kiem tra cac cot tieu de chu yeu
    If ws.Cells(1, 1).Value <> "TenCauHinh" Then GoTo StructureInvalid
    If ws.Cells(1, 2).Value <> "GiaTri" Then GoTo StructureInvalid
    
    ' Kiem tra dinh dang header
    If Not ws.Cells(1, 1).Font.Bold Then GoTo StructureInvalid
    
    ' Neu qua tat ca kiem tra
    ValidateConfigStructure = True
    Exit Function
    
StructureInvalid:
    Debug.Print "Cau truc sheet " & SHEET_CONFIG & " khong hop le"
    ValidateConfigStructure = False
    Exit Function
    
ErrorHandler:
    Debug.Print "Loi khi kiem tra cau truc " & SHEET_CONFIG & ": " & Err.Number & " - " & Err.description
    ValidateConfigStructure = False
End Function

' Kiem tra cau truc sheet Users
Private Function ValidateUsersStructure(ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    ValidateUsersStructure = False ' Mac dinh la that bai
    
    ' Kiem tra cac cot tieu de chu yeu
    If ws.Cells(1, 1).Value <> "ID" Then GoTo StructureInvalid
    If ws.Cells(1, 2).Value <> "TenDangNhap" Then GoTo StructureInvalid
    
    ' Kiem tra dinh dang header
    If Not ws.Cells(1, 1).Font.Bold Then GoTo StructureInvalid
    
    ' Neu qua tat ca kiem tra
    ValidateUsersStructure = True
    Exit Function
    
StructureInvalid:
    Debug.Print "Cau truc sheet " & SHEET_USERS & " khong hop le"
    ValidateUsersStructure = False
    Exit Function
    
ErrorHandler:
    Debug.Print "Loi khi kiem tra cau truc " & SHEET_USERS & ": " & Err.Number & " - " & Err.description
    ValidateUsersStructure = False
End Function

' Kiem tra va khoi phuc cau truc du lieu chi tiet
Public Sub ValidateAndRepairDataStructure()
    On Error GoTo ErrorHandler
    
    ' Toi uu hoa hieu suat
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    
    ' Danh sach cac sheet can kiem tra
    Dim requiredSheets As Variant
    requiredSheets = GetRequiredDataSheets()
    
    Dim configSheets As Variant
    configSheets = GetRequiredConfigSheets()
    
    Dim i As Integer
    Dim invalidStructure As Boolean
    Dim missingSheets As String
    
    invalidStructure = False
    missingSheets = ""
    
    ' Kiem tra cac sheet du lieu
    For i = LBound(requiredSheets) To UBound(requiredSheets)
        If Not modUtility.sheetExists(CStr(requiredSheets(i))) Then
            ' Sheet khong ton tai, can khoi tao moi
            missingSheets = missingSheets & CStr(requiredSheets(i)) & ", "
            invalidStructure = True
        Else
            ' Sheet ton tai, kiem tra cau truc chi tiet
            If Not ValidateSheetStructure(CStr(requiredSheets(i))) Then
                invalidStructure = True
                Exit For
            End If
        End If
    Next i
    
    ' Kiem tra cac sheet cau hinh
    If Not invalidStructure Then
        For i = LBound(configSheets) To UBound(configSheets)
            If Not modUtility.sheetExists(CStr(configSheets(i))) Then
                ' Sheet khong ton tai, can khoi tao moi
                missingSheets = missingSheets & CStr(configSheets(i)) & ", "
                invalidStructure = True
            Else
                ' Sheet ton tai, kiem tra cau truc chi tiet
                If Not ValidateSheetStructure(CStr(configSheets(i))) Then
                    invalidStructure = True
                    Exit For
                End If
            End If
        Next i
    End If
    
    ' Neu phat hien cau truc khong hop le
    If invalidStructure Then
        ' Sao luu du lieu truoc khi khoi phuc
        If modBackupRestore.BackupBeforeRepair() Then
            Dim response As VbMsgBoxResult
            
            If missingSheets <> "" Then
                missingSheets = Left(missingSheets, Len(missingSheets) - 2) ' Loai bo dau phay cuoi cung
                response = MsgBox("Phat hien cac sheet bi thieu: " & missingSheets & vbCrLf & _
                               "He thong se sao luu du lieu hien tai va khoi phuc cau truc chuan." & vbCrLf & _
                               "Ban co muon tiep tuc?", vbYesNo + vbQuestion, "Khoi phuc cau truc du lieu")
            Else
                response = MsgBox("Phat hien cau truc du lieu khong hop le." & vbCrLf & _
                               "He thong se sao luu du lieu hien tai va khoi phuc cau truc chuan." & vbCrLf & _
                               "Ban co muon tiep tuc?", vbYesNo + vbQuestion, "Khoi phuc cau truc du lieu")
            End If
            
            If response = vbYes Then
                ' Khoi phuc cau truc du lieu
                modDataStructure.InitializeDataStructure
                
                ' Thong bao khoi phuc thanh cong
                MsgBox "Cau truc du lieu da duoc khoi phuc thanh cong." & vbCrLf & _
                      "Du lieu goc da duoc sao luu truoc khi khoi phuc.", vbInformation, "Khoi phuc thanh cong"
            End If
        Else
            MsgBox "Khong the sao luu du lieu hien tai." & vbCrLf & _
                  "Qua trinh khoi phuc cau truc se khong duoc thuc hien.", vbExclamation, "Loi sao luu"
        End If
    End If
    
CleanUp:
    ' Khoi phuc cac cai dat
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Exit Sub
    
ErrorHandler:
    LogError "ValidateAndRepairDataStructure", Err.Number, Err.description
    Resume CleanUp
End Sub

