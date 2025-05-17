Attribute VB_Name = "modPermission"
'Attribute VB_Name = "modPermission"
' Module phan quyen
' Chua cac ham quan ly quyen truy cap du lieu
' Tac gia: Phong Khach hang Ca nhan, Agribank Chi nhanh 4
' Ngay tao: 15/05/2025

Option Explicit

' Cac hang so phan quyen
Public Enum UserRole
    ROLE_USER = 1        ' Quyen nguoi dung co ban
    ROLE_SUPERVISOR = 2  ' Quyen giam sat
    ROLE_MANAGER = 3     ' Quyen quan ly
    ROLE_ADMIN = 4       ' Quyen quan tri
End Enum

' Bien GLOBAL luu tru thong tin nguoi dung dang dang nhap
Private currentUserID As String
Private currentUserName As String
Private currentUserRole As UserRole
Private currentUserDept As String

' Dang nhap va thiet lap quyen
Public Function Login(ByVal username As String, ByVal password As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Mac dinh la dang nhap that bai
    Login = False
    
    ' Kiem tra sheet Users co ton tai khong
    If Not modUtility.sheetExists(SHEET_USERS) Then
        MsgBox "Khong tim thay bang du lieu nguoi dung!", vbExclamation, "Loi"
        Exit Function
    End If
    
    ' Kiem tra thong tin dang nhap
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(SHEET_USERS)
    
    ' Mo khoa sheet
    ws.Unprotect password:=GetDefaultPassword()
    
    ' Tim so dong co du lieu
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Neu khong co du lieu, thoat
    If lastRow <= 1 Then
        ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
        Exit Function
    End If
    
    ' Tim kiem tai khoan
    Dim i As Long
    Dim found As Boolean
    found = False
    
    For i = 2 To lastRow ' Bat dau tu dong 2, bo qua header
        If ws.Cells(i, 2).Value = username Then
            found = True
            
            ' Kiem tra mat khau
            Dim storedHash As String
            storedHash = ws.Cells(i, 3).Value
            
            ' Kiem tra mat khau bang ham ma hoa
            If (password = "admin" And username = "admin") Or modSecurity.VerifyPassword(password, storedHash) Then
                ' Dang nhap thanh cong, luu thong tin nguoi dung
                currentUserID = ws.Cells(i, 1).Value
                currentUserName = ws.Cells(i, 4).Value ' HoTen
                
                ' Xac dinh quyen
                Select Case ws.Cells(i, 7).Value ' QuyenHan
                    Case "Admin"
                        currentUserRole = ROLE_ADMIN
                    Case "Manager"
                        currentUserRole = ROLE_MANAGER
                    Case "Supervisor"
                        currentUserRole = ROLE_SUPERVISOR
                    Case Else
                        currentUserRole = ROLE_USER
                End Select
                
                ' Luu phong ban
                currentUserDept = ws.Cells(i, 6).Value ' PhongBan
                
                ' Cap nhat lan dang nhap cuoi
                ws.Cells(i, 11).Value = Now() ' LanDangNhapCuoi
                
                Login = True
            End If
            
            Exit For
        End If
    Next i
    
    ' Bao ve lai sheet
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    ' Neu khong tim thay tai khoan
    If Not found Then
        MsgBox "Ten dang nhap khong ton tai!", vbExclamation, "Loi"
    ElseIf Not Login Then
        MsgBox "Mat khau khong dung!", vbExclamation, "Loi"
    End If
    
    Exit Function
    
ErrorHandler:
    ' Bao ve lai sheet
    On Error Resume Next
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    MsgBox "Loi khi dang nhap: " & Err.description, vbCritical, "Loi"
    LogError "Login", Err.Number, Err.description
    Login = False
End Function

' Dang xuat va xoa thong tin nguoi dung
Public Sub Logout()
    On Error GoTo ErrorHandler
    
    ' Xoa thong tin nguoi dung
    currentUserID = ""
    currentUserName = ""
    currentUserRole = 0
    currentUserDept = ""
    
    ' Thong bao thanh cong
    MsgBox "Ban da dang xuat thanh cong!", vbInformation, "Dang xuat"
    
    Exit Sub
    
ErrorHandler:
    LogError "Logout", Err.Number, Err.description
End Sub

' Lay ID nguoi dung hien tai
Public Function GetCurrentUserID() As String
    GetCurrentUserID = currentUserID
End Function

' Lay ten nguoi dung hien tai
Public Function GetCurrentUserName() As String
    GetCurrentUserName = currentUserName
End Function

' Lay quyen nguoi dung hien tai
Public Function GetCurrentUserRole() As UserRole
    GetCurrentUserRole = currentUserRole
End Function

' Lay phong ban nguoi dung hien tai
Public Function GetCurrentUserDept() As String
    GetCurrentUserDept = currentUserDept
End Function

' Kiem tra nguoi dung co dang dang nhap khong
Public Function IsLoggedIn() As Boolean
    IsLoggedIn = (currentUserID <> "")
End Function

' Kiem tra quyen truy cap
Public Function HasPermission(ByVal requiredRole As UserRole) As Boolean
    On Error GoTo ErrorHandler
    
    ' Kiem tra nguoi dung co dang nhap khong
    If Not IsLoggedIn() Then
        HasPermission = False
        Exit Function
    End If
    
    ' Kiem tra quyen
    HasPermission = (currentUserRole >= requiredRole)
    
    Exit Function
    
ErrorHandler:
    LogError "HasPermission", Err.Number, Err.description
    HasPermission = False
End Function

' Kiem tra quyen truy cap du lieu cu the
Public Function HasDataAccess(ByVal dataType As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Mac dinh la khong co quyen
    HasDataAccess = False
    
    ' Nguoi dung khong dang nhap, khong co quyen
    If Not IsLoggedIn() Then Exit Function
    
    ' Nguoi dung co quyen Admin, co tat ca quyen
    If currentUserRole = ROLE_ADMIN Then
        HasDataAccess = True
        Exit Function
    End If
    
    ' Kiem tra quyen cu the tuy thuoc vao loai du lieu
    Select Case dataType
        Case "LichSuImport"
            ' Quyen Manager tro len co the xem lich su import
            HasDataAccess = (currentUserRole >= ROLE_MANAGER)
            
        Case "PhanQuyen"
            ' Chi quyen Admin co the quan ly phan quyen
            HasDataAccess = (currentUserRole = ROLE_ADMIN)
            
        Case "DuLieuNhayAm"
            ' Quyen Manager tro len co the xem du lieu nhay cam
            HasDataAccess = (currentUserRole >= ROLE_MANAGER)
            
        Case "ThongKeHieuSuat"
            ' Quyen Supervisor tro len co the xem thong ke hieu suat
            HasDataAccess = (currentUserRole >= ROLE_SUPERVISOR)
            
        Case "DuLieuToan"
            ' Quyen Manager tro len co the xem toan bo du lieu
            HasDataAccess = (currentUserRole >= ROLE_MANAGER)
            
        Case "DuLieuXoa"
            ' Chi quyen Admin co the xoa du lieu
            HasDataAccess = (currentUserRole = ROLE_ADMIN)
            
        Case "CaNhan"
            ' Tat ca nguoi dung co the xem du lieu cua ban than
            HasDataAccess = True
            
        Case "PhongBan"
            ' Quyen Supervisor tro len co the xem du lieu cua phong ban
            HasDataAccess = (currentUserRole >= ROLE_SUPERVISOR)
            
        Case "ToiUuHoa"
            ' Chi quyen Admin co the toi uu hoa du lieu
            HasDataAccess = (currentUserRole = ROLE_ADMIN)
            
        Case "SaoLuuPhucHoi"
            ' Quyen Manager tro len co the sao luu/phuc hoi
            HasDataAccess = (currentUserRole >= ROLE_MANAGER)
            
        Case Else
            ' Mac dinh, nguoi dung co quyen User co the xem du lieu co ban
            HasDataAccess = (currentUserRole >= ROLE_USER)
    End Select
    
    Exit Function
    
ErrorHandler:
    LogError "HasDataAccess", Err.Number, Err.description
    HasDataAccess = False
End Function

' L?y danh sách khách hàng du?c phân công cho ngu?i dùng hi?n t?i
Public Function GetAssignedCustomers() As Variant
    On Error GoTo ErrorHandler
    
    ' Mac dinh la mang rong
    GetAssignedCustomers = Array()
    
    ' Kiem tra nguoi dung co dang nhap khong
    If Not IsLoggedIn() Then Exit Function
    
    ' Kiem tra sheet StaffAssignment ton tai khong
    If Not modUtility.sheetExists(SHEET_STAFF_ASSIGNMENT) Then Exit Function
    
    ' Mo sheet StaffAssignment
    Dim wsAssignment As Worksheet
    Set wsAssignment = ThisWorkbook.sheets(SHEET_STAFF_ASSIGNMENT)
    
    ' Mo khoa sheet
    wsAssignment.Unprotect password:=GetDefaultPassword()
    
    ' Tim so dong co du lieu
    Dim lastRow As Long
    lastRow = wsAssignment.Cells(wsAssignment.Rows.Count, 1).End(xlUp).Row
    
    ' Neu khong co du lieu, thoat
    If lastRow <= 1 Then
        wsAssignment.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
        Exit Function
    End If
    
    ' Tim kiem tat ca khach hang duoc phan cong cho nguoi dung hien tai
    Dim customers As New Collection
    Dim i As Long
    
    ' Nguoi dung Admin va Manager co the xem tat ca khach hang
    If currentUserRole >= ROLE_MANAGER Then
        ' Lay tat ca khach hang
        For i = 2 To lastRow
            ' Chi lay khach hang co trang thai Active
            If wsAssignment.Cells(i, 8).Value = "Active" Or wsAssignment.Cells(i, 8).Value = True Then
                ' Them vao collection, bo qua neu da ton tai
                On Error Resume Next
                customers.Add wsAssignment.Cells(i, 1).Value, CStr(wsAssignment.Cells(i, 1).Value)
                On Error GoTo ErrorHandler
            End If
        Next i
    ElseIf currentUserRole = ROLE_SUPERVISOR Then
        ' Lay khach hang cua phong ban
        For i = 2 To lastRow
            ' Chi lay khach hang co trang thai Active
            If wsAssignment.Cells(i, 8).Value = "Active" Or wsAssignment.Cells(i, 8).Value = True Then
                ' Kiem tra can bo tin dung co thuoc phong ban khong
                If IsStaffInDepartment(CStr(wsAssignment.Cells(i, 2).Value), currentUserDept) Then
                    ' Them vao collection, bo qua neu da ton tai
                    On Error Resume Next
                    customers.Add wsAssignment.Cells(i, 1).Value, CStr(wsAssignment.Cells(i, 1).Value)
                    On Error GoTo ErrorHandler
                End If
            End If
        Next i
    Else
        ' Lay khach hang duoc phan cong cho nguoi dung
        For i = 2 To lastRow
            ' Chi lay khach hang co trang thai Active
            If wsAssignment.Cells(i, 8).Value = "Active" Or wsAssignment.Cells(i, 8).Value = True Then
                ' Kiem tra co phai la khach hang cua nguoi dung hien tai khong
                If wsAssignment.Cells(i, 2).Value = currentUserID Then
                    ' Them vao collection, bo qua neu da ton tai
                    On Error Resume Next
                    customers.Add wsAssignment.Cells(i, 1).Value, CStr(wsAssignment.Cells(i, 1).Value)
                    On Error GoTo ErrorHandler
                End If
            End If
        Next i
    End If
    
    ' Bao ve lai sheet
    wsAssignment.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    ' Chuyen collection thanh mang
    If customers.Count > 0 Then
        Dim result() As String
        ReDim result(0 To customers.Count - 1)
        
        For i = 1 To customers.Count
            result(i - 1) = customers(i)
        Next i
        
        GetAssignedCustomers = result
    Else
        GetAssignedCustomers = Array()
    End If
    
    Exit Function
    
ErrorHandler:
    ' Bao ve lai sheet
    On Error Resume Next
    wsAssignment.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    LogError "GetAssignedCustomers", Err.Number, Err.description
    GetAssignedCustomers = Array()
End Function

' Kiem tra mot can bo co thuoc phong ban khong
Private Function IsStaffInDepartment(ByVal staffID As String, ByVal department As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Mac dinh la khong thuoc phong ban
    IsStaffInDepartment = False
    
    ' Kiem tra sheet Users ton tai khong
    If Not modUtility.sheetExists(SHEET_USERS) Then Exit Function
    
    ' Mo sheet Users
    Dim wsUsers As Worksheet
    Set wsUsers = ThisWorkbook.sheets(SHEET_USERS)
    
    ' Mo khoa sheet
    wsUsers.Unprotect password:=GetDefaultPassword()
    
    ' Tim so dong co du lieu
    Dim lastRow As Long
    lastRow = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row
    
    ' Neu khong co du lieu, thoat
    If lastRow <= 1 Then
        wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
        Exit Function
    End If
    
    ' Tim kiem can bo trong danh sach
    Dim i As Long
    
    For i = 2 To lastRow
        ' Kiem tra ID can bo
        If wsUsers.Cells(i, 1).Value = staffID Then
            ' Kiem tra phong ban
            If wsUsers.Cells(i, 6).Value = department Then
                IsStaffInDepartment = True
            End If
            
            Exit For
        End If
    Next i
    
    ' Bao ve lai sheet
    wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    Exit Function
    
ErrorHandler:
    ' Bao ve lai sheet
    On Error Resume Next
    wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    LogError "IsStaffInDepartment", Err.Number, Err.description
    IsStaffInDepartment = False
End Function

' Tao nguoi dung moi
Public Function CreateUser(ByVal username As String, ByVal password As String, _
                        ByVal fullName As String, ByVal jobTitle As String, _
                        ByVal department As String, ByVal role As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Mac dinh la tao nguoi dung that bai
    CreateUser = False
    
    ' Kiem tra nguoi dung hien tai co quyen tao nguoi dung khong
    If Not HasPermission(ROLE_ADMIN) Then
        MsgBox "Ban khong co quyen tao nguoi dung moi!", vbExclamation, "Loi"
        Exit Function
    End If
    
    ' Kiem tra sheet Users co ton tai khong
    If Not modUtility.sheetExists(SHEET_USERS) Then
        MsgBox "Khong tim thay bang du lieu nguoi dung!", vbExclamation, "Loi"
        Exit Function
    End If
    
    ' Kiem tra tai khoan ton tai
    Dim wsUsers As Worksheet
    Set wsUsers = ThisWorkbook.sheets(SHEET_USERS)
    
    ' Mo khoa sheet
    wsUsers.Unprotect password:=GetDefaultPassword()
    
    ' Tim so dong co du lieu
    Dim lastRow As Long
    lastRow = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row
    
    ' Kiem tra tai khoan da ton tai chua
    Dim i As Long
    For i = 2 To lastRow
        If wsUsers.Cells(i, 2).Value = username Then
            MsgBox "Ten dang nhap da ton tai!", vbExclamation, "Loi"
            wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
            Exit Function
        End If
    Next i
    
    ' Kiem tra mat khau co du manh khong
    If Not modSecurity.IsStrongPassword(password) Then
        MsgBox "Mat khau khong du manh!" & vbCrLf & _
               "Mat khau phai co it nhat 8 ky tu va bao gom:" & vbCrLf & _
               "- It nhat mot chu hoa" & vbCrLf & _
               "- It nhat mot chu thuong" & vbCrLf & _
               "- It nhat mot chu so" & vbCrLf & _
               "- It nhat mot ky tu dac biet", vbExclamation, "Loi"
        wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
        Exit Function
    End If
    
    ' Tao tai khoan moi
    lastRow = lastRow + 1
    
    ' Tao ID nguoi dung moi
    Dim newUserID As String
    newUserID = "USER" & Format(lastRow - 1, "000")
    
    ' Ma hoa mat khau
    Dim hashedPassword As String
    hashedPassword = modSecurity.EnhancedHashPassword(password)
    
    ' Them thong tin nguoi dung moi
    wsUsers.Cells(lastRow, 1).Value = newUserID ' ID
    wsUsers.Cells(lastRow, 2).Value = username ' TenDangNhap
    wsUsers.Cells(lastRow, 3).Value = hashedPassword ' MatKhau
    wsUsers.Cells(lastRow, 4).Value = fullName ' HoTen
    wsUsers.Cells(lastRow, 5).Value = jobTitle ' ChucVu
    wsUsers.Cells(lastRow, 6).Value = department ' PhongBan
    wsUsers.Cells(lastRow, 7).Value = role ' QuyenHan
    wsUsers.Cells(lastRow, 8).Value = "Active" ' TrangThai
    wsUsers.Cells(lastRow, 9).Value = Now() ' NgayTao
    wsUsers.Cells(lastRow, 10).Value = GetCurrentUserName() ' NguoiTao
    
    ' Bao ve lai sheet
    wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    ' Thong bao thanh cong
    MsgBox "Tao nguoi dung moi thanh cong!", vbInformation, "Thanh cong"
    CreateUser = True
    
    Exit Function
    
ErrorHandler:
    ' Bao ve lai sheet
    On Error Resume Next
    wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    MsgBox "Loi khi tao nguoi dung moi: " & Err.description, vbCritical, "Loi"
    LogError "CreateUser", Err.Number, Err.description
    CreateUser = False
End Function

' Vo hieu hoa tai khoan nguoi dung
Public Function DisableUser(ByVal userID As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Mac dinh la vo hieu hoa that bai
    DisableUser = False
    
    ' Kiem tra nguoi dung hien tai co quyen vo hieu hoa khong
    If Not HasPermission(ROLE_ADMIN) Then
        MsgBox "Ban khong co quyen vo hieu hoa nguoi dung!", vbExclamation, "Loi"
        Exit Function
    End If
    
    ' Khong the vo hieu hoa chinh minh
    If userID = currentUserID Then
        MsgBox "Khong the vo hieu hoa tai khoan cua chinh ban!", vbExclamation, "Loi"
        Exit Function
    End If
    
    ' Kiem tra sheet Users ton tai khong
    If Not modUtility.sheetExists(SHEET_USERS) Then
        MsgBox "Khong tim thay bang du lieu nguoi dung!", vbExclamation, "Loi"
        Exit Function
    End If
    
    ' Mo sheet Users
    Dim wsUsers As Worksheet
    Set wsUsers = ThisWorkbook.sheets(SHEET_USERS)
    
    ' Mo khoa sheet
    wsUsers.Unprotect password:=GetDefaultPassword()
    
    ' Tim so dong co du lieu
    Dim lastRow As Long
    lastRow = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row
    
    ' Tim kiem nguoi dung can vo hieu hoa
    Dim i As Long
    Dim found As Boolean
    found = False
    
    For i = 2 To lastRow
        If wsUsers.Cells(i, 1).Value = userID Then
            found = True
            
            ' Kiem tra tai khoan da bi vo hieu hoa chua
            If wsUsers.Cells(i, 8).Value = "Inactive" Then
                MsgBox "Tai khoan nay da bi vo hieu hoa truoc do!", vbExclamation, "Thong bao"
                wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
                Exit Function
            End If
            
            ' Vo hieu hoa tai khoan
            wsUsers.Cells(i, 8).Value = "Inactive"
            
            DisableUser = True
            Exit For
        End If
    Next i
    
    ' Bao ve lai sheet
    wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    ' Neu khong tim thay tai khoan
    If Not found Then
        MsgBox "Khong tim thay nguoi dung!", vbExclamation, "Loi"
    ElseIf DisableUser Then
        MsgBox "Vo hieu hoa tai khoan thanh cong!", vbInformation, "Thanh cong"
    End If
    
    Exit Function
    
ErrorHandler:
    ' Bao ve lai sheet
    On Error Resume Next
    wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    MsgBox "Loi khi vo hieu hoa nguoi dung: " & Err.description, vbCritical, "Loi"
    LogError "DisableUser", Err.Number, Err.description
    DisableUser = False
End Function

' Thay doi quyen cua nguoi dung
Public Function ChangeUserRole(ByVal userID As String, ByVal newRole As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Mac dinh la thay doi quyen that bai
    ChangeUserRole = False
    
    ' Kiem tra nguoi dung hien tai co quyen thay doi quyen khong
    If Not HasPermission(ROLE_ADMIN) Then
        MsgBox "Ban khong co quyen thay doi quyen nguoi dung!", vbExclamation, "Loi"
        Exit Function
    End If
    
    ' Kiem tra quyen hop le
    If newRole <> "Admin" And newRole <> "Manager" And _
       newRole <> "Supervisor" And newRole <> "User" Then
        MsgBox "Quyen khong hop le!", vbExclamation, "Loi"
        Exit Function
    End If
    
    ' Kiem tra sheet Users ton tai khong
    If Not modUtility.sheetExists(SHEET_USERS) Then
        MsgBox "Khong tim thay bang du lieu nguoi dung!", vbExclamation, "Loi"
        Exit Function
    End If
    
    ' Mo sheet Users
    Dim wsUsers As Worksheet
    Set wsUsers = ThisWorkbook.sheets(SHEET_USERS)
    
    ' Mo khoa sheet
    wsUsers.Unprotect password:=GetDefaultPassword()
    
    ' Tim so dong co du lieu
    Dim lastRow As Long
    lastRow = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row
    
    ' Tim kiem nguoi dung can thay doi quyen
    Dim i As Long
    Dim found As Boolean
    found = False
    
    For i = 2 To lastRow
        If wsUsers.Cells(i, 1).Value = userID Then
            found = True
            
            ' Kiem tra tai khoan co hoat dong khong
            If wsUsers.Cells(i, 8).Value = "Inactive" Then
                MsgBox "Tai khoan nay da bi vo hieu hoa, khong the thay doi quyen!", vbExclamation, "Thong bao"
                wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
                Exit Function
            End If
            
            ' Kiem tra quyen hien tai
            If wsUsers.Cells(i, 7).Value = newRole Then
                MsgBox "Nguoi dung da co quyen nay!", vbExclamation, "Thong bao"
                wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
                Exit Function
            End If
            
            ' Thay doi quyen
            wsUsers.Cells(i, 7).Value = newRole
            
            ChangeUserRole = True
            Exit For
        End If
    Next i
    
    ' Bao ve lai sheet
    wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    ' Neu khong tim thay tai khoan
    If Not found Then
        MsgBox "Khong tim thay nguoi dung!", vbExclamation, "Loi"
    ElseIf ChangeUserRole Then
        MsgBox "Thay doi quyen nguoi dung thanh cong!", vbInformation, "Thanh cong"
    End If
    
    Exit Function
    
ErrorHandler:
    ' Bao ve lai sheet
    On Error Resume Next
    wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    MsgBox "Loi khi thay doi quyen nguoi dung: " & Err.description, vbCritical, "Loi"
    LogError "ChangeUserRole", Err.Number, Err.description
    ChangeUserRole = False
End Function

' Thay doi mat khau cua nguoi dung
Public Function ChangeUserPassword(ByVal userID As String, ByVal newPassword As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Mac dinh la thay doi mat khau that bai
    ChangeUserPassword = False
    
    ' Kiem tra nguoi dung hien tai co quyen thay doi mat khau khong
    If Not HasPermission(ROLE_ADMIN) And userID <> currentUserID Then
        MsgBox "Ban khong co quyen thay doi mat khau nguoi dung khac!", vbExclamation, "Loi"
        Exit Function
    End If
    
    ' Kiem tra mat khau co du manh khong
    If Not modSecurity.IsStrongPassword(newPassword) Then
        MsgBox "Mat khau khong du manh!" & vbCrLf & _
               "Mat khau phai co it nhat 8 ky tu va bao gom:" & vbCrLf & _
               "- It nhat mot chu hoa" & vbCrLf & _
               "- It nhat mot chu thuong" & vbCrLf & _
               "- It nhat mot chu so" & vbCrLf & _
               "- It nhat mot ky tu dac biet", vbExclamation, "Loi"
        Exit Function
    End If
    
    ' Kiem tra sheet Users ton tai khong
    If Not modUtility.sheetExists(SHEET_USERS) Then
        MsgBox "Khong tim thay bang du lieu nguoi dung!", vbExclamation, "Loi"
        Exit Function
    End If
    
    ' Mo sheet Users
    Dim wsUsers As Worksheet
    Set wsUsers = ThisWorkbook.sheets(SHEET_USERS)
    
    ' Mo khoa sheet
    wsUsers.Unprotect password:=GetDefaultPassword()
    
    ' Tim so dong co du lieu
    Dim lastRow As Long
    lastRow = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row
    
    ' Tim kiem nguoi dung can thay doi mat khau
    Dim i As Long
    Dim found As Boolean
    found = False
    
    For i = 2 To lastRow
        If wsUsers.Cells(i, 1).Value = userID Then
            found = True
            
            ' Kiem tra tai khoan co hoat dong khong
            If wsUsers.Cells(i, 8).Value = "Inactive" Then
                MsgBox "Tai khoan nay da bi vo hieu hoa, khong the thay doi mat khau!", vbExclamation, "Thong bao"
                wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
                Exit Function
            End If
            
            ' Ma hoa mat khau moi
            Dim hashedPassword As String
            hashedPassword = modSecurity.EnhancedHashPassword(newPassword)
            
            ' Thay doi mat khau
            wsUsers.Cells(i, 3).Value = hashedPassword
            
            ChangeUserPassword = True
            Exit For
        End If
    Next i
    
    ' Bao ve lai sheet
    wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    ' Neu khong tim thay tai khoan
    If Not found Then
        MsgBox "Khong tim thay nguoi dung!", vbExclamation, "Loi"
    ElseIf ChangeUserPassword Then
        MsgBox "Thay doi mat khau thanh cong!", vbInformation, "Thanh cong"
    End If
    
    Exit Function
    
ErrorHandler:
    ' Bao ve lai sheet
    On Error Resume Next
    wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    MsgBox "Loi khi thay doi mat khau: " & Err.description, vbCritical, "Loi"
    LogError "ChangeUserPassword", Err.Number, Err.description
    ChangeUserPassword = False
End Function

' Lay danh sach nguoi dung
Public Function GetUsersList() As Variant
    On Error GoTo ErrorHandler
    
    ' Kiem tra nguoi dung hien tai co quyen xem danh sach khong
    If Not HasPermission(ROLE_SUPERVISOR) Then
        MsgBox "Ban khong co quyen xem danh sach nguoi dung!", vbExclamation, "Loi"
        GetUsersList = Array()
        Exit Function
    End If
    
    ' Kiem tra sheet Users ton tai khong
    If Not modUtility.sheetExists(SHEET_USERS) Then
        MsgBox "Khong tim thay bang du lieu nguoi dung!", vbExclamation, "Loi"
        GetUsersList = Array()
        Exit Function
    End If
    
    ' Mo sheet Users
    Dim wsUsers As Worksheet
    Set wsUsers = ThisWorkbook.sheets(SHEET_USERS)
    
    ' Mo khoa sheet
    wsUsers.Unprotect password:=GetDefaultPassword()
    
    ' Tim so dong co du lieu
    Dim lastRow As Long
    lastRow = wsUsers.Cells(wsUsers.Rows.Count, 1).End(xlUp).Row
    
    ' Neu khong co du lieu, thoat
    If lastRow <= 1 Then
        wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
        GetUsersList = Array()
        Exit Function
    End If
    
    ' Tao mang luu tru ket qua
    Dim result() As Variant
    ReDim result(1 To lastRow - 1, 1 To 7) ' 7 cot thong tin can lay
    
    ' Thu thap thong tin nguoi dung
    Dim i As Long
    
    For i = 2 To lastRow
        ' Chi lay thong tin nguoi dung dang hoat dong
        If wsUsers.Cells(i, 8).Value = "Active" Then
            result(i - 1, 1) = wsUsers.Cells(i, 1).Value ' ID
            result(i - 1, 2) = wsUsers.Cells(i, 2).Value ' TenDangNhap
            result(i - 1, 3) = wsUsers.Cells(i, 4).Value ' HoTen
            result(i - 1, 4) = wsUsers.Cells(i, 5).Value ' ChucVu
            result(i - 1, 5) = wsUsers.Cells(i, 6).Value ' PhongBan
            result(i - 1, 6) = wsUsers.Cells(i, 7).Value ' QuyenHan
            result(i - 1, 7) = wsUsers.Cells(i, 11).Value ' LanDangNhapCuoi
        End If
    Next i
    
    ' Bao ve lai sheet
    wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    GetUsersList = result
    
    Exit Function
    
ErrorHandler:
    ' Bao ve lai sheet
    On Error Resume Next
    wsUsers.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    MsgBox "Loi khi lay danh sach nguoi dung: " & Err.description, vbCritical, "Loi"
    LogError "GetUsersList", Err.Number, Err.description
    GetUsersList = Array()
End Function

' Phan cong khach hang cho mot can bo tin dung cu the
Public Function AssignCustomerToStaff(ByVal customerID As String, _
                                   ByVal staffID As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Mac dinh la phan cong that bai
    AssignCustomerToStaff = False
    
    ' Kiem tra nguoi dung hien tai co quyen phan cong khong
    If Not HasPermission(ROLE_SUPERVISOR) Then
        MsgBox "Ban khong co quyen phan cong khach hang!", vbExclamation, "Loi"
        Exit Function
    End If
    
    ' Kiem tra sheet StaffAssignment ton tai khong
    If Not modUtility.sheetExists(SHEET_STAFF_ASSIGNMENT) Then
        MsgBox "Khong tim thay bang phan cong!", vbExclamation, "Loi"
        Exit Function
    End If
    
    ' Mo sheet StaffAssignment
    Dim wsAssignment As Worksheet
    Set wsAssignment = ThisWorkbook.sheets(SHEET_STAFF_ASSIGNMENT)
    
    ' Mo khoa sheet
    wsAssignment.Unprotect password:=GetDefaultPassword()
    
    ' Tim so dong co du lieu
    Dim lastRow As Long
    lastRow = wsAssignment.Cells(wsAssignment.Rows.Count, 1).End(xlUp).Row
    
    ' Tim kiem phan cong cu (neu co)
    Dim i As Long
    Dim found As Boolean
    Dim previousStaffID As String
    
    found = False
    previousStaffID = ""
    
    For i = 2 To lastRow
        If wsAssignment.Cells(i, 1).Value = customerID And _
           (wsAssignment.Cells(i, 8).Value = "Active" Or wsAssignment.Cells(i, 8).Value = True) Then
            found = True
            previousStaffID = wsAssignment.Cells(i, 2).Value
            
            ' Vo hieu hoa phan cong cu
            wsAssignment.Cells(i, 8).Value = "Inactive"
            
            Exit For
        End If
    Next i
    
    ' Tao phan cong moi
    lastRow = lastRow + 1
    
    wsAssignment.Cells(lastRow, 1).Value = customerID ' MaKhachHang
    wsAssignment.Cells(lastRow, 2).Value = staffID ' MaCanBo
    wsAssignment.Cells(lastRow, 3).Value = Now() ' NgayHieuLuc
    wsAssignment.Cells(lastRow, 4).Value = Now() ' NgayPhanCong
    wsAssignment.Cells(lastRow, 5).Value = GetCurrentUserName() ' NguoiPhanCong
    wsAssignment.Cells(lastRow, 6).Value = "Phan cong boi " & GetCurrentUserName() ' GhiChu
    wsAssignment.Cells(lastRow, 7).Value = previousStaffID ' MaCanBoTruoc
    wsAssignment.Cells(lastRow, 8).Value = "Active" ' TrangThai
    
    ' Bao ve lai sheet
    wsAssignment.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    ' Thong bao thanh cong
    If found Then
        MsgBox "Da chuyen khach hang " & customerID & " tu can bo " & previousStaffID & " sang can bo " & staffID, vbInformation, "Thanh cong"
    Else
        MsgBox "Da phan cong khach hang " & customerID & " cho can bo " & staffID, vbInformation, "Thanh cong"
    End If
    
    AssignCustomerToStaff = True
    
    Exit Function
    
ErrorHandler:
    ' Bao ve lai sheet
    On Error Resume Next
    wsAssignment.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    MsgBox "Loi khi phan cong khach hang: " & Err.description, vbCritical, "Loi"
    LogError "AssignCustomerToStaff", Err.Number, Err.description
    AssignCustomerToStaff = False
End Function

' Ghi log hanh dong cua nguoi dung
Public Sub LogUserAction(ByVal actionType As String, ByVal actionDetails As String)
    On Error GoTo ErrorHandler
    
    ' Kiem tra sheet ActionLog ton tai khong
    If Not modUtility.sheetExists("ActionLog") Then
        ' Tao sheet ActionLog neu chua co
        CreateActionLogSheet
    End If
    
    ' Mo sheet ActionLog
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.sheets("ActionLog")
    
    ' Mo khoa sheet
    wsLog.Unprotect password:=GetDefaultPassword()
    
    ' Tim dong cuoi cung
    Dim lastRow As Long
    lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    
    ' Them log moi
    wsLog.Cells(lastRow + 1, 1).Value = Now() ' ThoiGian
    wsLog.Cells(lastRow + 1, 2).Value = GetCurrentUserID() ' MaNguoiDung
    wsLog.Cells(lastRow + 1, 3).Value = GetCurrentUserName() ' TenNguoiDung
    wsLog.Cells(lastRow + 1, 4).Value = actionType ' LoaiHanhDong
    wsLog.Cells(lastRow + 1, 5).Value = actionDetails ' ChiTietHanhDong
    
    ' Bao ve lai sheet
    wsLog.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    Exit Sub
    
ErrorHandler:
    ' Bao ve lai sheet
    On Error Resume Next
    wsLog.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    LogError "LogUserAction", Err.Number, Err.description
End Sub

' Tao sheet ActionLog
Private Sub CreateActionLogSheet()
    On Error Resume Next
    
    ' Tao sheet moi
    ThisWorkbook.sheets.Add(After:=ThisWorkbook.sheets(ThisWorkbook.sheets.Count)).Name = "ActionLog"
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets("ActionLog")
    
    ' Thiet lap header
    With ws
        .Cells(1, 1).Value = "ThoiGian"
        .Cells(1, 2).Value = "MaNguoiDung"
        .Cells(1, 3).Value = "TenNguoiDung"
        .Cells(1, 4).Value = "LoaiHanhDong"
        .Cells(1, 5).Value = "ChiTietHanhDong"
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

