Attribute VB_Name = "modSecurity"
'Attribute VB_Name = "modSecurity"
' Module bao mat du lieu
' Chua cac ham ho tro bao mat va ma hoa du lieu
' Tac gia: Phong Khach hang Ca nhan, Agribank Chi nhanh 4
' Ngay tao: 15/05/2025

Option Explicit

' Cac hang so lien quan den bao mat
Private Const HASH_ITERATIONS As Integer = 1000 ' So vong lap khi ma hoa mat khau
Private Const MIN_PASSWORD_LENGTH As Integer = 8 ' Do dai mat khau toi thieu
Private Const SALT_LENGTH As Integer = 16 ' Do dai salt cho ma hoa mat khau

' Bao ve noi dung mot sheet
Public Sub SecureSheet(ByVal sheetName As String, Optional ByVal strongProtection As Boolean = False)
    On Error GoTo ErrorHandler
    
    ' Kiem tra sheet ton tai
    If Not modUtility.sheetExists(sheetName) Then
        MsgBox "Sheet " & sheetName & " khong ton tai!", vbExclamation, "Loi"
        Exit Sub
    End If
    
    ' Ma hoa du lieu nhay cam trong sheet (neu co)
    EncryptSensitiveData sheetName
    
    ' Bao ve sheet voi cac tuy chon
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(sheetName)
    
    ' Bo khoa truoc khi dat lai
    On Error Resume Next ' Tranh loi neu sheet chua duoc bao ve
    ws.Unprotect password:=GetDefaultPassword()
    On Error GoTo ErrorHandler
    
    ' Bao ve sheet voi cac tuy chon khac nhau tuy thuoc vao muc do bao mat
    If strongProtection Then
        ' Bao ve chat che, khong cho phep thay doi bat ky thu gi
        ws.Protect password:=GetDefaultPassword(), _
                UserInterfaceOnly:=False, _
                Contents:=True, _
                DrawingObjects:=True, _
                Scenarios:=True, _
                AllowFormattingCells:=False, _
                AllowFormattingColumns:=False, _
                AllowFormattingRows:=False, _
                AllowInsertingColumns:=False, _
                AllowInsertingRows:=False, _
                AllowInsertingHyperlinks:=False, _
                AllowDeletingColumns:=False, _
                AllowDeletingRows:=False, _
                AllowSorting:=False, _
                AllowFiltering:=False, _
                AllowUsingPivotTables:=False
    Else
        ' Bao ve co ban, cho phep mot so thao tac
        ws.Protect password:=GetDefaultPassword(), _
                UserInterfaceOnly:=True, _
                Contents:=True, _
                DrawingObjects:=True, _
                Scenarios:=True, _
                AllowFormattingCells:=True, _
                AllowFormattingColumns:=True, _
                AllowFormattingRows:=True, _
                AllowInsertingColumns:=False, _
                AllowInsertingRows:=False, _
                AllowInsertingHyperlinks:=False, _
                AllowDeletingColumns:=False, _
                AllowDeletingRows:=False, _
                AllowSorting:=True, _
                AllowFiltering:=True, _
                AllowUsingPivotTables:=True
    End If
    
    ' An sheet neu la sheet du lieu - SUA PHAN NAY DE TRANH LOI
    If IsDataSheet(sheetName) Then
        On Error Resume Next ' Tranh loi khi thiet lap thuoc tinh Visible
        
        ' Kiem tra trang thai hien tai truoc khi thay doi
        If ws.Visible <> xlSheetVeryHidden Then
            ws.Visible = xlSheetVeryHidden
        End If
        
        If Err.Number <> 0 Then
            ' Neu khong the set VeryHidden, thu Hidden
            Err.Clear
            If ws.Visible <> xlSheetHidden Then
                ws.Visible = xlSheetHidden
            End If
            
            If Err.Number <> 0 Then
                ' Ghi log va bo qua neu van khong duoc
                Debug.Print "Khong the an sheet " & sheetName & " (Sheet co the da duoc an truoc do)"
                Err.Clear
            End If
        End If
        
        On Error GoTo ErrorHandler
    End If

    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi bao ve sheet " & sheetName & ": " & Err.description, vbCritical, "Loi"
    LogError "SecureSheet", Err.Number, Err.description
End Sub

' Kiem tra mot sheet co phai la sheet du lieu khong
' Them ham IsDataSheet de ho tro
Private Function IsDataSheet(ByVal sheetName As String) As Boolean
    On Error Resume Next
    
    Dim dataSheets As Variant
    dataSheets = GetRequiredDataSheets()
    
    Dim i As Integer
    For i = LBound(dataSheets) To UBound(dataSheets)
        If dataSheets(i) = sheetName Then
            IsDataSheet = True
            Exit Function
        End If
    Next i
    
    IsDataSheet = False
End Function

' Ma hoa du lieu nhay cam trong sheet
Private Sub EncryptSensitiveData(ByVal sheetName As String)
    On Error GoTo ErrorHandler
    
    ' Chi ma hoa du lieu nhay cam trong cac sheet cu the
    Select Case sheetName
        Case SHEET_USERS
            EncryptUserPasswords
        
        Case Else
            ' Khong lam gi doi voi cac sheet khac
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi ma hoa du lieu nhay cam trong sheet " & sheetName & ": " & Err.description, vbCritical, "Loi"
    LogError "EncryptSensitiveData", Err.Number, Err.description
End Sub

' Ma hoa mat khau nguoi dung
Private Sub EncryptUserPasswords()
    On Error GoTo ErrorHandler
    
    ' Kiem tra sheet Users co ton tai khong
    If Not modUtility.sheetExists(SHEET_USERS) Then Exit Sub
    
    ' Mo khoa sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(SHEET_USERS)
    ws.Unprotect password:=GetDefaultPassword()
    
    ' Tim so dong co du lieu
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Neu khong co du lieu, thoat
    If lastRow <= 1 Then
        ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
        Exit Sub
    End If
    
    ' Ma hoa tung mat khau chua duoc ma hoa
    Dim i As Long
    For i = 2 To lastRow ' Bat dau tu dong 2, bo qua header
        ' Chi ma hoa mat khau chua duoc ma hoa
        If Len(ws.Cells(i, 3).Value) > 0 And Len(ws.Cells(i, 3).Value) <= 20 Then
            ' Kiem tra xem mat khau da duoc ma hoa chua (mat khau ma hoa thong thuong dai hon 30 ky tu)
            ' Neu mat khau ngan hon 30 ky tu, co the day la mat khau chua ma hoa
            ws.Cells(i, 3).Value = EnhancedHashPassword(ws.Cells(i, 3).Value)
        End If
    Next i
    
    ' Bao ve lai sheet
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    Exit Sub
    
ErrorHandler:
    ' Bao ve lai sheet
    On Error Resume Next
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    LogError "EncryptUserPasswords", Err.Number, Err.description
End Sub

' Ma hoa mat khau voi bao mat cao hon
Public Function EnhancedHashPassword(ByVal password As String) As String
    On Error GoTo ErrorHandler
    
    ' Ham ma hoa mat khau nang cao hon, ket hop salt va nhieu vong lap
    ' Luu y: Day van chi la mo phong, trong thuc te can su dung thu vien ma hoa tot hon
    
    ' Tao salt ngau nhien
    Dim salt As String
    salt = GenerateRandomSalt(SALT_LENGTH)
    
    ' Ket hop mat khau voi salt
    Dim base As String
    base = password & salt
    
    ' Thuc hien nhieu vong lap ma hoa
    Dim hash As String
    hash = base
    
    Dim i As Integer
    For i = 1 To HASH_ITERATIONS
        hash = SimpleHash(hash)
    Next i
    
    ' Ket qua cuoi cung bao gom salt de co the kiem tra mat khau sau nay
    EnhancedHashPassword = hash & ":" & salt
    
    Exit Function
    
ErrorHandler:
    LogError "EnhancedHashPassword", Err.Number, Err.description
    EnhancedHashPassword = ""
End Function

' Kiem tra mat khau co dung voi hash khong
Public Function VerifyPassword(ByVal inputPassword As String, ByVal storedHash As String) As Boolean
    On Error GoTo ErrorHandler
    
    VerifyPassword = False
    
    ' Tach hash va salt
    Dim parts() As String
    parts = Split(storedHash, ":")
    
    ' Kiem tra dinh dang hash
    If UBound(parts) < 1 Then
        ' Kiem tra mat khau voi ham ma hoa don gian
        ' (cho tuong thich voi cach ma hoa cu)
        If HashPassword(inputPassword) = storedHash Then
            VerifyPassword = True
        End If
        Exit Function
    End If
    
    ' Lay salt tu hash da luu
    Dim salt As String
    salt = parts(1)
    
    ' Ket hop mat khau nhap vao voi salt
    Dim base As String
    base = inputPassword & salt
    
    ' Thuc hien nhieu vong lap ma hoa
    Dim hash As String
    hash = base
    
    Dim i As Integer
    For i = 1 To HASH_ITERATIONS
        hash = SimpleHash(hash)
    Next i
    
    ' So sanh hash
    If hash = parts(0) Then
        VerifyPassword = True
    End If
    
    Exit Function
    
ErrorHandler:
    LogError "VerifyPassword", Err.Number, Err.description
    VerifyPassword = False
End Function

' Tao salt ngau nhien
Private Function GenerateRandomSalt(ByVal length As Integer) As String
    On Error GoTo ErrorHandler
    
    Dim salt As String
    Dim chars As String
    Dim i As Integer
    
    ' Cac ky tu duoc phep trong salt
    chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    
    ' Tao chuoi salt ngau nhien
    salt = ""
    
    ' Khoi tao bo tao so ngau nhien
    Randomize
    
    For i = 1 To length
        salt = salt & Mid(chars, Int((Len(chars) * Rnd) + 1), 1)
    Next i
    
    GenerateRandomSalt = salt
    
    Exit Function
    
ErrorHandler:
    LogError "GenerateRandomSalt", Err.Number, Err.description
    GenerateRandomSalt = ""
End Function

' Ham ma hoa don gian
Private Function SimpleHash(ByVal text As String) As String
    On Error GoTo ErrorHandler
    
    ' Mo phong mot ham ma hoa don gian dua tren MD5 (khong phai MD5 that)
    ' Trong thuc te, nen su dung thu vien ma hoa tot hon
    
    Dim result As String
    Dim i As Long
    Dim char As Integer
    Dim charSum As Long
    
    ' Tinh tong ma ASCII cua cac ky tu
    charSum = 0
    For i = 1 To Len(text)
        char = Asc(Mid(text, i, 1))
        charSum = charSum + char * i
    Next i
    
    ' Tao ket qua dua tren tong ma ASCII
    result = ""
    For i = 1 To 32 ' Do dai 32 ky tu, tuong tu MD5
        char = ((charSum * i + Asc(Mid(text, i Mod Len(text) + 1, 1))) Mod 16) + 1
        result = result & Mid("0123456789abcdef", char, 1)
    Next i
    
    SimpleHash = result
    
    Exit Function
    
ErrorHandler:
    LogError "SimpleHash", Err.Number, Err.description
    SimpleHash = ""
End Function

' Bao ve ma VBA
Public Sub ProtectVBACode()
    On Error GoTo ErrorHandler
    
    ' Luu y: Trong VBA, khong the tu dong bao ve ma VBA
    ' Can huong dan nguoi dung cach bao ve bang tay
    
    MsgBox "De bao ve ma VBA, thuc hien cac buoc sau:" & vbCrLf & _
           "1. Mo Tools -> VBAProject Properties" & vbCrLf & _
           "2. Chon tab Protection" & vbCrLf & _
           "3. Chon 'Lock project for viewing'" & vbCrLf & _
           "4. Nhap mat khau va xac nhan mat khau" & vbCrLf & _
           "5. Nhan OK", vbInformation, "Huong dan bao ve ma VBA"
    
    Exit Sub
    
ErrorHandler:
    LogError "ProtectVBACode", Err.Number, Err.description
End Sub

' Kiem tra do phuc tap cua mat khau
Public Function IsStrongPassword(ByVal password As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Mac dinh la mat khau yeu
    IsStrongPassword = False
    
    ' Kiem tra do dai
    If Len(password) < MIN_PASSWORD_LENGTH Then Exit Function
    
    ' Kiem tra co it nhat mot chu hoa
    If Not ContainsUpperCase(password) Then Exit Function
    
    ' Kiem tra co it nhat mot chu thuong
    If Not ContainsLowerCase(password) Then Exit Function
    
    ' Kiem tra co it nhat mot chu so
    If Not ContainsDigit(password) Then Exit Function
    
    ' Kiem tra co it nhat mot ky tu dac biet
    If Not ContainsSpecialChar(password) Then Exit Function
    
    ' Neu qua tat ca kiem tra, mat khau manh
    IsStrongPassword = True
    
    Exit Function
    
ErrorHandler:
    LogError "IsStrongPassword", Err.Number, Err.description
    IsStrongPassword = False
End Function

' Kiem tra chuoi co chua chu hoa khong
Private Function ContainsUpperCase(ByVal text As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Long
    
    For i = 1 To Len(text)
        If Asc(Mid(text, i, 1)) >= 65 And Asc(Mid(text, i, 1)) <= 90 Then
            ContainsUpperCase = True
            Exit Function
        End If
    Next i
    
    ContainsUpperCase = False
    
    Exit Function
    
ErrorHandler:
    ContainsUpperCase = False
End Function

' Kiem tra chuoi co chua chu thuong khong
Private Function ContainsLowerCase(ByVal text As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Long
    
    For i = 1 To Len(text)
        If Asc(Mid(text, i, 1)) >= 97 And Asc(Mid(text, i, 1)) <= 122 Then
            ContainsLowerCase = True
            Exit Function
        End If
    Next i
    
    ContainsLowerCase = False
    
    Exit Function
    
ErrorHandler:
    ContainsLowerCase = False
End Function

' Kiem tra chuoi co chua chu so khong
Private Function ContainsDigit(ByVal text As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Long
    
    For i = 1 To Len(text)
        If Asc(Mid(text, i, 1)) >= 48 And Asc(Mid(text, i, 1)) <= 57 Then
            ContainsDigit = True
            Exit Function
        End If
    Next i
    
    ContainsDigit = False
    
    Exit Function
    
ErrorHandler:
    ContainsDigit = False
End Function

' Kiem tra chuoi co chua ky tu dac biet khong
Private Function ContainsSpecialChar(ByVal text As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim specialChars As String
    specialChars = "!@#$%^&*()_-+={}[]|\:;""'<>,.?/~`"
    
    Dim i As Long, j As Long
    
    For i = 1 To Len(text)
        For j = 1 To Len(specialChars)
            If Mid(text, i, 1) = Mid(specialChars, j, 1) Then
                ContainsSpecialChar = True
                Exit Function
            End If
        Next j
    Next i
    
    ContainsSpecialChar = False
    
    Exit Function
    
ErrorHandler:
    ContainsSpecialChar = False
End Function

' Bao ve toan bo workbook
Public Sub SecureWorkbook()
    On Error GoTo ErrorHandler
    
    ' Bao ve tung sheet
    Dim ws As Worksheet
    
    ' Bao ve cac sheet du lieu voi bao mat manh
    Dim dataSheets As Variant
    dataSheets = GetRequiredDataSheets()
    
    Dim i As Integer
    For i = LBound(dataSheets) To UBound(dataSheets)
        If modUtility.sheetExists(CStr(dataSheets(i))) Then
            On Error Resume Next ' Tranh loi khi bao ve tung sheet
            SecureSheet CStr(dataSheets(i)), True ' Bao mat manh cho sheet du lieu
            If Err.Number <> 0 Then
                Debug.Print "Loi khi bao ve sheet " & dataSheets(i) & ": " & Err.description
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        End If
    Next i
    
    ' Bao ve cac sheet cau hinh voi bao mat rat manh
    Dim configSheets As Variant
    configSheets = GetRequiredConfigSheets()
    
    For i = LBound(configSheets) To UBound(configSheets)
        If modUtility.sheetExists(CStr(configSheets(i))) Then
            On Error Resume Next ' Tranh loi khi bao ve tung sheet
            SecureSheet CStr(configSheets(i)), True ' Bao mat manh cho sheet cau hinh
            If Err.Number <> 0 Then
                Debug.Print "Loi khi bao ve sheet " & configSheets(i) & ": " & Err.description
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        End If
    Next i
    
    ' Bao ve cac sheet hien thi voi bao mat thap hon
    For Each ws In ThisWorkbook.sheets
        If Not IsDataOrConfigSheet(ws.Name) Then
            On Error Resume Next ' Tranh loi khi bao ve tung sheet
            SecureSheet ws.Name, False ' Bao mat thap hon cho sheet hien thi
            If Err.Number <> 0 Then
                Debug.Print "Loi khi bao ve sheet " & ws.Name & ": " & Err.description
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        End If
    Next ws
    
    ' Bao ve cau truc workbook - TRY/CATCH de tranh loi
    On Error Resume Next
    ThisWorkbook.Protect password:=GetDefaultPassword(), Structure:=True, Windows:=False
    If Err.Number <> 0 Then
        Debug.Print "Loi khi bao ve workbook: " & Err.description
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    MsgBox "Da bao ve toan bo workbook thanh cong!", vbInformation, "Bao mat"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi bao ve workbook: " & Err.description, vbCritical, "Loi"
    LogError "SecureWorkbook", Err.Number, Err.description
End Sub

' Kiem tra co phai sheet du lieu hoac cau hinh khong
Public Function IsDataOrConfigSheet(sheetName As String) As Boolean
    On Error Resume Next
    
    ' Danh sach cac sheet can bao ve
    Dim protectedSheets As Variant
    protectedSheets = Array(SHEET_RAW_DU_NO, SHEET_RAW_TAI_SAN, SHEET_RAW_TRA_GOC, SHEET_RAW_TRA_LAI, _
                           SHEET_PROCESSED_DATA, SHEET_IMPORT_LOG, SHEET_TRANSACTION_DATA, _
                           SHEET_STAFF_ASSIGNMENT, SHEET_CONFIG, SHEET_USERS)
    
    Dim i As Integer
    For i = LBound(protectedSheets) To UBound(protectedSheets)
        If sheetName = protectedSheets(i) Then
            IsDataOrConfigSheet = True
            Exit Function
        End If
    Next i
    
    IsDataOrConfigSheet = False
End Function

' Bao ve du lieu nhay cam
Public Sub MaskSensitiveData()
    On Error GoTo ErrorHandler
    
    ' Bao ve du lieu nhay cam cho muc dich hien thi
    ' Luu y: Du lieu goc van duoc luu tru, chi co cach hien thi la bi che
    
    ' Chi ap dung cho sheet du lieu
    Dim dataSheets As Variant
    dataSheets = GetRequiredDataSheets()
    
    Dim i As Integer
    For i = LBound(dataSheets) To UBound(dataSheets)
        If modUtility.sheetExists(CStr(dataSheets(i))) Then
            MaskSensitiveDataInSheet CStr(dataSheets(i))
        End If
    Next i
    
    MsgBox "Da che du lieu nhay cam thanh cong!", vbInformation, "Bao mat"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi che du lieu nhay cam: " & Err.description, vbCritical, "Loi"
    LogError "MaskSensitiveData", Err.Number, Err.description
End Sub

' Che du lieu nhay cam trong mot sheet
Private Sub MaskSensitiveDataInSheet(ByVal sheetName As String)
    On Error GoTo ErrorHandler
    
    ' Mo khoa sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(sheetName)
    ws.Unprotect password:=GetDefaultPassword()
    
    ' Xac dinh cac cot chua du lieu nhay cam
    Dim sensitiveColumns As New Collection
    
    ' Them cac cot nhay cam tuy thuoc vao loai sheet
    Select Case sheetName
        Case SHEET_RAW_DU_NO
            ' Che so du, so tien
            sensitiveColumns.Add 6 ' SoTienPheDuyet
            sensitiveColumns.Add 7 ' SoTienGiaiNgan
            sensitiveColumns.Add 9 ' SoDuHienTai
            
        Case SHEET_RAW_TAI_SAN
            ' Che gia tri tai san
            sensitiveColumns.Add 11 ' GiaTriTaiSan
            sensitiveColumns.Add 16 ' GiaTriKhaDung
            sensitiveColumns.Add 17 ' GiaTriTheChan
            
        Case SHEET_RAW_TRA_GOC, SHEET_RAW_TRA_LAI
            ' Che so tien
            sensitiveColumns.Add 6 ' SoTienPhaiTra
            sensitiveColumns.Add 7 ' SoDuHienTai
            
        Case SHEET_USERS
            ' Che mat khau
            sensitiveColumns.Add 3 ' MatKhau
    End Select
    
    ' Ap dung NumberFormat de che du lieu nhay cam
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim j As Long, k As Long
    For j = 1 To sensitiveColumns.Count
        ' Ap dung cho tat ca cac o trong cot (tru dong tieu de)
        For k = 2 To lastRow
            ' Chi ap dung neu o khong rong
            If Not IsEmpty(ws.Cells(k, sensitiveColumns(j)).Value) Then
                ' Ap dung dinh dang che du lieu
                Select Case sensitiveColumns(j)
                    Case 3 ' MatKhau
                        ' Che hoan toan mat khau
                        ws.Cells(k, sensitiveColumns(j)).NumberFormat = ";;;"
                    Case Else
                        ' Che mot phan du lieu so
                        ws.Cells(k, sensitiveColumns(j)).NumberFormat = """***""0"
                End Select
            End If
        Next k
    Next j
    
    ' Bao ve lai sheet
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    Exit Sub
    
ErrorHandler:
    ' Bao ve lai sheet
    On Error Resume Next
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True
    
    LogError "MaskSensitiveDataInSheet", Err.Number, Err.description
End Sub

' Thay doi mat khau workbook
Public Sub ChangeWorkbookPassword()
    On Error GoTo ErrorHandler
    
    ' Hien thi form thay doi mat khau
    Dim oldPassword As String
    Dim newPassword As String
    Dim confirmPassword As String
    
    ' Nhap mat khau cu
    oldPassword = InputBox("Nhap mat khau hien tai:", "Thay doi mat khau")
    
    ' Kiem tra mat khau cu
    If oldPassword <> GetDefaultPassword() Then
        MsgBox "Mat khau khong dung!", vbExclamation, "Loi"
        Exit Sub
    End If
    
    ' Nhap mat khau moi
    newPassword = InputBox("Nhap mat khau moi (toi thieu " & MIN_PASSWORD_LENGTH & " ky tu):", "Thay doi mat khau")
    
    ' Kiem tra mat khau moi
    If Len(newPassword) < MIN_PASSWORD_LENGTH Then
        MsgBox "Mat khau moi qua ngan!", vbExclamation, "Loi"
        Exit Sub
    End If
    
    ' Kiem tra mat khau manh
    If Not IsStrongPassword(newPassword) Then
        MsgBox "Mat khau moi khong du manh!" & vbCrLf & _
               "Mat khau phai co it nhat " & MIN_PASSWORD_LENGTH & " ky tu va bao gom:" & vbCrLf & _
               "- It nhat mot chu hoa" & vbCrLf & _
               "- It nhat mot chu thuong" & vbCrLf & _
               "- It nhat mot chu so" & vbCrLf & _
               "- It nhat mot ky tu dac biet", vbExclamation, "Loi"
        Exit Sub
    End If
    
    ' Xac nhan mat khau moi
    confirmPassword = InputBox("Xac nhan mat khau moi:", "Thay doi mat khau")
    
    ' Kiem tra xac nhan
    If newPassword <> confirmPassword Then
        MsgBox "Xac nhan mat khau khong khop!", vbExclamation, "Loi"
        Exit Sub
    End If
    
    ' Cap nhat mat khau trong cau hinh
    ' Luu y: Can thay doi mat khau o tat ca cac noi su dung mat khau
    ' Thay doi mat khau trong modConfig
    MsgBox "Khong the thay doi mat khau tu dong." & vbCrLf & _
           "Ban can sua mat khau trong module modConfig, hang so DEFAULT_PASSWORD.", vbInformation, "Thong bao"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Loi khi thay doi mat khau: " & Err.description, vbCritical, "Loi"
    LogError "ChangeWorkbookPassword", Err.Number, Err.description
End Sub

