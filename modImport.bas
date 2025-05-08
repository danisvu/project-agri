Attribute VB_Name = "modImport"
' Module xu ly import va dong bo du lieu
' Mo ta: Module quan ly viec import va dong bo du lieu tu cac file Excel vao he thong
' Tac gia: Phong KHCN, Agribank Chi nhanh 4
' Ngay tao: 08/05/2025

Option Explicit

' ===========================
' HANG SO VA KHAI BAO
' ===========================

' Khai bao hang so duong dan va ten sheet
' Cac hang so nay da duoc dinh nghia trong modConfig va modDataStructure
' Private Const SHEET_DU_NO As String = "Raw_DuNo"
' Private Const SHEET_TAI_SAN As String = "Raw_TaiSan"
' Private Const SHEET_TRA_GOC As String = "Raw_TraGoc"
' Private Const SHEET_TRA_LAI As String = "Raw_TraLai"
' Private Const SHEET_IMPORT_LOG As String = "ImportLog"
' Private Const DEFAULT_IMPORT_PATH As String = "C:\Agribank\Import\"

' Hang so dac thu cho module import
Private Const MAX_IMPORT_ROWS As Long = 1000000        ' So dong toi da co the import
Private Const MAX_COLS_DU_NO As Integer = 80           ' So cot toi da trong file Du no
Private Const MAX_COLS_TAI_SAN As Integer = 40         ' So cot toi da trong file Tai san
Private Const MAX_COLS_TRA_GOC As Integer = 50         ' So cot toi da trong file Tra goc
Private Const MAX_COLS_TRA_LAI As Integer = 50         ' So cot toi da trong file Tra lai
Private Const FIELD_DELIMITER As String = ";"          ' Ky tu phan cach trong file CSV/XLS

' ===========================
' CAC PROCEDURE CHINH
' ===========================

' Procedure chinh xu ly import du lieu
' @param filePaths: Mang duong dan den cac file can import
' @param dataTypes: Mang loai du lieu tuong ung voi tung file
Public Sub ImportData(ByRef filePaths() As String, ByRef dataTypes() As String)
    On Error GoTo ErrorHandler
    
    ' Toi uu hieu suat
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    
    ' Khai bao bien
    Dim i As Integer
    Dim fileCount As Integer
    Dim importStatus As Boolean
    Dim fileDate As Date
    Dim lastImportDate As Date
    Dim cancelImport As Boolean
    Dim importResults() As String
    Dim ws As Worksheet
    Dim startTime As Double
    
    ' Thiet lap mang ket qua
    fileCount = UBound(filePaths) - LBound(filePaths) + 1
    ReDim importResults(1 To fileCount, 1 To 5) ' [FilePath, DataType, Status, RecordsProcessed, Notes]
    
    ' Bat dau do thoi gian
    startTime = Timer
    
    ' Thong bao bat dau
    Application.StatusBar = "Bat dau qua trinh import..."
    
    ' Xu ly tung file
    For i = LBound(filePaths) To UBound(filePaths)
        If filePaths(i) <> "" Then
            ' Cap nhat thanh trang thai
            Application.StatusBar = "Dang xu ly file " & (i - LBound(filePaths) + 1) & "/" & fileCount & ": " & _
                                   GetFileNameFromPath(filePaths(i))
            
            ' Kiem tra tinh hop le cua file
            If ValidateImportFile(filePaths(i), dataTypes(i)) Then
                ' Kiem tra tinh cap nhat cua file
                fileDate = ExtractDateFromFileName(GetFileNameFromPath(filePaths(i)), dataTypes(i))
                lastImportDate = GetLastImportDate(dataTypes(i))
                
                ' Canh bao neu file cu hon du lieu hien tai
                If fileDate < lastImportDate Then
                    If MsgBox("File " & GetFileNameFromPath(filePaths(i)) & " co ve cu hon du lieu hien tai!" & vbCrLf & _
                            "Ngay cua file: " & Format(fileDate, "dd/mm/yyyy") & vbCrLf & _
                            "Ngay import gan nhat: " & Format(lastImportDate, "dd/mm/yyyy") & vbCrLf & vbCrLf & _
                            "Ban co chac muon tiep tuc import?", _
                            vbQuestion + vbYesNo, "Canh bao file cu") = vbNo Then
                        ' Danh dau khong import file nay
                        importResults(i, 1) = filePaths(i)
                        importResults(i, 2) = dataTypes(i)
                        importResults(i, 3) = "Cancelled"
                        importResults(i, 4) = "0"
                        importResults(i, 5) = "File cu hon du lieu hien tai"
                        GoTo NextFile
                    End If
                End If
                
                ' Thuc hien import du lieu
                importStatus = ProcessImportFile(filePaths(i), dataTypes(i), importResults(i, 4), importResults(i, 5))
                
                ' Ghi nhan ket qua
                importResults(i, 1) = filePaths(i)
                importResults(i, 2) = dataTypes(i)
                importResults(i, 3) = IIf(importStatus, "Success", "Failed")
                
                ' Ghi log import
                LogImportActivity filePaths(i), dataTypes(i), fileDate, importStatus, _
                                 CLng(importResults(i, 4)), importResults(i, 5)
            Else
                ' Ghi nhan loi validate
                importResults(i, 1) = filePaths(i)
                importResults(i, 2) = dataTypes(i)
                importResults(i, 3) = "Failed"
                importResults(i, 4) = "0"
                importResults(i, 5) = "File khong hop le hoac khong dung dinh dang"
                
                ' Ghi log loi
                LogImportActivity filePaths(i), dataTypes(i), fileDate, False, 0, _
                                "File khong hop le hoac khong dung dinh dang"
            End If
        End If
NextFile:
    Next i
    
    ' Cap nhat bang tong hop (Processed_Data)
    Application.StatusBar = "Dang cap nhat du lieu tong hop..."
    UpdateProcessedData
    
    ' Hien thi ket qua import
    DisplayImportResults importResults, Timer - startTime
    
    ' Xoa mang ket qua
    Erase importResults
    
    ' Thong bao hoan thanh
    Application.StatusBar = "Hoan thanh qua trinh import!"
    Application.Wait Now + TimeValue("00:00:01")
    Application.StatusBar = False
    
    ' Cap nhat bang bien toan cuc
    UpdateGlobalDataStatus

CleanUp:
    ' Khoi phuc cai dat
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    ' Xu ly loi
    LogErrorDetailed "ImportData", Err.Number, Err.description, ErrorSeverity_High, _
                    "Import paths: " & Join(filePaths, ", ")
    
    MsgBox "Da xay ra loi trong qua trinh import: " & vbCrLf & _
           Err.description, vbCritical, "Loi Import"
    
    Resume CleanUp
End Sub

' Procedure xu ly import du lieu tu mot file
' @param filePath: Duong dan den file can import
' @param dataType: Loai du lieu cua file
' @param recordsProcessed: Tham chieu den bien chua so ban ghi da xu ly
' @param notes: Tham chieu den bien chua ghi chu
' @return: TRUE neu import thanh cong, FALSE neu that bai
Private Function ProcessImportFile(ByVal filePath As String, ByVal dataType As String, _
                                 ByRef recordsProcessed As String, ByRef notes As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Khai bao bien
    Dim wbSource As Workbook
    Dim wsTarget As Worksheet
    Dim data As Variant
    Dim rowCount As Long
    Dim targetSheet As String
    Dim recordsNew As Long
    Dim recordsUpdated As Long
    Dim recordsUnchanged As Long
    
    ' Xac dinh target sheet tuong ung
    targetSheet = GetSheetNameForDataType(dataType)
    If targetSheet = "" Then
        notes = "Khong xac dinh duoc sheet dich cho loai du lieu " & dataType
        ProcessImportFile = False
        Exit Function
    End If
    
    ' Mo file Excel nguon voi che do chi doc
    Set wbSource = Workbooks.Open(filePath, ReadOnly:=True)
    
    ' Doc du lieu vao mang de toi uu hieu suat
    data = ReadDataFromWorkbook(wbSource, dataType)
    
    ' Dong file nguon
    wbSource.Close SaveChanges:=False
    
    ' Kiem tra du lieu
    If Not IsArray(data) Then
        notes = "Loi khi doc du lieu tu file nguon"
        ProcessImportFile = False
        Exit Function
    End If
    
    ' Tim sheet dich
    Set wsTarget = ThisWorkbook.Sheets(targetSheet)
    
    ' Dong bo du lieu vao sheet dich
    SynchronizeData data, wsTarget, dataType, recordsNew, recordsUpdated, recordsUnchanged
    
    ' Cap nhat ket qua
    rowCount = UBound(data, 1) - 1 ' Tru di dong tieu de
    recordsProcessed = CStr(rowCount)
    notes = "Them moi: " & recordsNew & ", Cap nhat: " & recordsUpdated & _
           ", Khong doi: " & recordsUnchanged
    
    ' Tra ve ket qua thanh cong
    ProcessImportFile = True
    Exit Function
    
ErrorHandler:
    notes = "Loi: " & Err.description
    
    ' Dong workbook nguon neu da mo
    If Not wbSource Is Nothing Then
        wbSource.Close SaveChanges:=False
    End If
    
    ' Ghi log loi
    LogErrorDetailed "ProcessImportFile", Err.Number, Err.description, ErrorSeverity_High, _
                    "File: " & filePath & ", DataType: " & dataType
    
    ProcessImportFile = False
End Function

' Ham doc du lieu tu workbook nguon vao mang
' @param wbSource: Workbook nguon
' @param dataType: Loai du lieu
' @return: Mang chua du lieu hoac NULL neu co loi
Private Function ReadDataFromWorkbook(ByVal wbSource As Workbook, ByVal dataType As String) As Variant
    On Error GoTo ErrorHandler
    
    ' Khai bao bien
    Dim wsSource As Worksheet
    Dim data As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim maxCols As Integer
    Dim usedRange As Range
    
    ' Xac dinh so cot toi da theo loai du lieu
    Select Case dataType
        Case DATA_TYPE_DU_NO
            maxCols = MAX_COLS_DU_NO
        Case DATA_TYPE_TAI_SAN
            maxCols = MAX_COLS_TAI_SAN
        Case DATA_TYPE_TRA_GOC
            maxCols = MAX_COLS_TRA_GOC
        Case DATA_TYPE_TRA_LAI
            maxCols = MAX_COLS_TRA_LAI
        Case Else
            maxCols = 100 ' Gia tri mac dinh
    End Select
    
    ' Su dung Worksheet dau tien
    Set wsSource = wbSource.Sheets(1)
    
    ' Tim dong va cot cuoi cung co du lieu
    lastRow = wsSource.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = WorksheetFunction.Min(wsSource.Cells.Find(What:="*", SearchOrder:=xlByColumns, _
                                                      SearchDirection:=xlPrevious).Column, maxCols)
    
    ' Kiem tra so dong
    If lastRow > MAX_IMPORT_ROWS Then
        MsgBox "File qua lon! So dong vuot qua gioi han cho phep (" & MAX_IMPORT_ROWS & ").", _
               vbExclamation, "Loi Import"
        ReadDataFromWorkbook = Null
        Exit Function
    End If
    
    ' Gioi han so cot theo cau hinh
    lastCol = WorksheetFunction.Min(lastCol, maxCols)
    
    ' Doc toan bo du lieu vao mang
    Set usedRange = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastCol))
    data = usedRange.Value
    
    ' Tra ve mang du lieu
    ReadDataFromWorkbook = data
    Exit Function
    
ErrorHandler:
    ' Ghi log loi
    LogErrorDetailed "ReadDataFromWorkbook", Err.Number, Err.description, ErrorSeverity_High, _
                    "DataType: " & dataType
    
    ReadDataFromWorkbook = Null
End Function

' Ham dong bo du lieu vao sheet dich
' @param data: Mang chua du lieu nguon
' @param wsTarget: Worksheet dich
' @param dataType: Loai du lieu
' @param recordsNew: Tham chieu den bien dem so ban ghi moi
' @param recordsUpdated: Tham chieu den bien dem so ban ghi cap nhat
' @param recordsUnchanged: Tham chieu den bien dem so ban ghi khong thay doi
Private Sub SynchronizeData(ByRef data As Variant, ByRef wsTarget As Worksheet, ByVal dataType As String, _
                          ByRef recordsNew As Long, ByRef recordsUpdated As Long, _
                          ByRef recordsUnchanged As Long)
    On Error GoTo ErrorHandler
    
    ' Khai bao bien
    Dim sourceRowCount As Long
    Dim targetLastRow As Long
    Dim i As Long, j As Long, k As Long
    Dim primaryKeyCol As Integer
    Dim primaryKeyExists As Boolean
    Dim primaryKey As String
    Dim rowData() As Variant
    Dim targetHeaderRow As Variant
    Dim dataMapping() As Integer
    Dim targetRows As Object
    Dim currentRow As Long
    Dim oldProtection As Boolean
    
    ' Khoi tao bien dem
    recordsNew = 0
    recordsUpdated = 0
    recordsUnchanged = 0
    
    ' Lay so dong cua du lieu nguon (tru dong tieu de)
    sourceRowCount = UBound(data, 1)
    
    ' Bo bao ve sheet tam thoi
    oldProtection = UnprotectSheetWithPassword(wsTarget)
    
    ' Doc dong tieu de cua sheet dich
    targetHeaderRow = wsTarget.Range("A1").Resize(1, wsTarget.Cells(1, Columns.Count).End(xlToLeft).Column).Value
    
    ' Xac dinh cot khoa chinh theo loai du lieu
    primaryKeyCol = GetPrimaryKeyColumn(dataType)
    
    ' Tao mapping giua cot nguon va cot dich
    ReDim dataMapping(1 To UBound(data, 2))
    For j = 1 To UBound(data, 2)
        dataMapping(j) = FindColumnIndexInTarget(data(1, j), targetHeaderRow)
    Next j
    
    ' Tao Dictionary de luu tru cac dong hien co trong sheet dich (dua theo khoa chinh)
    Set targetRows = CreateObject("Scripting.Dictionary")
    
    ' Tim dong cuoi cung co du lieu trong sheet dich
    targetLastRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
    If targetLastRow > 1 Then ' Co du lieu (ngoai dong tieu de)
        ' Doc tat ca cac khoa chinh hien co trong sheet dich
        Dim targetKeys As Variant
        targetKeys = wsTarget.Range(wsTarget.Cells(2, primaryKeyCol), wsTarget.Cells(targetLastRow, primaryKeyCol)).Value
        
        ' Luu tru vi tri dong cua moi khoa chinh
        For i = 1 To UBound(targetKeys, 1)
            If Not IsEmpty(targetKeys(i, 1)) And Not IsNull(targetKeys(i, 1)) Then
                If Not targetRows.Exists(CStr(targetKeys(i, 1))) Then
                    targetRows.Add CStr(targetKeys(i, 1)), i + 1 ' +1 vi bat dau tu dong 2 (sau tieu de)
                End If
            End If
        Next i
    End If
    
    ' Hien thi thanh tien trinh
    Application.StatusBar = "Dang dong bo du lieu 0%"
    
    ' Xu ly tung dong du lieu trong mang nguon (bo qua dong tieu de)
    For i = 2 To sourceRowCount
        ' Cap nhat thanh tien trinh
        If i Mod 100 = 0 Then
            Application.StatusBar = "Dang dong bo du lieu " & Format((i - 1) / (sourceRowCount - 1) * 100, "0") & "%"
            DoEvents
        End If
        
        ' Lay gia tri khoa chinh cua dong hien tai
        primaryKey = CStr(data(i, primaryKeyCol))
        
        ' Bo qua neu khoa chinh trong
        If Trim(primaryKey) = "" Then GoTo NextSourceRow
        
        ' Kiem tra xem khoa chinh da ton tai trong sheet dich chua
        primaryKeyExists = targetRows.Exists(primaryKey)
        
        ' Xu ly them moi hoac cap nhat
        If primaryKeyExists Then
            ' Lay vi tri dong trong sheet dich
            currentRow = targetRows(primaryKey)
            
            ' Kiem tra xem du lieu co thay doi khong
            Dim dataChanged As Boolean
            dataChanged = False
            
            ' Tao mang chua du lieu cua dong hien tai
            ReDim rowData(1 To UBound(targetHeaderRow, 2))
            
            ' Gan gia tri tu du lieu nguon vao mang rowData theo mapping
            For j = 1 To UBound(data, 2)
                If dataMapping(j) > 0 Then
                    rowData(dataMapping(j)) = data(i, j)
                End If
            Next j
            
            ' So sanh voi du lieu hien co trong sheet dich
            Dim currentRowData As Variant
            currentRowData = wsTarget.Range(wsTarget.Cells(currentRow, 1), _
                                         wsTarget.Cells(currentRow, UBound(targetHeaderRow, 2))).Value
            
            ' Kiem tra su khac biet
            For j = 1 To UBound(targetHeaderRow, 2)
                If j <> primaryKeyCol Then ' Khong so sanh cot khoa chinh
                    If Not IsEmpty(rowData(j)) And Not IsNull(rowData(j)) Then
                        If CStr(rowData(j)) <> CStr(currentRowData(1, j)) Then
                            dataChanged = True
                            Exit For
                        End If
                    End If
                End If
            Next j
            
            ' Cap nhat neu co thay doi
            If dataChanged Then
                ' Cap nhat tung cot co du lieu moi
                For j = 1 To UBound(data, 2)
                    If dataMapping(j) > 0 Then
                        wsTarget.Cells(currentRow, dataMapping(j)).Value = data(i, j)
                    End If
                Next j
                
                recordsUpdated = recordsUpdated + 1
            Else
                recordsUnchanged = recordsUnchanged + 1
            End If
        Else
            ' Them dong moi
            currentRow = targetLastRow + 1
            targetLastRow = currentRow
            
            ' Tao mang chua du lieu cua dong moi
            ReDim rowData(1 To UBound(targetHeaderRow, 2))
            
            ' Gan gia tri tu du lieu nguon vao mang rowData theo mapping
            For j = 1 To UBound(data, 2)
                If dataMapping(j) > 0 Then
                    rowData(dataMapping(j)) = data(i, j)
                End If
            Next j
            
            ' Ghi du lieu vao sheet dich
            For j = 1 To UBound(rowData)
                wsTarget.Cells(currentRow, j).Value = rowData(j)
            Next j
            
            ' Them khoa chinh moi vao Dictionary
            targetRows.Add primaryKey, currentRow
            
            recordsNew = recordsNew + 1
        End If
        
NextSourceRow:
    Next i
    
    ' Khoi phuc bao ve sheet
    If oldProtection Then ProtectSheetWithPassword wsTarget
    
    Exit Sub
    
ErrorHandler:
    ' Khoi phuc bao ve sheet
    If oldProtection Then ProtectSheetWithPassword wsTarget
    
    ' Ghi log loi
    LogErrorDetailed "SynchronizeData", Err.Number, Err.description, ErrorSeverity_High, _
                    "DataType: " & dataType
    
    ' Hien thi thong bao loi
    MsgBox "Da xay ra loi khi dong bo du lieu: " & vbCrLf & _
           Err.description, vbCritical, "Loi Dong bo"
           
    ' Re-throw loi
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

' ===========================
' CAC FUNCTION HO TRO
' ===========================

' Ham kiem tra tinh hop le cua file import
' @param filePath: Duong dan den file can kiem tra
' @param dataType: Loai du lieu cua file
' @return: TRUE neu file hop le, FALSE neu khong
Private Function ValidateImportFile(ByVal filePath As String, ByVal dataType As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Khai bao bien
    Dim fileName As String
    Dim fileExt As String
    Dim expectedPattern As String
    Dim validExtensions As String
    
    ' Lay ten file va phan mo rong
    fileName = GetFileNameFromPath(filePath)
    fileExt = GetFileExtension(filePath)
    
    ' Kiem tra phan mo rong
    validExtensions = ".xls;.xlsx;.xlsm;"
    If InStr(1, validExtensions, "." & LCase(fileExt) & ";") = 0 Then
        MsgBox "File khong dung dinh dang Excel: " & fileName, vbExclamation, "Loi dinh dang"
        ValidateImportFile = False
        Exit Function
    End If
    
    ' Kiem tra ten file theo loai du lieu
    Select Case dataType
        Case DATA_TYPE_DU_NO
            expectedPattern = DU_NO_FILE_PATTERN
            If Not MatchesPattern(fileName, expectedPattern) Then
                MsgBox "Ten file khong dung dinh dang yeu cau: " & fileName & vbCrLf & _
                      "Dinh dang yeu cau: " & DU_NO_FILE_PATTERN, vbExclamation, "Loi ten file"
                ValidateImportFile = False
                Exit Function
            End If
            
        Case DATA_TYPE_TAI_SAN
            expectedPattern = TAI_SAN_FILE_PATTERN
            If Not MatchesPattern(fileName, expectedPattern) Then
                MsgBox "Ten file khong dung dinh dang yeu cau: " & fileName & vbCrLf & _
                      "Dinh dang yeu cau: " & TAI_SAN_FILE_PATTERN, vbExclamation, "Loi ten file"
                ValidateImportFile = False
                Exit Function
            End If
            
        Case DATA_TYPE_TRA_GOC
            expectedPattern = TRA_GOC_FILE_PATTERN
            If Not MatchesPattern(fileName, expectedPattern) Then
                MsgBox "Ten file khong dung dinh dang yeu cau: " & fileName & vbCrLf & _
                      "Dinh dang yeu cau: " & TRA_GOC_FILE_PATTERN, vbExclamation, "Loi ten file"
                ValidateImportFile = False
                Exit Function
            End If
            
        Case DATA_TYPE_TRA_LAI
            expectedPattern = TRA_LAI_FILE_PATTERN
            If Not MatchesPattern(fileName, expectedPattern) Then
                MsgBox "Ten file khong dung dinh dang yeu cau: " & fileName & vbCrLf & _
                      "Dinh dang yeu cau: " & TRA_LAI_FILE_PATTERN, vbExclamation, "Loi ten file"
                ValidateImportFile = False
                Exit Function
            End If
            
        Case Else
            MsgBox "Loai du lieu khong hop le: " & dataType, vbExclamation, "Loi loai du lieu"
            ValidateImportFile = False
            Exit Function
    End Select
    
    ' Kiem tra file co ton tai khong
    If Dir(filePath) = "" Then
        MsgBox "File khong ton tai: " & filePath, vbExclamation, "Loi file"
        ValidateImportFile = False
        Exit Function
    End If
    
    ' Kiem tra kich thuoc file
    Dim fileSize As Long
    fileSize = FileLen(filePath)
    If fileSize > MAX_IMPORT_FILE_SIZE Then
        MsgBox "File qua lon! Kich thuoc toi da cho phep la " & _
              Format(MAX_IMPORT_FILE_SIZE / ONE_MB, "0.00") & " MB." & vbCrLf & _
              "Kich thuoc file hien tai: " & Format(fileSize / ONE_MB, "0.00") & " MB.", _
              vbExclamation, "Loi kich thuoc file"
        ValidateImportFile = False
        Exit Function
    End If
    
    ' Neu qua tat ca kiem tra, file hop le
    ValidateImportFile = True
    Exit Function
    
ErrorHandler:
    ' Ghi log loi
    LogErrorDetailed "ValidateImportFile", Err.Number, Err.description, ErrorSeverity_Medium, _
                    "File: " & filePath & ", DataType: " & dataType
    
    ValidateImportFile = False
End Function

' Ham kiem tra su phu hop giua ten file va mau quy dinh
' @param fileName: Ten file can kiem tra
' @param pattern: Mau quy dinh
' @return: TRUE neu phu hop, FALSE neu khong
Private Function MatchesPattern(ByVal fileName As String, ByVal pattern As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Chuyen doi mau thanh Regular Expression
    Dim regexPattern As String
    regexPattern = "^" & Replace(Replace(Replace(Replace(pattern, ".", "\."), "?", "."), "*", ".*"), " ", " ") & "$"
    
    ' Tao doi tuong RegExp
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Thiet lap thuoc tinh
    regex.Global = False
    regex.IgnoreCase = True
    regex.pattern = regexPattern
    
    ' Kiem tra su phu hop
    MatchesPattern = regex.Test(fileName)
    Exit Function
    
ErrorHandler:
    ' Neu co loi, su dung phuong phap don gian hon
    Dim filePattern As String
    filePattern = pattern
    
    ' Thay the cac ky tu dac biet
    filePattern = Replace(filePattern, ".", "\.")
    filePattern = Replace(filePattern, "?", ".")
    filePattern = Replace(filePattern, "*", ".*")
    
    ' So sanh voi mau don gian
    ' NOTE: Day chi la giai phap thay the tam thoi
    ' Trong thuc te, nen su dung Regular Expression de so sanh chinh xac hon
    MatchesPattern = (fileName Like pattern)
End Function

' Ham lay ten file tu duong dan day du
' @param filePath: Duong dan day du den file
' @return: Ten file (bao gom phan mo rong)
Private Function GetFileNameFromPath(ByVal filePath As String) As String
    Dim pos As Integer
    pos = InStrRev(filePath, "\")
    
    If pos > 0 Then
        GetFileNameFromPath = Mid(filePath, pos + 1)
    Else
        GetFileNameFromPath = filePath
    End If
End Function

' Ham lay phan mo rong cua file
' @param filePath: Duong dan hoac ten file
' @return: Phan mo rong cua file (khong bao gom dau cham)
Private Function GetFileExtension(ByVal filePath As String) As String
    Dim fileName As String
    Dim pos As Integer
    
    ' Lay ten file truoc
    fileName = GetFileNameFromPath(filePath)
    
    ' Tim vi tri dau cham cuoi cung
    pos = InStrRev(fileName, ".")
    
    If pos > 0 Then
        GetFileExtension = Mid(fileName, pos + 1)
    Else
        GetFileExtension = ""
    End If
End Function

' Ham lay ngay import gan nhat cho loai du lieu
' @param dataType: Loai du lieu
' @return: Ngay import gan nhat
Private Function GetLastImportDate(ByVal dataType As String) As Date
    On Error GoTo ErrorHandler
    
    ' Khai bao bien
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim lastDate As Date
    
    ' Khoi tao gia tri mac dinh
    lastDate = DateSerial(1900, 1, 1)
    
    ' Tim worksheet ImportLog
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_IMPORT_LOG)
    On Error GoTo ErrorHandler
    
    ' Neu khong tim thay sheet, tra ve gia tri mac dinh
    If ws Is Nothing Then
        GetLastImportDate = lastDate
        Exit Function
    End If
    
    ' Tim dong cuoi cung co du lieu
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Neu chi co dong tieu de, tra ve gia tri mac dinh
    If lastRow <= 1 Then
        GetLastImportDate = lastDate
        Exit Function
    End If
    
    ' Tim ngay import gan nhat cho loai du lieu
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 3).Value = dataType And ws.Cells(i, 11).Value = "Success" Then
            lastDate = ws.Cells(i, 4).Value
            Exit For
        End If
    Next i
    
    ' Tra ve ngay tim duoc
    GetLastImportDate = lastDate
    Exit Function
    
ErrorHandler:
    ' Ghi log loi
    LogErrorDetailed "GetLastImportDate", Err.Number, Err.description, ErrorSeverity_Low, _
                    "DataType: " & dataType
    
    ' Tra ve gia tri mac dinh
    GetLastImportDate = DateSerial(1900, 1, 1)
End Function

' Ham ghi log hoat dong import
' @param filePath: Duong dan file da import
' @param dataType: Loai du lieu
' @param fileDate: Ngay tao file
' @param success: Trang thai thanh cong hay that bai
' @param recordCount: So ban ghi da xu ly
' @param notes: Ghi chu
Private Sub LogImportActivity(ByVal filePath As String, ByVal dataType As String, _
                            ByVal fileDate As Date, ByVal success As Boolean, _
                            ByVal recordCount As Long, ByVal notes As String)
    On Error GoTo ErrorHandler
    
    ' Khai bao bien
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim fileName As String
    Dim importID As String
    Dim recordsNew As Long
    Dim recordsUpdated As Long
    Dim recordsUnchanged As Long
    
    ' Lay ten file
    fileName = GetFileNameFromPath(filePath)
    
    ' Tim worksheet ImportLog
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_IMPORT_LOG)
    On Error GoTo ErrorHandler
    
    ' Neu khong tim thay sheet, thoat
    If ws Is Nothing Then Exit Sub
    
    ' Bo bao ve sheet tam thoi
    Dim oldProtection As Boolean
    oldProtection = UnprotectSheetWithPassword(ws)
    
    ' Tim dong cuoi cung co du lieu
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Tao ID cho ban ghi import
    importID = Format(Now, "yyyymmddhhnnss") & "_" & dataType
    
    ' Phan tich ghi chu de lay so ban ghi moi/cap nhat/khong doi
    If InStr(1, notes, "Them moi: ") > 0 Then
        ' Format notes: "Them moi: X, Cap nhat: Y, Khong doi: Z"
        recordsNew = CLng(Mid(notes, InStr(1, notes, "Them moi: ") + 10, _
                            InStr(1, notes, ", Cap nhat:") - (InStr(1, notes, "Them moi: ") + 10)))
        recordsUpdated = CLng(Mid(notes, InStr(1, notes, "Cap nhat: ") + 10, _
                              InStr(1, notes, ", Khong doi:") - (InStr(1, notes, "Cap nhat: ") + 10)))
        recordsUnchanged = CLng(Mid(notes, InStr(1, notes, "Khong doi: ") + 11))
    End If
    
    ' Ghi ban ghi moi vao ImportLog
    lastRow = lastRow + 1
    ws.Cells(lastRow, 1).Value = importID
    ws.Cells(lastRow, 2).Value = fileName
    ws.Cells(lastRow, 3).Value = dataType
    ws.Cells(lastRow, 4).Value = fileDate
    ws.Cells(lastRow, 5).Value = Now
    ws.Cells(lastRow, 6).Value = gCurrentUser
    ws.Cells(lastRow, 7).Value = recordCount
    ws.Cells(lastRow, 8).Value = recordsNew
    ws.Cells(lastRow, 9).Value = recordsUpdated
    ws.Cells(lastRow, 10).Value = 0 ' So ban ghi xoa (khong co trong truong hop nay)
    ws.Cells(lastRow, 11).Value = IIf(success, "Success", "Failed")
    ws.Cells(lastRow, 12).Value = notes
    
    ' Dinh dang cac o
    With ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, 12))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' Khoi phuc bao ve sheet
    If oldProtection Then ProtectSheetWithPassword ws
    
    Exit Sub
    
ErrorHandler:
    ' Khoi phuc bao ve sheet
    If oldProtection Then ProtectSheetWithPassword ws
    
    ' Ghi log loi
    LogErrorDetailed "LogImportActivity", Err.Number, Err.description, ErrorSeverity_Medium, _
                    "File: " & filePath & ", DataType: " & dataType
End Sub

' Ham cap nhat bang du lieu tong hop (Processed_Data)
Private Sub UpdateProcessedData()
    On Error GoTo ErrorHandler
    
    ' Thong bao, phan nay se duoc phat trien sau
    Application.StatusBar = "Dang cap nhat du lieu tong hop..."
    
    ' TODO: Phat trien code cap nhat bang tong hop tu cac nguon du lieu khac nhau
    ' Day la tinh nang phuc tap can phat trien rieng
    
    Exit Sub
    
ErrorHandler:
    ' Ghi log loi
    LogErrorDetailed "UpdateProcessedData", Err.Number, Err.description, ErrorSeverity_Medium, ""
End Sub

' Ham cap nhat bien toan cuc ve trang thai du lieu
Private Sub UpdateGlobalDataStatus()
    On Error GoTo ErrorHandler
    
    ' Khai bao bien
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim latestDate As Date
    Dim latestUser As String
    Dim latestType As String
    
    ' Tim worksheet ImportLog
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_IMPORT_LOG)
    On Error GoTo ErrorHandler
    
    ' Neu khong tim thay sheet, thoat
    If ws Is Nothing Then Exit Sub
    
    ' Tim dong cuoi cung co du lieu
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Neu chi co dong tieu de, thoat
    If lastRow <= 1 Then Exit Sub
    
    ' Tim ban ghi import thanh cong gan nhat
    Dim i As Long
    latestDate = DateSerial(1900, 1, 1)
    
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 11).Value = "Success" Then
            If ws.Cells(i, 5).Value > latestDate Then
                latestDate = ws.Cells(i, 5).Value
                latestUser = ws.Cells(i, 6).Value
                latestType = ws.Cells(i, 3).Value
            End If
        End If
    Next i
    
    ' Cap nhat bien toan cuc
    gDataLastImportDate = latestDate
    gDataLastImportBy = latestUser
    gDataLastImportType = latestType
    
    Exit Sub
    
ErrorHandler:
    ' Ghi log loi
    LogErrorDetailed "UpdateGlobalDataStatus", Err.Number, Err.description, ErrorSeverity_Low, ""
End Sub

' Ham hien thi ket qua import
' @param importResults: Mang chua ket qua import
' @param elapsedTime: Thoi gian da troi qua (giay)
Private Sub DisplayImportResults(ByRef importResults() As String, ByVal elapsedTime As Double)
    On Error GoTo ErrorHandler
    
    ' Khai bao bien
    Dim resultMsg As String
    Dim successCount As Integer
    Dim failCount As Integer
    Dim cancelCount As Integer
    Dim totalRecords As Long
    Dim i As Integer
    
    ' Tinh tong so ban ghi va ket qua
    successCount = 0
    failCount = 0
    cancelCount = 0
    totalRecords = 0
    
    For i = LBound(importResults, 1) To UBound(importResults, 1)
        Select Case importResults(i, 3)
            Case "Success"
                successCount = successCount + 1
                totalRecords = totalRecords + CLng(importResults(i, 4))
            Case "Failed"
                failCount = failCount + 1
            Case "Cancelled"
                cancelCount = cancelCount + 1
        End Select
    Next i
    
    ' Tao thong bao ket qua
    resultMsg = "Ket qua import:" & vbCrLf & vbCrLf
    resultMsg = resultMsg & "- Tong so file da xu ly: " & UBound(importResults, 1) - LBound(importResults, 1) + 1 & vbCrLf
    resultMsg = resultMsg & "- So file import thanh cong: " & successCount & vbCrLf
    resultMsg = resultMsg & "- So file gap loi: " & failCount & vbCrLf
    resultMsg = resultMsg & "- So file bo qua: " & cancelCount & vbCrLf
    resultMsg = resultMsg & "- Tong so ban ghi da xu ly: " & Format(totalRecords, "#,##0") & vbCrLf
    resultMsg = resultMsg & "- Thoi gian xu ly: " & Format(elapsedTime, "0.00") & " giay" & vbCrLf & vbCrLf
    
    ' Them thong tin chi tiet cho tung file
    resultMsg = resultMsg & "Thong tin chi tiet:" & vbCrLf
    
    For i = LBound(importResults, 1) To UBound(importResults, 1)
        If importResults(i, 1) <> "" Then
            resultMsg = resultMsg & vbCrLf & "* File: " & GetFileNameFromPath(importResults(i, 1)) & vbCrLf
            resultMsg = resultMsg & "  - Loai du lieu: " & importResults(i, 2) & vbCrLf
            resultMsg = resultMsg & "  - Trang thai: " & importResults(i, 3) & vbCrLf
            resultMsg = resultMsg & "  - So ban ghi: " & importResults(i, 4) & vbCrLf
            resultMsg = resultMsg & "  - Ghi chu: " & importResults(i, 5) & vbCrLf
        End If
    Next i
    
    ' Hien thi thong bao
    If successCount > 0 Then
        MsgBox resultMsg, vbInformation, "Ket qua Import"
    ElseIf cancelCount > 0 And failCount = 0 Then
        MsgBox resultMsg, vbExclamation, "Import bi huy"
    Else
        MsgBox resultMsg, vbCritical, "Loi Import"
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Ghi log loi
    LogErrorDetailed "DisplayImportResults", Err.Number, Err.description, ErrorSeverity_Low, ""
    
    ' Hien thi thong bao don gian
    MsgBox "Da hoan thanh qua trinh import, xem chi tiet trong sheet ImportLog.", _
          vbInformation, "Import hoan tat"
End Sub

' Ham tim vi tri cot khoa chinh theo loai du lieu
' @param dataType: Loai du lieu
' @return: Vi tri cot khoa chinh
Private Function GetPrimaryKeyColumn(ByVal dataType As String) As Integer
    On Error GoTo ErrorHandler
    
    Select Case dataType
        Case DATA_TYPE_DU_NO
            GetPrimaryKeyColumn = 1 ' MaKhoanVay
        Case DATA_TYPE_TAI_SAN
            GetPrimaryKeyColumn = 1 ' MaTaiSan
        Case DATA_TYPE_TRA_GOC
            GetPrimaryKeyColumn = 1 ' MaLichTraGoc
        Case DATA_TYPE_TRA_LAI
            GetPrimaryKeyColumn = 1 ' MaLichTraLai
        Case Else
            GetPrimaryKeyColumn = 1 ' Mac dinh
    End Select
    
    Exit Function
    
ErrorHandler:
    GetPrimaryKeyColumn = 1 ' Tra ve gia tri mac dinh
End Function

' Ham tim vi tri cot trong mang tieu de dich
' @param sourceColumnName: Ten cot nguon
' @param targetHeaderRow: Mang chua tieu de dich
' @return: Vi tri cot trong bang dich, 0 neu khong tim thay
Private Function FindColumnIndexInTarget(ByVal sourceColumnName As Variant, ByRef targetHeaderRow As Variant) As Integer
    On Error GoTo ErrorHandler
    
    Dim j As Integer
    
    ' Mac dinh la khong tim thay
    FindColumnIndexInTarget = 0
    
    ' Kiem tra tham so
    If IsNull(sourceColumnName) Or IsEmpty(sourceColumnName) Then Exit Function
    
    ' Tim vi tri cot trong targetHeaderRow
    For j = 1 To UBound(targetHeaderRow, 2)
        If Trim(CStr(sourceColumnName)) = Trim(CStr(targetHeaderRow(1, j))) Then
            FindColumnIndexInTarget = j
            Exit Function
        End If
    Next j
    
    Exit Function
    
ErrorHandler:
    FindColumnIndexInTarget = 0
End Function

' Ham bo bao ve sheet voi mat khau
' @param ws: Worksheet can bo bao ve
' @return: TRUE neu sheet da duoc bao ve truoc do
Private Function UnprotectSheetWithPassword(ByRef ws As Worksheet) As Boolean
    On Error Resume Next
    
    ' Kiem tra xem sheet co dang duoc bao ve khong
    UnprotectSheetWithPassword = ws.ProtectContents
    
    ' Bo bao ve neu co
    If UnprotectSheetWithPassword Then
        ws.Unprotect password:=GetDefaultPassword()
    End If
    
    On Error GoTo 0
End Function

' Ham bao ve sheet voi mat khau
' @param ws: Worksheet can bao ve
Private Sub ProtectSheetWithPassword(ByRef ws As Worksheet)
    On Error Resume Next
    
    ' Bao ve sheet
    ws.Protect password:=GetDefaultPassword(), UserInterfaceOnly:=True, _
             AllowSorting:=True, AllowFiltering:=True
    
    On Error GoTo 0
End Sub

' ===========================
' PROCEDURES FORM IMPORT
' ===========================

' Procedure hien thi form import va xu ly du lieu
Public Sub ShowImportForm()
    On Error GoTo ErrorHandler
    
    ' Khai bao bien
    Dim filePaths(1 To 4) As String
    Dim dataTypes(1 To 4) As String
    Dim importAll As Boolean
    Dim i As Integer  ' Khai bao bien i
    
    ' Hien thi form import
    If Not frmDataImport Is Nothing Then
        Unload frmDataImport
    End If
    
    ' Khi nao form duoc phat trien, se thay bang viec hien thi form
    ' Tam thoi, su dung OpenDialog de chon file
    
    ' Khai bao cac chuoi loai du lieu
    dataTypes(1) = DATA_TYPE_DU_NO
    dataTypes(2) = DATA_TYPE_TAI_SAN
    dataTypes(3) = DATA_TYPE_TRA_GOC
    dataTypes(4) = DATA_TYPE_TRA_LAI
    
    ' Cho phep nguoi dung chon cac file tuong ung
    For i = 1 To 4
        Select Case i
            Case 1
                If MsgBox("Ban co muon import du lieu DU NO?", vbQuestion + vbYesNo, "Import du lieu") = vbYes Then
                    filePaths(i) = GetFileFromDialog("Chon file Du no (Du no yyyy-mm-dd.xls)", _
                                                  "Files Excel (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", _
                                                  DEFAULT_IMPORT_PATH)
                End If
            Case 2
                If MsgBox("Ban co muon import du lieu TAI SAN?", vbQuestion + vbYesNo, "Import du lieu") = vbYes Then
                    filePaths(i) = GetFileFromDialog("Chon file Tai san (Tai san yyyy-mm-dd.xls)", _
                                                  "Files Excel (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", _
                                                  DEFAULT_IMPORT_PATH)
                End If
            Case 3
                If MsgBox("Ban co muon import du lieu TRA GOC?", vbQuestion + vbYesNo, "Import du lieu") = vbYes Then
                    filePaths(i) = GetFileFromDialog("Chon file Tra goc (Tra goc mm-yyyy.xls)", _
                                                  "Files Excel (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", _
                                                  DEFAULT_IMPORT_PATH)
                End If
            Case 4
                If MsgBox("Ban co muon import du lieu TRA LAI?", vbQuestion + vbYesNo, "Import du lieu") = vbYes Then
                    filePaths(i) = GetFileFromDialog("Chon file Tra lai (Tra lai mm-yyyy.xls)", _
                                                  "Files Excel (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", _
                                                  DEFAULT_IMPORT_PATH)
                End If
        End Select
    Next i
    
    ' Kiem tra xem co it nhat mot file duoc chon
    Dim hasFile As Boolean
    hasFile = False
    
    For i = 1 To 4
        If filePaths(i) <> "" Then
            hasFile = True
            Exit For
        End If
    Next i
    
    If Not hasFile Then
        MsgBox "Khong co file nao duoc chon de import!", vbExclamation, "Import du lieu"
        Exit Sub
    End If
    
    ' Xac nhan import
    If MsgBox("Ban co chac chan muon import du lieu tu cac file da chon?", _
             vbQuestion + vbYesNo, "Xac nhan import") = vbYes Then
        ' Goi ham import du lieu
        ImportData filePaths, dataTypes
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Ghi log loi
    LogErrorDetailed "ShowImportForm", Err.Number, Err.description, ErrorSeverity_Medium, ""
    
    MsgBox "Da xay ra loi khi hien thi form import: " & vbCrLf & _
           Err.description, vbCritical, "Loi Import"
End Sub

' Ham hien thi hop thoai chon file
' @param dialogTitle: Tieu de hop thoai
' @param fileFilter: Bo loc file
' @param initialFolder: Thu muc ban dau
' @return: Duong dan den file duoc chon
Private Function GetFileFromDialog(ByVal dialogTitle As String, ByVal fileFilter As String, _
                                 ByVal initialFolder As String) As String
    On Error GoTo ErrorHandler
    
    ' Khai bao bien
    Dim fd As Object
    
    ' Tao doi tuong FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Thiet lap thuoc tinh
    With fd
        .Title = dialogTitle
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add fileFilter, "*.xls; *.xlsx; *.xlsm"
        
        ' Thiet lap thu muc ban dau
        If Dir(initialFolder, vbDirectory) <> "" Then
            .InitialFileName = initialFolder
        End If
        
        ' Hien thi hop thoai
        If .Show Then
            GetFileFromDialog = .SelectedItems(1)
        Else
            GetFileFromDialog = ""
        End If
    End With
    
    Exit Function
    
ErrorHandler:
    GetFileFromDialog = ""
End Function

