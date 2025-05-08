Attribute VB_Name = "ModuleDataStructure"
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
        MsgBox "Cau truc du lieu he thong bi hong. Cac sheet sau day bi thieu: " & vbCrLf & _
               missingSheets & vbCrLf & vbCrLf & _
               "He thong se co gang khoi phuc cau truc du lieu. " & _
               "Vui long khoi dong lai ung dung sau khi qua trinh khoi phuc hoan tat.", _
               vbCritical, "Loi cau truc du lieu"
        
        ' Thu hoi khoi phuc cau truc du lieu
        RecreateDataStructure
        
        ValidateRequiredSheets = False
    Else
        ValidateRequiredSheets = True
    End If
    
    On Error GoTo 0
End Function

