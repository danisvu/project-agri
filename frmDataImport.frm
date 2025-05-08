VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataImport 
   Caption         =   "IMPORT DU LIEU KHACH HANG"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765.001
   OleObjectBlob   =   "frmDataImport.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmDataImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form Import du lieu
' Mo ta: Form cho phep nguoi dung chon va import du lieu tu cac file Excel
' Tac gia: Phong KHCN, Agribank Chi nhanh 4
' Ngay tao: 08/05/2025

Option Explicit

' ===========================
' CAC BIEN CUC BO
' ===========================
Private filePaths(1 To 4) As String
Private fileLabels(1 To 4) As String
Private fileChecked(1 To 4) As Boolean
Private fileButtons(1 To 4) As CommandButton
Private fileTextBoxes(1 To 4) As TextBox
Private fileCheckBoxes(1 To 4) As CheckBox
Private fileStatusLabels(1 To 4) As Label

' ===========================
' FORM INITIALIZE/TERMINATE
' ===========================

Private Sub UserForm_Initialize()
    ' Khoi tao cac bien cuc bo
    InitializeControlVariables
    
    ' Thiet lap tieu de form
    Me.Caption = "Import Du Lieu"
    
    ' Thiet lap nhan cho cac loai du lieu
    fileLabels(1) = "Du no (Du no yyyy-mm-dd.xls)"
    fileLabels(2) = "Tai san (Tai san yyyy-mm-dd.xls)"
    fileLabels(3) = "Tra goc (Tra goc mm-yyyy.xls)"
    fileLabels(4) = "Tra lai (Tra lai mm-yyyy.xls)"
    
    ' Cap nhat giao dien
    UpdateFormInterface
    
    ' Kiem tra trang thai du lieu - Hien thi thong tin import gan nhat
    UpdateDataStatus
End Sub

Private Sub UserForm_Terminate()
    ' Giai phong cac bien
    Erase filePaths
    Erase fileLabels
    Erase fileChecked
End Sub

' ===========================
' EVENT HANDLERS
' ===========================

' Xu ly khi nguoi dung nhan nut Browse
Private Sub cmdBrowse1_Click()
    SelectFile 1
End Sub

Private Sub cmdBrowse2_Click()
    SelectFile 2
End Sub

Private Sub cmdBrowse3_Click()
    SelectFile 3
End Sub

Private Sub cmdBrowse4_Click()
    SelectFile 4
End Sub

' Xu ly khi nguoi dung thay doi checkbox
Private Sub chkFile1_Click()
    fileChecked(1) = chkFile1.Value
    UpdateButtonStates
End Sub

Private Sub chkFile2_Click()
    fileChecked(2) = chkFile2.Value
    UpdateButtonStates
End Sub

Private Sub chkFile3_Click()
    fileChecked(3) = chkFile3.Value
    UpdateButtonStates
End Sub

Private Sub chkFile4_Click()
    fileChecked(4) = chkFile4.Value
    UpdateButtonStates
End Sub

' Xu ly khi nguoi dung nhan nut Import
Private Sub cmdImport_Click()
    Dim dataTypes(1 To 4) As String
    Dim importFilePaths(1 To 4) As String
    Dim importCount As Integer
    Dim i As Integer
    
    ' Kiem tra xem co file nao duoc chon khong
    importCount = 0
    For i = 1 To 4
        If fileChecked(i) And Trim(filePaths(i)) <> "" Then
            importCount = importCount + 1
        End If
    Next i
    
    If importCount = 0 Then
        MsgBox "Vui long chon it nhat mot file de import!", vbExclamation, "Import du lieu"
        Exit Sub
    End If
    
    ' Xac nhan import
    If MsgBox("Ban co chac chan muon import du lieu tu " & importCount & " file da chon?", _
             vbQuestion + vbYesNo, "Xac nhan import") = vbNo Then
        Exit Sub
    End If
    
    ' Chuan bi du lieu cho import
    dataTypes(1) = DATA_TYPE_DU_NO
    dataTypes(2) = DATA_TYPE_TAI_SAN
    dataTypes(3) = DATA_TYPE_TRA_GOC
    dataTypes(4) = DATA_TYPE_TRA_LAI
    
    For i = 1 To 4
        If fileChecked(i) Then
            importFilePaths(i) = filePaths(i)
        Else
            importFilePaths(i) = ""
        End If
    Next i
    
    ' An form
    Me.Hide
    
    ' Goi ham import
    ImportData importFilePaths, dataTypes
    
    ' Dong form
    Unload Me
End Sub

' Xu ly khi nguoi dung nhan nut Cancel
Private Sub cmdCancel_Click()
    ' Dong form
    Unload Me
End Sub

' Xu ly khi nguoi dung nhan nut Check All
Private Sub cmdCheckAll_Click()
    Dim i As Integer
    
    ' Chon tat ca cac checkbox
    For i = 1 To 4
        fileChecked(i) = True
    Next i
    
    ' Cap nhat trang thai checkbox
    chkFile1.Value = True
    chkFile2.Value = True
    chkFile3.Value = True
    chkFile4.Value = True
    
    ' Cap nhat trang thai cac nut
    UpdateButtonStates
End Sub

' Xu ly khi nguoi dung nhan nut Uncheck All
Private Sub cmdUncheckAll_Click()
    Dim i As Integer
    
    ' Bo chon tat ca cac checkbox
    For i = 1 To 4
        fileChecked(i) = False
    Next i
    
    ' Cap nhat trang thai checkbox
    chkFile1.Value = False
    chkFile2.Value = False
    chkFile3.Value = False
    chkFile4.Value = False
    
    ' Cap nhat trang thai cac nut
    UpdateButtonStates
End Sub

' ===========================
' CAC PROCEDURE HO TRO
' ===========================

' Khoi tao cac bien cuc bo cho cac control
Private Sub InitializeControlVariables()
    ' Gan cac control vao mang de de dang quan ly
    On Error Resume Next
    
    ' Gan textboxes
    Set fileTextBoxes(1) = txtFile1
    Set fileTextBoxes(2) = txtFile2
    Set fileTextBoxes(3) = txtFile3
    Set fileTextBoxes(4) = txtFile4
    
    ' Gan buttons
    Set fileButtons(1) = cmdBrowse1
    Set fileButtons(2) = cmdBrowse2
    Set fileButtons(3) = cmdBrowse3
    Set fileButtons(4) = cmdBrowse4
    
    ' Gan checkboxes
    Set fileCheckBoxes(1) = chkFile1
    Set fileCheckBoxes(2) = chkFile2
    Set fileCheckBoxes(3) = chkFile3
    Set fileCheckBoxes(4) = chkFile4
    
    ' Gan status labels
    Set fileStatusLabels(1) = lblStatus1
    Set fileStatusLabels(2) = lblStatus2
    Set fileStatusLabels(3) = lblStatus3
    Set fileStatusLabels(4) = lblStatus4
    
    On Error GoTo 0
End Sub

' Cap nhat giao dien form
Private Sub UpdateFormInterface()
    Dim i As Integer
    
    ' Cap nhat cac label
    For i = 1 To 4
        fileCheckBoxes(i).Caption = fileLabels(i)
        fileCheckBoxes(i).Value = False
        fileChecked(i) = False
        
        fileTextBoxes(i).Text = ""
        fileStatusLabels(i).Caption = ""
    Next i
    
    ' Cap nhat trang thai cac nut
    UpdateButtonStates
End Sub

' Cap nhat trang thai cac nut
Private Sub UpdateButtonStates()
    Dim anyChecked As Boolean
    Dim i As Integer
    
    ' Kiem tra xem co checkbox nao duoc chon khong
    anyChecked = False
    For i = 1 To 4
        If fileChecked(i) Then
            anyChecked = True
            Exit For
        End If
    Next i
    
    ' Cap nhat trang thai nut Import
    cmdImport.Enabled = anyChecked
End Sub

' Cap nhat thong tin trang thai du lieu
Private Sub UpdateDataStatus()
    Dim lastImportDate As Date
    Dim daysSinceLastImport As Integer
    Dim statusColor As Long
    Dim statusText As String
    Dim i As Integer
    
    ' Lay thong tin import gan nhat cho tung loai du lieu
    For i = 1 To 4
        Select Case i
            Case 1 ' Du no
                lastImportDate = modImport.GetLastImportDate(ModuleConfig.DATA_TYPE_DU_NO)
            Case 2 ' Tai san
                lastImportDate = modImport.GetLastImportDate(ModuleConfig.DATA_TYPE_TAI_SAN)
            Case 3 ' Tra goc
                lastImportDate = modImport.GetLastImportDate(ModuleConfig.DATA_TYPE_TRA_GOC)
            Case 4 ' Tra lai
                lastImportDate = modImport.GetLastImportDate(ModuleConfig.DATA_TYPE_TRA_LAI)
        End Select
        
        ' Tinh so ngay ke tu lan import cuoi
        If lastImportDate > DateSerial(1900, 1, 2) Then ' Co du lieu import
            daysSinceLastImport = DateDiff("d", lastImportDate, Date)
            
            ' Xac dinh mau va thong bao theo so ngay
            If daysSinceLastImport <= 3 Then
                ' Du lieu moi (0-3 ngay)
                statusColor = ModuleConfig.COLOR_SUCCESS ' Mau xanh la
                statusText = "Import gan nhat: " & Format(lastImportDate, "dd/mm/yyyy") & " (Du lieu moi)"
            ElseIf daysSinceLastImport <= ModuleConfig.DATA_WARNING_DAYS Then
                ' Du lieu binh thuong (4-7 ngay)
                statusColor = ModuleConfig.COLOR_WARNING ' Mau vang
                statusText = "Import gan nhat: " & Format(lastImportDate, "dd/mm/yyyy") & " (" & daysSinceLastImport & " ngay truoc)"
            Else
                ' Du lieu cu (> 7 ngay)
                statusColor = ModuleConfig.COLOR_DANGER ' Mau do
                statusText = "Import gan nhat: " & Format(lastImportDate, "dd/mm/yyyy") & " (" & daysSinceLastImport & " ngay truoc - Du lieu cu)"
            End If
        Else
            ' Chua co du lieu import
            statusColor = ModuleConfig.COLOR_DANGER ' Mau do
            statusText = "Chua co du lieu import"
        End If
        
        ' Cap nhat label trang thai
        fileStatusLabels(i).ForeColor = statusColor
        fileStatusLabels(i).Caption = statusText
    Next i
End Sub

' Lua chon file tu dialog
' @param fileIndex: Chi so cua file (1-4)
Private Sub SelectFile(ByVal fileIndex As Integer)
    On Error Resume Next
    
    ' Tao dialog chon file
    Dim fd As Object
    Dim fileTitle As String
    Dim fileFilter As String
    Dim initialFolder As String
    
    ' Xac dinh tieu de va bo loc cho tung loai file
    Select Case fileIndex
        Case 1 ' Du no
            fileTitle = "Chon file Du no (Du no yyyy-mm-dd.xls)"
            fileFilter = "Du no*.xls*"
        Case 2 ' Tai san
            fileTitle = "Chon file Tai san (Tai san yyyy-mm-dd.xls)"
            fileFilter = "Tai san*.xls*"
        Case 3 ' Tra goc
            fileTitle = "Chon file Tra goc (Tra goc mm-yyyy.xls)"
            fileFilter = "Tra goc*.xls*"
        Case 4 ' Tra lai
            fileTitle = "Chon file Tra lai (Tra lai mm-yyyy.xls)"
            fileFilter = "Tra lai*.xls*"
    End Select
    
    ' Lay duong dan thu muc import
    initialFolder = DEFAULT_IMPORT_PATH
    
    ' Tao doi tuong file dialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Thiet lap thuoc tinh
    With fd
        .Title = fileTitle
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        
        ' Thiet lap thu muc ban dau
        If Dir(initialFolder, vbDirectory) <> "" Then
            .InitialFileName = initialFolder
        End If
        
        ' Hien thi dialog
        If .Show Then
            ' Luu duong dan file
            filePaths(fileIndex) = .SelectedItems(1)
            
            ' Hien thi ten file
            fileTextBoxes(fileIndex).Text = GetFileNameFromPath(filePaths(fileIndex))
            
            ' Tu dong chon checkbox
            fileCheckBoxes(fileIndex).Value = True
            fileChecked(fileIndex) = True
            
            ' Kiem tra tinh hop le cua file
            Dim dataType As String
            Select Case fileIndex
                Case 1: dataType = DATA_TYPE_DU_NO
                Case 2: dataType = DATA_TYPE_TAI_SAN
                Case 3: dataType = DATA_TYPE_TRA_GOC
                Case 4: dataType = DATA_TYPE_TRA_LAI
            End Select
            
            ' Validate va cap nhat trang thai
            ValidateFileAndUpdateStatus fileIndex, dataType
        End If
    End With
    
    ' Cap nhat trang thai cac nut
    UpdateButtonStates
    
    On Error GoTo 0
End Sub

' Kiem tra tinh hop le va cap nhat trang thai cua file
' @param fileIndex: Chi so cua file (1-4)
' @param dataType: Loai du lieu
Private Sub ValidateFileAndUpdateStatus(ByVal fileIndex As Integer, ByVal dataType As String)
    On Error Resume Next
    
    Dim filePath As String
    Dim fileName As String
    Dim isValid As Boolean
    Dim fileDate As Date
    Dim lastImportDate As Date
    Dim statusColor As Long
    Dim statusText As String
    
    ' Lay duong dan va ten file
    filePath = filePaths(fileIndex)
    fileName = modImport.GetFileNameFromPath(filePath)
    
    ' Kiem tra file co ton tai khong
    If Dir(filePath) = "" Then
        statusColor = ModuleConfig.COLOR_DANGER
        statusText = "Loi: File khong ton tai"
        isValid = False
    Else
        ' Kiem tra tinh hop le cua file
        isValid = modImport.ValidateImportFile(filePath, dataType)
        
        If isValid Then
            ' Lay ngay tu ten file
            fileDate = ModuleConfig.ExtractDateFromFileName(fileName, dataType)
            
            ' Lay ngay import gan nhat
            lastImportDate = modImport.GetLastImportDate(dataType)
            
            ' Kiem tra tinh cap nhat
            If fileDate < lastImportDate Then
                ' File cu hon du lieu hien tai
                statusColor = ModuleConfig.COLOR_WARNING
                statusText = "Canh bao: File cu hon du lieu hien tai (" & Format(fileDate, "dd/mm/yyyy") & " < " & Format(lastImportDate, "dd/mm/yyyy") & ")"
            Else
                ' File moi hon hoac bang du lieu hien tai
                statusColor = ModuleConfig.COLOR_SUCCESS
                statusText = "File hop le - Ngay: " & Format(fileDate, "dd/mm/yyyy")
            End If
        Else
            ' File khong hop le
            statusColor = ModuleConfig.COLOR_DANGER
            statusText = "Loi: Ten file khong dung dinh dang yeu cau"
        End If
    End If
    
    ' Cap nhat trang thai
    fileStatusLabels(fileIndex).ForeColor = statusColor
    fileStatusLabels(fileIndex).Caption = statusText
    
    On Error GoTo 0
End Sub

