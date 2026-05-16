Attribute VB_Name = "XLSX成果表表头自动最终版"
Option Explicit

Dim startTime As Double
Dim processedCount As Long
Dim skippedCount As Long
Dim ErrorCount As Long
Dim dict1 As Object  ' Sheet1：键→行数据集合（Collection）
Dim dict2 As Object  ' Sheet2：键→行数据集合（Collection）
Dim renameMapStage1 As Object  ' 阶段1：原始路径→新路径
Dim renameMapStage2 As Object  ' 阶段2：原始路径→新路径
Dim fso As Object  ' 文件系统对象
Dim keyCounts1 As Object  ' Sheet1：键→出现次数
Dim keyCounts2 As Object  ' Sheet2：键→出现次数

Sub ProcessFiles()
    Dim rootFolder As Object
    Dim mappingTablePath As String
    
    ' 初始化统计
    InitializeStats
    
    ' 选择根目录
    Set rootFolder = SelectRootFolder
    If rootFolder Is Nothing Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 查找对应表
    mappingTablePath = FindMappingTable(rootFolder.path)
    If mappingTablePath = "" Then
        MsgBox "未找到对应表.xlsx！", vbExclamation
        Exit Sub
    End If
    
' ########## 阶段1处理（Sheet1：独立文件，如Q开头） ##########
If Not LoadMappingTable(mappingTablePath, 1) Then
    MsgBox "加载Sheet1失败！", vbCritical
    Exit Sub
End If
Application.ScreenUpdating = False
CreateFileCopies rootFolder, 1  ' 为Sheet1重复键创建副本（如Q178→Q178_stage1_copy1.xlsx）
ProcessFolders rootFolder, 1    ' 遍历所有文件，处理Sheet1对应的文件
BatchRenameStage1               ' 阶段1重命名（用Sheet1的K列）
Application.ScreenUpdating = True

' ########## 阶段2处理（Sheet2：独立文件，如Z开头） ##########
If Not LoadMappingTable(mappingTablePath, 2) Then
    MsgBox "加载Sheet2失败！", vbCritical
    Exit Sub
End If
Application.ScreenUpdating = False
CreateFileCopies rootFolder, 2  ' 添加：为Sheet2重复键创建副本（如Z10→Z10_stage2_copy1.xlsx）
ProcessFolders rootFolder, 2    ' 修改：遍历所有文件，处理Sheet2对应的文件（不再依赖阶段1）
BatchRenameStage2               ' 阶段2重命名（用Sheet2的H列）
Application.ScreenUpdating = True
    
    ' 显示结果
    ShowSummary
End Sub

' 初始化全局变量
Private Sub InitializeStats()
    startTime = Timer
    processedCount = 0
    skippedCount = 0
    ErrorCount = 0
    Set dict1 = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    Set renameMapStage1 = CreateObject("Scripting.Dictionary")
    Set renameMapStage2 = CreateObject("Scripting.Dictionary")
    Set keyCounts1 = CreateObject("Scripting.Dictionary")
    Set keyCounts2 = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub

' 根据对应表创建副本文件（阶段1专用）
Private Sub CreateFileCopies(folder As Object, stage As Integer)
    Dim subFolder As Object
    Dim file As Object
    Dim keyCounts As Object
    
    ' 选择当前阶段的键次数字典（Sheet1→keyCounts1，Sheet2→keyCounts2）
    Set keyCounts = IIf(stage = 1, keyCounts1, keyCounts2)
    
    On Error Resume Next
    For Each file In folder.files  ' 修正：Files（大写）
        If IsValidExcelFile(file) Then
            Dim fileName As String: fileName = GetBaseName(file.name)
            ' 为重复键创建副本（数量=出现次数-1）
            If keyCounts.Exists(fileName) Then
                Dim count As Long: count = keyCounts(fileName)
                If count > 1 Then
                    ' 修正：去掉括号（未用Call关键字）
                    CreateCopiesForFile file, count - 1, stage
                End If
            End If
        End If
    Next
    
    ' 递归处理子文件夹（修正：SubFolders（大写））
    For Each subFolder In folder.subFolders
        CreateFileCopies subFolder, stage
    Next
    On Error GoTo 0
End Sub

' 创建副本文件（通用化，支持阶段1/2）
Private Sub CreateCopiesForFile(file As Object, copyCount As Long, stage As Integer)
    Dim i As Long
    Dim parentPath As String: parentPath = file.parentFolder.path
    Dim baseName As String: baseName = GetBaseName(file.name)
    
    For i = 1 To copyCount
        Dim newPath As String: newPath = parentPath & "\" & baseName & "_stage" & stage & "_copy" & i & ".xlsx"
        If Not fso.FileExists(newPath) Then
            fso.CopyFile file.path, newPath  ' 正确：CopyFile方法（FSO）
        End If
    Next
End Sub

' 加载对应表（Sheet1/Sheet2），存储为键→行数据集合
Private Function LoadMappingTable(filePath As String, sheetType As Integer) As Boolean
    On Error GoTo ErrorHandler
    Dim wb As Workbook, ws As Worksheet
    Dim dict As Object, keyCounts As Object
    Dim lastRow As Long, i As Long, j As Long, key As String
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)
    Set ws = wb.Sheets(IIf(sheetType = 1, 1, 2))
    Set dict = IIf(sheetType = 1, dict1, dict2)
    Set keyCounts = IIf(sheetType = 1, keyCounts1, keyCounts2)
    
    ' 清空旧数据
    dict.RemoveAll
    keyCounts.RemoveAll
    
    ' 读取行数据（从第2行开始）
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    For i = 2 To lastRow
        If Not IsEmpty(ws.Cells(i, 1)) Then
            key = CStr(ws.Cells(i, 1).value)
            ' 存储整行11列数据
            Dim arr(1 To 11)
            For j = 1 To 11
                arr(j) = IIf(j <= ws.Columns.count, ws.Cells(i, j).value, "")
            Next
            ' 添加到集合（不覆盖，保留所有行）
            If Not dict.Exists(key) Then
                Set dict(key) = New Collection
            End If
            dict(key).Add arr
            ' 更新键出现次数
            keyCounts(key) = dict(key).count
        End If
    Next
    
    wb.Close SaveChanges:=False
    LoadMappingTable = True
    Exit Function
    
ErrorHandler:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    LoadMappingTable = False
End Function

' 处理文件夹中的所有文件（阶段1专用）
Private Sub ProcessFolders(folder As Object, stage As Integer)
    Dim subFolder As Object, file As Object
    On Error Resume Next
    For Each file In folder.files
        ProcessSingleFile file, stage
    Next
    For Each subFolder In folder.subFolders
        ProcessFolders subFolder, stage
    Next
    On Error GoTo 0
End Sub

' 处理单个文件（核心逻辑：映射行数据+重命名）
Private Sub ProcessSingleFile(file As Object, stage As Integer)
    If Not IsValidExcelFile(file) Then
        skippedCount = skippedCount + 1
        Exit Sub
    End If
    
    Dim wb As Workbook, fileName As String, baseKey As String, rowIndex As Integer
    Dim dict As Object, arrData As Variant, newName As String
    Dim renameMap As Object ' 当前阶段的重命名映射（阶段1→renameMapStage1，阶段2→renameMapStage2）
    
    ' 选择当前阶段的字典和重命名映射
    Select Case stage
        Case 1
            Set dict = dict1
            Set renameMap = renameMapStage1
        Case 2
            Set dict = dict2
            Set renameMap = renameMapStage2
    End Select
    
    fileName = GetBaseName(file.name)
    
    ' ########## 关键：判断是原始文件还是副本，映射到对应行数据 ##########
    If InStr(fileName, "_stage" & stage & "_copy") > 0 Then
        ' 副本文件：提取基础键和序号（如"Z10_stage2_copy1"→基础键"Z10"，序号1→行2）
        Dim parts() As String: parts = Split(fileName, "_stage" & stage & "_copy")
        baseKey = parts(0)
        rowIndex = CInt(parts(1)) + 1 ' 副本1对应行2（原始行1）
    Else
        ' 原始文件：基础键=文件名，行索引1
        baseKey = fileName
        rowIndex = 1
    End If
    
    ' 检查键是否存在，行索引是否有效
    If dict.Exists(baseKey) Then
        Dim dataColl As Collection: Set dataColl = dict(baseKey)
        If rowIndex >= 1 And rowIndex <= dataColl.count Then
            ' 打开文件并更新单元格
            On Error Resume Next
            Set wb = Workbooks.Open(file.path, ReadOnly:=False)
            If Err.Number <> 0 Then
                ErrorCount = ErrorCount + 1
                Exit Sub
            End If
            arrData = dataColl(rowIndex) ' 取对应行数据
            
            ' 根据阶段执行更新和重命名
            Select Case stage
                Case 1
                    UpdateCellsStage1 wb.Sheets(1), arrData ' 阶段1更新规则（Sheet1）
                    newName = arrData(11) & ".xlsx" ' 阶段1用K列（第11列）重命名
                Case 2
                    UpdateCellsStage2 wb.Sheets(1), arrData ' 阶段2更新规则（Sheet2）
                    newName = arrData(8) & ".xlsx" ' 阶段2用H列（第8列）重命名（用户要求）
            End Select
            
            ' 添加到当前阶段的重命名映射
            renameMap(file.path) = file.parentFolder.path & "\" & newName
            
            ' 保存并关闭
            wb.Close SaveChanges:=True
            processedCount = processedCount + 1
            On Error GoTo 0
        Else
            skippedCount = skippedCount + 1 ' 行索引无效（如副本序号超过Sheet2的行数量）
        End If
    Else
        skippedCount = skippedCount + 1 ' 键不存在（文件名不在当前阶段的A列中）
    End If
End Sub

' 处理阶段2文件（遍历阶段1的输出文件）
Private Sub ProcessStage2Files(renameMap As Object, stage As Integer)
    Dim oldPath As Variant, newPath As String, file As Object
    On Error Resume Next
    For Each oldPath In renameMap.keys
        newPath = renameMap(oldPath)
        If fso.FileExists(newPath) Then
            Set file = fso.GetFile(newPath)
            ProcessSingleFile file, stage ' 处理阶段1的输出文件
        End If
    Next
    On Error GoTo 0
End Sub

' 阶段1单元格更新规则（根据Sheet1数据）
Private Sub UpdateCellsStage1(ws As Worksheet, arrData As Variant)
    With ws
        .Range("B2").value = arrData(2) ' B列→B2
        .Range("E2").value = arrData(3) ' C列→E2
        .Range("B3").value = arrData(4) ' D列→B3
        .Range("E3").value = arrData(5) ' E列→E3
        .Range("B4").value = arrData(6) ' F列→B4
        .Range("B5").value = arrData(7) ' G列→B5
        .Range("E4").value = arrData(8) ' H列→E4
        .Range("E5").value = arrData(9) ' I列→E5
        .Range("E6").value = arrData(10) ' J列→E6
    End With
End Sub

' 阶段2单元格更新规则（根据Sheet2数据）
Private Sub UpdateCellsStage2(ws As Worksheet, arrData As Variant)
    With ws
        .Range("B2").value = arrData(2) ' B列→B2
        .Range("D4").value = arrData(3) ' C列→D4
        .Range("B3").value = arrData(4) ' D列→B3
        .Range("D3").value = arrData(5) ' E列→D3
        .Range("B4").value = arrData(6) ' F列→B4
    End With
End Sub

' 批量重命名（阶段1）
Private Sub BatchRenameStage1()
    On Error Resume Next
    Dim oldPath As Variant  ' 声明循环变量（关键修正）
    For Each oldPath In renameMapStage1.keys
        Name oldPath As renameMapStage1(oldPath)
        If Err.Number <> 0 Then
            ErrorCount = ErrorCount + 1
            Err.Clear
        End If
    Next
End Sub

' 批量重命名（阶段2）
Private Sub BatchRenameStage2()
    On Error Resume Next
    Dim oldPath As Variant ' 声明循环变量（避免编译错误）
    For Each oldPath In renameMapStage2.keys
        Name oldPath As renameMapStage2(oldPath)
        If Err.Number <> 0 Then
            ErrorCount = ErrorCount + 1
            Err.Clear
        End If
    Next
End Sub

' 辅助函数：获取文件名（不含后缀）
Private Function GetBaseName(fullName As String) As String
    GetBaseName = Left(fullName, InStrRev(fullName, ".") - 1)
End Function

' 辅助函数：判断是否为有效Excel文件（排除对应表和临时文件）
Private Function IsValidExcelFile(file As Object) As Boolean
    Dim fileName As String: fileName = LCase(file.name)
    IsValidExcelFile = (Right(fileName, 5) = ".xlsx") And _
                       (fileName <> "对应表.xlsx") And _
                       (Left(fileName, 2) <> "~$")
End Function

' 辅助函数：选择根目录
Private Function SelectRootFolder() As Object
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "选择根目录"
        If .Show Then
            Set SelectRootFolder = fso.GetFolder(.SelectedItems(1))
        End If
    End With
End Function

' 辅助函数：查找对应表（根目录下）
Private Function FindMappingTable(rootPath As String) As String
    Dim tablePath As String: tablePath = rootPath & "\对应表.xlsx"
    FindMappingTable = IIf(fso.FileExists(tablePath), tablePath, "")
End Function

' 辅助函数：显示处理结果
Private Sub ShowSummary()
    MsgBox "处理完成！" & vbCrLf & _
           "成功处理: " & processedCount & " 个文件" & vbCrLf & _
           "跳过文件: " & skippedCount & " 个文件" & vbCrLf & _
           "错误文件: " & ErrorCount & " 个文件" & vbCrLf & _
           "耗时: " & Format(Timer - startTime, "0.00") & " 秒", vbInformation, "结果"
End Sub

