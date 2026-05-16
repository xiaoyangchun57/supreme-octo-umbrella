Attribute VB_Name = "A整合文件夹及成图合并"
Option Explicit

' 模块9的全局变量
Dim fsoMerge As Object
Dim logText As String
Dim gRootFolder As Object

' ========================= 主流程 =========================
Sub MainOperation()
    Dim rootPath As String
    
    ' 先执行模块8的文件整理
    rootPath = OrganizeFilesWithLog()
    If rootPath = "" Then Exit Sub
    
    ' 再执行模块9的文件合并
    Call MergeFiles(rootPath)
End Sub

' ======================== 模块8代码 ========================
Function OrganizeFilesWithLog() As String
    Dim fso As Object, rootPath As String, logFile As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 选择根目录
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "请选择要处理的根目录（所有文件将在此目录下）："
        If .Show <> -1 Then
            OrganizeFilesWithLog = ""
            Exit Function
        End If
        rootPath = .SelectedItems(1)
    End With
    
    ' 创建日志文件
    Set logFile = fso.CreateTextFile(rootPath & "\整理日志.txt", True)
    LogAction logFile, "===== 开始整理 " & Now & " =====" ' 调用LogAction
    
    ' 处理文件夹
    ProcessFolders rootPath, fso, logFile
    
    ' 结束日志
    LogAction logFile, "===== 整理完成 " & Now & " ====="
    logFile.Close
    
    OrganizeFilesWithLog = rootPath
End Function

' 添加LogAction的定义（用于写日志到文件）
Sub LogAction(logFile As Object, message As String)
    If Not logFile Is Nothing Then
        logFile.WriteLine message
    End If
End Sub

Sub ProcessFolders(folderPath As String, fso As Object, logFile As Object)
    Dim folder As Object, subFolder As Object
    Dim wb As Workbook, ws As Worksheet
    Dim mainFolderA1 As String, subFolderA2 As String
    Dim mainPathA1 As String, subPathA2 As String
    Dim colIndex As Integer, lastCol As Integer
    Dim lastRow As Long, row As Long
    
    Set folder = fso.GetFolder(folderPath)
    
    On Error Resume Next ' 忽略文件占用等错误
    
    ' 检查当前文件夹是否存在`对应表.XLSX`配置文件
    If fso.FileExists(folderPath & "\对应表.XLSX") Then
        ' 打开对应表（只读模式）
        Set wb = Workbooks.Open(folderPath & "\对应表.XLSX", ReadOnly:=True)
        
        ' === 修改点：同时处理Sheet4和Sheet5 ===
        Dim sheetNumbers
        sheetNumbers = Array(4, 5) ' 同时处理Sheet4和Sheet5
        
        Dim sheetNum As Variant
        For Each sheetNum In sheetNumbers
            On Error Resume Next
            Set ws = wb.Sheets(sheetNum)
            On Error GoTo 0
            
            If ws Is Nothing Then
                LogAction logFile, "警告：未找到Sheet" & sheetNum & "，跳过"
                GoTo NextSheet
            End If
            
            ' ==================== 1. 读取并创建第一、第二层级文件夹 ====================
            mainFolderA1 = Trim(ws.Cells(1, 1).value) ' 第一层级名称
            subFolderA2 = Trim(ws.Cells(2, 1).value) ' 第二层级名称
            
            ' 检查有效性
            If mainFolderA1 = "" Or subFolderA2 = "" Then
                LogAction logFile, "警告：`对应表.Sheet" & sheetNum & "`中A1（第一层级）或A2（第二层级）为空，跳过当前工作表"
                GoTo NextSheet
            End If
            
            ' 创建第一层级文件夹
            mainPathA1 = folderPath & "\" & mainFolderA1
            If Not fso.FolderExists(mainPathA1) Then
                fso.CreateFolder mainPathA1
                LogAction logFile, "创建第一层级文件夹：" & mainPathA1
            End If
            
            ' 创建第二层级文件夹
            subPathA2 = mainPathA1 & "\" & subFolderA2
            If Not fso.FolderExists(subPathA2) Then
                fso.CreateFolder subPathA2
                LogAction logFile, "创建第二层级文件夹：" & subPathA2
            End If
            
            ' ==================== 2. 创建第三层级文件夹（每列的名称） ====================
            lastCol = ws.Cells(3, ws.Columns.count).End(xlToLeft).Column ' 获取有效列数
            
            ' 遍历每列
            For colIndex = 1 To lastCol
                Dim subFolderName As String
                subFolderName = Trim(ws.Cells(3, colIndex).value) ' 第三层级名称
                
                If subFolderName = "" Then
                    LogAction logFile, "警告：第" & colIndex & "列（第三层级名称）为空，跳过"
                    GoTo NextColumn
                End If
                
                ' 创建第三层级文件夹
                Dim subFolderPath As String
                subFolderPath = subPathA2 & "\" & subFolderName
                If Not fso.FolderExists(subFolderPath) Then
                    fso.CreateFolder subFolderPath
                    LogAction logFile, "创建第三层级文件夹：" & subFolderPath
                End If
                
                ' ==================== 3. 收集并复制文件 ====================
                Dim mappingDict As Object
                Set mappingDict = CreateObject("Scripting.Dictionary")
                lastRow = ws.Cells(ws.Rows.count, colIndex).End(xlUp).row
                
                ' 收集文件名
                If lastRow >= 4 Then
                    For row = 4 To lastRow
                        Dim fileName As String
                        fileName = Trim(ws.Cells(row, colIndex).value)
                        If fileName <> "" Then
                            mappingDict.Add fileName, fileName
                        End If
                    Next row
                End If
                
                ' 复制文件
                If mappingDict.count > 0 Then
                    SearchAndCopyFiles folderPath, subFolderPath, mappingDict, fso, logFile
                Else
                    LogAction logFile, "警告：第" & colIndex & "列没有文件列表，跳过"
                End If
                
NextColumn:
            Next colIndex
            
NextSheet:
        Next sheetNum
        ' ==================== 修改结束 ====================
        
        wb.Close False ' 关闭工作簿（不保存）
    End If
    
    ' 递归处理子文件夹
    For Each subFolder In folder.subFolders
        ProcessFolders subFolder.path, fso, logFile
    Next subFolder
    
    On Error GoTo 0 ' 恢复错误处理
End Sub
Sub SearchAndCopyFiles(rootPath As String, destFolderPath As String, dict As Object, fso As Object, logFile As Object)
    Dim folder As Object, subFolder As Object, file As Object
    Dim fileName As String
    
    Set folder = fso.GetFolder(rootPath)
    
    ' 处理当前文件夹中的文件
    For Each file In folder.files
        ' 跳过日志文件和对应表（避免误操作）
        If file.name <> "操作日志.txt" And file.name <> "对应表.XLSX" Then
            ' 获取文件名（不含扩展名）—— 用于匹配`对应表`中的文件名称
            fileName = fso.GetBaseName(file.name)
            ' 检查是否匹配当前列的文件名称（第四行及以下）
            If dict.Exists(fileName) Then
                Dim destFilePath As String
                destFilePath = destFolderPath & "\" & file.name ' 保留原文件扩展名
                ' 避免覆盖已存在的文件（防止数据丢失）
                If Not fso.FileExists(destFilePath) Then
                    file.Copy destFilePath ' 复制文件（保留原文件）
                    LogAction logFile, "复制文件：" & file.path & " → " & destFilePath
                Else
                    LogAction logFile, "跳过：" & file.name & " 已存在于 " & destFolderPath
                End If
            End If
        End If
    Next file
    
    ' 递归处理子文件夹（搜索深层文件夹中的文件）
    For Each subFolder In folder.subFolders
        ' 避免无限循环（跳过目标文件夹）
        If subFolder.path <> destFolderPath Then
            SearchAndCopyFiles subFolder.path, destFolderPath, dict, fso, logFile
        End If
    Next subFolder
End Sub



' 其余模块8函数保持不变（SearchAndCopyFiles和LogAction）...

' ======================== 模块9代码 ========================
Sub MergeFiles(rootPath As String)
    Set fsoMerge = CreateObject("Scripting.FileSystemObject")
    logText = "合并日志：" & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbCrLf
    
    If rootPath <> "" Then
        Set gRootFolder = fsoMerge.GetFolder(rootPath)
    Else
        MsgBox "根目录路径无效", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ProcessFoldersForMerge gRootFolder
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    CreateLogFile
    MsgBox "文件合并完成：" & vbCrLf & "详细日志已保存到桌面", vbInformation
End Sub

Sub ProcessFoldersForMerge(currentFolder As Object)
    Dim subFolder As Object
    Dim file As Object
    
    ' 查找并处理对应表
    For Each file In currentFolder.files
        If LCase(file.name) = "对应表.xlsx" Then
            ProcessMasterFile file
            Exit For
        End If
    Next file
    
    ' 递归处理子文件夹
    For Each subFolder In currentFolder.subFolders
        ProcessFoldersForMerge subFolder
    Next subFolder
End Sub

Private Sub ProcessMasterFile(masterFile As Object)
    Dim wbMaster As Workbook
    Dim wsMaster As Worksheet
    Dim col As Range
    Dim headerCell As Range
    Dim lastCol As Long
    Dim baseFolderName As String
    Dim newFolderName As String
    Dim newFolderPath As String
    
    On Error GoTo ErrorHandler
    
    Set wbMaster = Workbooks.Open(masterFile.path, ReadOnly:=True)
    
    ' === 修改点：同时处理Sheet4和Sheet5 ===
    Dim sheetNumbers
    sheetNumbers = Array(4, 5)
    
    Dim sheetNum As Variant
    For Each sheetNum In sheetNumbers
        On Error Resume Next
        Set wsMaster = wbMaster.Sheets(sheetNum)
        On Error GoTo ErrorHandler
        
        If wsMaster Is Nothing Then
            logText = logText & vbCrLf & "源文件 " & masterFile.name & " 缺少Sheet" & sheetNum
            GoTo NextSheet
        End If
        
        ' 读取配置
        baseFolderName = Trim(wsMaster.Range("A1").value)
        newFolderName = Trim(wsMaster.Range("B2").value)
        
        If baseFolderName = "" Or newFolderName = "" Then
            logText = logText & vbCrLf & "源文件 " & masterFile.name & " Sheet" & sheetNum & " 中A1或B2为空"
            GoTo NextSheet
        End If
        
        ' 创建目标文件夹
        newFolderPath = gRootFolder.path & "\" & baseFolderName
        If Not fsoMerge.FolderExists(newFolderPath) Then
            logText = logText & vbCrLf & "错误：找不到文件夹 " & newFolderPath
            GoTo NextSheet
        End If
        
        newFolderPath = newFolderPath & "\" & newFolderName
        If Not fsoMerge.FolderExists(newFolderPath) Then
            fsoMerge.CreateFolder newFolderPath
            logText = logText & vbCrLf & "已创建文件夹：" & newFolderPath
        End If
        
        ' 处理列数据
        lastCol = wsMaster.Cells(3, wsMaster.Columns.count).End(xlToLeft).Column
        For Each col In wsMaster.Range(wsMaster.Cells(3, 1), wsMaster.Cells(3, lastCol))
            Set headerCell = col.Cells(1, 1)
            If Len(Trim(headerCell.value)) > 0 Then
                ProcessColumn headerCell, newFolderPath
            End If
        Next
        
NextSheet:
        Set wsMaster = Nothing
    Next sheetNum
    ' ==================== 修改结束 ====================
    
CloseMaster:
    wbMaster.Close False
    Exit Sub
    
ErrorHandler:
    logText = logText & vbCrLf & "处理文件出错：" & masterFile.path & vbCrLf & "错误信息：" & Err.Description
    Resume CloseMaster
End Sub

' 其余模块9函数保持不变（ProcessColumn, CreateWorkbook等）...
Private Sub ProcessColumn(headerCell As Range, targetFolderPath As String)
    Dim targetFileName As String
    Dim wbTarget As Workbook
    Dim dataRange As Range
    Dim cell As Range
    Dim fileCounter As Long
    Dim headerParts() As String
    
    ' 分割标题（取_前部分）
    headerParts = Split(headerCell.value, "_")
    If UBound(headerParts) >= 0 Then
        targetFileName = CleanFileName(headerParts(0))
    Else
        targetFileName = CleanFileName(headerCell.value)
    End If
    
    If Len(targetFileName) = 0 Then Exit Sub
    
    ' 创建目标工作簿
    Set wbTarget = CreateWorkbook(targetFolderPath, targetFileName)
    
    ' 获取数据范围（第四行开始）
    Set dataRange = headerCell.Parent.Range(headerCell.Offset(1, 0), _
                     headerCell.Parent.Cells(headerCell.Parent.Rows.count, headerCell.Column).End(xlUp))
    
    ' 处理每个数据项
    For Each cell In dataRange
        If Len(Trim(cell.value)) > 0 Then
            Dim fileNameParts() As String
            Dim searchFileName As String
            
            ' 分割文件名（取_后部分）
            fileNameParts = Split(cell.value, "_")
            If UBound(fileNameParts) >= 1 Then
                searchFileName = Trim(fileNameParts(1))
            Else
                searchFileName = Trim(cell.value)
            End If
            
            If Len(searchFileName) > 0 Then
                fileCounter = fileCounter + 1
                FindAndCopyFile searchFileName, gRootFolder, wbTarget
            End If
        End If
    Next
    
    ' 自动删除Sheet1
    Application.DisplayAlerts = False
    On Error Resume Next
    wbTarget.Sheets("Sheet1").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' 保存工作簿
    wbTarget.Close SaveChanges:=True
    
    logText = logText & vbCrLf & "已创建：" & targetFileName & ".xlsx" & _
              "（包含" & fileCounter & "个工作表）"
End Sub

Private Function CreateWorkbook(folderPath As String, fileName As String) As Workbook
    Dim fullPath As String
    ' ==== 修复点：将 fso 替换为 fsoMerge ====
    fullPath = fsoMerge.BuildPath(folderPath, fileName & ".xlsx")
    
    ' 删除已存在文件
    ' ==== 修复点：将 fso 替换为 fsoMerge ====
    If fsoMerge.FileExists(fullPath) Then
        logText = logText & vbCrLf & "删除旧文件：" & fullPath
        fsoMerge.DeleteFile fullPath
    End If
    
    ' 创建新工作簿
    Set CreateWorkbook = Workbooks.Add
    Application.DisplayAlerts = False
    CreateWorkbook.SaveAs fullPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
End Function

Private Sub FindAndCopyFile(fileName As String, searchFolder As Object, targetWorkbook As Workbook)
    Dim foundFile As Object
    Dim searchPath As String
    
    ' ==== 修复点：将 fso 替换为 fsoMerge ====
    searchPath = fsoMerge.BuildPath(searchFolder.path, fileName & ".xlsx")
    
    ' 首先检查当前目录
    ' ==== 修复点：将 fso 替换为 fsoMerge ====
    If fsoMerge.FileExists(searchPath) Then
        CopyWorksheet fsoMerge.GetFile(searchPath), targetWorkbook, fileName
        Exit Sub
    End If
    
    ' 递归搜索子文件夹
    For Each foundFile In searchFolder.subFolders
        SearchSubFolders foundFile, fileName, targetWorkbook
    Next
End Sub

Private Sub SearchSubFolders(folder As Object, fileName As String, targetWorkbook As Workbook)
    Dim subFile As Object
    Dim searchPath As String
    
    ' ==== 修复点：将 fso 替换为 fsoMerge ====
    searchPath = fsoMerge.BuildPath(folder.path, fileName & ".xlsx")
    
    ' ==== 修复点：将 fso 替换为 fsoMerge ====
    If fsoMerge.FileExists(searchPath) Then
        CopyWorksheet fsoMerge.GetFile(searchPath), targetWorkbook, fileName
        Exit Sub
    End If
    
    For Each subFile In folder.subFolders
        SearchSubFolders subFile, fileName, targetWorkbook
    Next
End Sub

Private Sub CopyWorksheet(sourceFile As Object, targetWorkbook As Workbook, sheetName As String)
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim cleanedName As String
    
    On Error GoTo CopyError
    
    ' 打开源文件
    Set wbSource = Workbooks.Open(sourceFile.path, ReadOnly:=True)
    Set wsSource = wbSource.Sheets(1)
    
    ' 复制工作表
    wsSource.Copy After:=targetWorkbook.Sheets(targetWorkbook.Sheets.count)
    Set wsTarget = targetWorkbook.Sheets(targetWorkbook.Sheets.count)
    
    ' 清理并尝试重命名工作表
    cleanedName = CleanSheetName(sheetName)
    On Error Resume Next
    wsTarget.name = cleanedName
    If Err.Number <> 0 Then
        logText = logText & vbCrLf & "警告：工作表名称 """ & cleanedName & """ 已存在，已保留默认名称：" & wsTarget.name
        Err.Clear
    End If
    On Error GoTo CopyError
    
    ' 关闭源文件
    wbSource.Close False
    Exit Sub
    
CopyError:
    logText = logText & vbCrLf & "复制失败：" & sourceFile.name & " -> " & sheetName & " | 错误描述：" & Err.Description
    If Not wbSource Is Nothing Then wbSource.Close False
End Sub

Private Function CleanSheetName(inputName As String) As String
    Dim badChars As Variant
    Dim i As Integer
    badChars = Array("\", "/", "?", "*", "[", "]", ":")
    
    CleanSheetName = inputName
    For i = LBound(badChars) To UBound(badChars)
        CleanSheetName = Replace(CleanSheetName, badChars(i), "")
    Next
End Function

Private Function CleanFileName(inputName As String) As String
    Dim badChars As Variant
    Dim i As Integer
    badChars = Array("\", "/", "?", "*", ":", "|", "<", ">", """")
    
    CleanFileName = inputName
    For i = LBound(badChars) To UBound(badChars)
        CleanFileName = Replace(CleanFileName, badChars(i), "_")
    Next
End Function

Private Sub CreateLogFile()
    Dim logPath As String
    logPath = Environ("USERPROFILE") & "\Desktop\文件合并日志.txt"
    
    Open logPath For Output As #1
    Print #1, logText
    Close #1
End Sub

