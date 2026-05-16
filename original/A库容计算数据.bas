Attribute VB_Name = "A库容计算数据"

Option Explicit

' 共享变量声明
Dim fso As Object, logFile As Object
Dim rootPath As String
Dim templateName As String
Dim logPath As String
Dim module8Completed As Boolean

' ==================== 主执行流程 ====================
Sub RunBothMacros()
    On Error GoTo ErrorHandler
    Dim startTime As Double
    startTime = Timer
    
    ' 初始化文件系统对象
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 用户选择项目目录
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "请选择项目目录"
        If .Show Then
            rootPath = .SelectedItems(1)
            If Right(rootPath, 1) <> "\" Then rootPath = rootPath & "\"
        Else
            Exit Sub
        End If
    End With
    
    ' 创建运行日志文件
    logPath = rootPath & Format(Now, "yyyymmdd_hhmmss") & "_运行日志.txt"
    Set logFile = fso.CreateTextFile(logPath, True)
    
    ' 运行模块8
    LogAction "===== 模块8开始执行 ====="
    MasterWorkflow
    LogAction ">> 模块8执行完成"
    module8Completed = True
    
    ' 自动运行模块9
    LogAction vbCrLf & "===== 模块9开始执行 ====="
    ProcessFolders_FixMerge rootPath
    LogAction ">> 模块9执行完成"
    
Cleanup:
    ' 关闭日志文件
    If Not logFile Is Nothing Then
        LogAction vbCrLf & "总耗时: " & Format(Timer - startTime, "0.00") & "秒"
        logFile.Close
    End If
    Set logFile = Nothing
    Set fso = Nothing
    
    ' 最终提示
    If module8Completed Then
        MsgBox "所有操作已完成!" & vbCrLf & "日志路径: " & logPath, vbInformation
    Else
        MsgBox "处理过程中发生错误，请查看日志: " & logPath, vbExclamation
    End If
    Exit Sub
    
ErrorHandler:
    LogAction "错误: " & Err.Description & " (代码:" & Err.Number & ")"
    Resume Cleanup
End Sub

' ====================================================
'               模块8的核心代码 (库容计算)
' ====================================================
Sub MasterWorkflow()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    ' 定义模板名称
    templateName = "面积计算模板.xlsm"
    
    LogAction "项目目录: " & rootPath
    LogAction "使用模板: " & templateName
    
    ' 步骤1: 模板复制
    LogAction vbCrLf & ">> 开始模板复制操作"
    ProcessFolder rootPath, "CopyTemplate"
    LogAction ">> 模板复制完成"
    
    ' 步骤2: 数据同步
    LogAction vbCrLf & ">> 开始数据同步操作"
    ProcessFolder rootPath, "DataSync"
    LogAction ">> 数据同步完成"
    
    ' 步骤3: 文件归档
    LogAction vbCrLf & ">> 开始文件归档操作"
    ProcessFolder rootPath, "OrganizeFiles"
    LogAction ">> 文件归档完成"
    
    ' 步骤4: 文件标记
    LogAction vbCrLf & ">> 开始文件标记操作"
    ProcessFolder rootPath, "FolderMark"
    LogAction ">> 文件标记完成"
    
    Exit Sub
    
ErrorHandler:
    LogAction "模块8错误: " & Err.Description & " (代码:" & Err.Number & ")"
    Err.Raise Err.Number
End Sub

' 递归处理文件夹
Sub ProcessFolder(folderPath As String, operationType As String)
    On Error Resume Next
    Dim folder As Object, subFolder As Object
    Set folder = fso.GetFolder(folderPath)
    If Err.Number <> 0 Then
        LogAction "错误: 无法访问文件夹 " & folderPath
        Exit Sub
    End If
    
    Select Case operationType
        Case "CopyTemplate": Macro3_CopyTemplate folder
        Case "DataSync": Macro1_DataSync folder
        Case "OrganizeFiles": Macro4_OrganizeFiles folder
        Case "FolderMark": Macro2_FolderMark folder
    End Select
    
    ' 递归处理子文件夹
    For Each subFolder In folder.subFolders
        ProcessFolder subFolder.path, operationType
    Next
End Sub

' 模板复制功能
Sub Macro3_CopyTemplate(folder As Object)
    Dim xlsxFiles As New Collection
    Dim file As Object, xlsxFile As Object
    Dim templatePath As String
    
    ' 在当前文件夹中查找模板
    templatePath = folder.path & "\" & templateName
    If Not fso.FileExists(templatePath) Then
        ' 在根目录查找模板
        templatePath = rootPath & templateName
        If Not fso.FileExists(templatePath) Then
            LogAction "警告: " & folder.path & " 缺少模板文件"
            Exit Sub
        End If
    End If
    
   For Each file In folder.files
    If LCase(fso.GetExtensionName(file.name)) = "xlsx" Then
        Dim fileNameLower As String
        fileNameLower = LCase(file.name)
        ' 检查文件名是否包含"横断面"或B/Q/K/G/J
        If InStr(1, fileNameLower, "横断面", vbTextCompare) > 0 Or _
           InStr(1, fileNameLower, "b") > 0 Or _
           InStr(1, fileNameLower, "q") > 0 Or _
           InStr(1, fileNameLower, "k") > 0 Or _
           InStr(1, fileNameLower, "g") > 0 Or _
           InStr(1, fileNameLower, "j") > 0 Then
            xlsxFiles.Add file
        End If
    End If
Next
    
    ' 复制逻辑
    Dim copyCount As Long: copyCount = 0
    For Each xlsxFile In xlsxFiles
        Dim newName As String
        newName = fso.GetBaseName(xlsxFile.name) & ".xlsm"
        Dim newFilePath As String
        newFilePath = folder.path & "\" & newName
        
        ' 目标文件不存在时才复制
        If Not fso.FileExists(newFilePath) Then
            fso.CopyFile templatePath, newFilePath
            LogAction "复制: " & templateName & " 到 " & newName & " (位置: " & folder.path & ")"
            copyCount = copyCount + 1
        Else
            LogAction "跳过: " & newName & " 已存在"
        End If
    Next
    
    LogAction "模板复制: 完成 " & copyCount & " 个复制 (位置: " & folder.path & ")"
End Sub

' 文件归档处理
Sub Macro4_OrganizeFiles(folder As Object)
    Dim mappingFile As String, wb As Workbook, ws As Worksheet
    Dim mappingDict As Object, file As Object, cell As Range
    Dim colIndex As Integer, destFolderPath As String, fileName As String
    
    ' 检查对应表
    mappingFile = folder.path & "\对应表.xlsx"
    If Not fso.FileExists(mappingFile) Then Exit Sub
    
    ' 递归收集所有文件
    Dim allFiles As New Collection
    CollectFilesRecursive folder, allFiles
    
    Set mappingDict = CreateObject("Scripting.Dictionary")
    Set wb = Workbooks.Open(mappingFile, ReadOnly:=True)
    
    ' 使用Sheet6作为对应表
    On Error Resume Next
    Set ws = wb.Sheets("Sheet6")
    If ws Is Nothing Then
        LogAction "错误: " & mappingFile & " 缺少Sheet6"
        wb.Close False
        Exit Sub
    End If
    
    ' 读取Sheet6对应表
    For colIndex = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        Dim targetFolderName As String
        targetFolderName = Trim(ws.Cells(1, colIndex).value)
        If targetFolderName <> "" Then
            ' 创建目标文件夹
            destFolderPath = folder.path & "\" & targetFolderName
            If Not fso.FolderExists(destFolderPath) Then
                fso.CreateFolder destFolderPath
                LogAction "创建文件夹: " & destFolderPath
            End If
            
            ' 收集映射关系
            mappingDict.RemoveAll
            For Each cell In ws.Range(ws.Cells(2, colIndex), ws.Cells(ws.Rows.count, colIndex).End(xlUp))
                If Not IsEmpty(cell.value) Then
                    mappingDict.Add fso.GetBaseName(CStr(cell.value)), True
                End If
            Next cell
            
            ' 移动匹配的文件到目标文件夹
            For Each file In allFiles
                fileName = fso.GetBaseName(file.name)
                If mappingDict.Exists(fileName) Then
                    Dim destPath As String
                    destPath = destFolderPath & "\" & file.name
                    
                    ' 确保源文件不在目标文件夹中
                    If StrComp(file.parentFolder.path, destFolderPath, vbTextCompare) <> 0 Then
                        If Not fso.FileExists(destPath) Or OverwritePrompt(file.name) Then
                            file.Copy destPath
                            LogAction "归档文件: " & file.name & " -> " & targetFolderName
                        End If
                    End If
                End If
            Next file
        End If
    Next colIndex
    
    wb.Close False
End Sub

' 递归收集文件
Sub CollectFilesRecursive(folder As Object, ByRef allFiles As Collection)
    Dim file As Object, subFolder As Object
    
    ' 添加当前文件夹的文件
    For Each file In folder.files
        ' 排除系统文件
        If file.name <> "对应表.xlsx" And file.name <> templateName Then
            allFiles.Add file
        End If
    Next
    
    ' 递归处理子文件夹
    For Each subFolder In folder.subFolders
        CollectFilesRecursive subFolder, allFiles
    Next
End Sub

' 文件覆盖确认
Function OverwritePrompt(fileName As String) As Boolean
    ' 直接返回True自动确认覆盖
    OverwritePrompt = True
End Function

Sub Macro1_DataSync(folder As Object)
    Dim xlsxFiles As New Collection, xlsmFiles As New Collection
    Dim file As Object, xlsxFile As Object, xlsmFile As Object
    
    ' 收集文件 - 修改为包含所有符合条件的xlsx文件
    For Each file In folder.files
        If LCase(fso.GetExtensionName(file.name)) = "xlsx" Then
            Dim fileNameLower As String
            fileNameLower = LCase(file.name)
            ' 检查文件名是否包含"横断面"或B/Q/K/G/J
            If InStr(1, fileNameLower, "横断面", vbTextCompare) > 0 Or _
               InStr(1, fileNameLower, "b") > 0 Or _
               InStr(1, fileNameLower, "q") > 0 Or _
               InStr(1, fileNameLower, "k") > 0 Or _
               InStr(1, fileNameLower, "g") > 0 Or _
               InStr(1, fileNameLower, "j") > 0 Then
                xlsxFiles.Add file
            End If
        ElseIf LCase(fso.GetExtensionName(file.name)) = "xlsm" Then
            ' 排除包含"模板"的xlsm文件
            If InStr(1, file.name, "模板", vbTextCompare) = 0 Then
                xlsmFiles.Add file
            End If
        End If
    Next
    
    ' 数据同步逻辑
    For Each xlsxFile In xlsxFiles
        Set xlsmFile = FindMatchingXlsm(xlsmFiles, xlsxFile.name)
        If Not xlsmFile Is Nothing Then
            CopyData xlsxFile.path, xlsmFile.path
        Else
            LogAction "警告: 未找到与 " & xlsxFile.name & " 匹配的xlsm文件"
        End If
    Next
End Sub

' 文件标记功能
Sub Macro2_FolderMark(folder As Object)
    Dim file As Object, wb As Workbook, ws As Worksheet
    Dim folderName As String, markText As String
    Dim qiaoPos As Integer
    
    ' 获取文件夹名称
    folderName = folder.name
    
    ' 查找"桥"的位置
    qiaoPos = InStr(1, folderName, "桥", vbTextCompare)
    
    ' 提取标记文本
    If qiaoPos > 0 Then
        markText = Trim(Mid(folderName, qiaoPos + 1))  ' 取"桥"之后的内容
    Else
        markText = folderName  ' 如果没有"桥"，使用全名
    End If
    
    For Each file In folder.files
        If LCase(fso.GetExtensionName(file.name)) = "xlsm" Then
            ' 排除模板文件
            If InStr(1, file.name, "模板", vbTextCompare) = 0 Then
                On Error Resume Next
                Set wb = Workbooks.Open(file.path)
                Set ws = wb.Sheets("面积计算")
                
                If Err.Number = 0 Then
                    ' 写入提取的标记文本
                    ws.Range("A2").value = markText
                    ws.Range("A2").NumberFormat = "@"
                    wb.Close True
                    LogAction "标记: " & file.name & " <- " & markText & " (原始文件夹: " & folderName & ")"
                Else
                    LogAction "错误: " & file.name & " 缺少[面积计算]工作表"
                    If Not wb Is Nothing Then wb.Close False
                End If
                On Error GoTo 0  ' 重置错误处理
            End If
        End If
    Next
End Sub

' 辅助函数
Function FindMatchingXlsm(files As Collection, xlsxName As String) As Object
    Dim baseName As String, file As Object
    baseName = Left(xlsxName, InStrRev(xlsxName, ".") - 1)
    
    For Each file In files
        If StrComp(Left(file.name, InStrRev(file.name, ".") - 1), baseName, vbTextCompare) = 0 Then
            Set FindMatchingXlsm = file
            Exit Function
        End If
    Next
End Function

' 数据复制函数
Sub CopyData(xlsxPath As String, xlsmPath As String)
    On Error GoTo ErrorHandler
    Dim wbSource As Workbook, wbTarget As Workbook
    Dim wsSource As Worksheet, wsTarget As Worksheet
    
    Set wbSource = Workbooks.Open(xlsxPath, ReadOnly:=True)
    Set wsSource = wbSource.Sheets(1)
    
    Set wbTarget = Workbooks.Open(xlsmPath)
    Set wsTarget = wbTarget.Sheets("大断面")
    
    ' 数据复制
    wsSource.Range("C13:D200").Copy
    wsTarget.Range("A2:B200").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
    wbTarget.Close True
    wbSource.Close False
    
    LogAction "同步: " & fso.GetFileName(xlsxPath) & " 到 " & fso.GetFileName(xlsmPath)
    Exit Sub
    
ErrorHandler:
    LogAction "同步失败: " & fso.GetFileName(xlsmPath) & " - " & Err.Description
    If Not wbTarget Is Nothing Then wbTarget.Close False
    If Not wbSource Is Nothing Then wbSource.Close False
End Sub

' ====================================================
'               模块9的核心代码 (库容计算2)
' ====================================================
Function WorksheetExists(sheetName As String, wb As Workbook) As Boolean
    On Error Resume Next
    WorksheetExists = Not wb.Sheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

' 模块9主程序（自动使用模块8的根目录）
Sub ProcessFolders_FixMerge(mainFolder As String)
    On Error GoTo ErrorHandler
    Dim mainWb As Workbook ' 库容数据.xlsx的工作簿
    Dim templateWS As Worksheet ' 模板Sheet1
    Dim subFolderSheet As Worksheet ' 子文件夹对应Sheet
    Dim subFolder As Object ' 子文件夹
    Dim file As Object ' 文件
    Dim dict As Object ' 映射字典
    Dim fileName As String ' 文件名（不含扩展名）
    Dim fileExt As String ' 文件扩展名
    Dim lastRow As Long ' 最后行号
    
    LogAction "模块9目标目录: " & mainFolder
    
    ' 检查库容数据是否存在
    If Dir(mainFolder & "\库容数据.xlsx") = "" Then
        LogAction "错误: 库容数据.xlsx 不存在"
        MsgBox "库容数据夹中未找到'库容数据.xlsx'", vbCritical
        Exit Sub
    End If
    
    ' 打开库容数据并获取模板
    Set mainWb = Workbooks.Open(mainFolder & "\库容数据.xlsx")
    On Error Resume Next
    Set templateWS = mainWb.Sheets("Sheet1")
    On Error GoTo 0
    
    If templateWS Is Nothing Then
        LogAction "错误: 缺少Sheet1模板"
        MsgBox "库容数据中未找到'Sheet1'模板", vbExclamation
        mainWb.Close False
        Exit Sub
    End If
    
    ' 处理每个子文件夹
    For Each subFolder In fso.GetFolder(mainFolder).subFolders
        Dim subFolderName As String: subFolderName = subFolder.name
        LogAction "处理子文件夹: " & subFolderName
        
        ' 检查Sheet是否已存在
        If WorksheetExists(subFolderName, mainWb) Then
            LogAction "提示: " & subFolderName & " 工作表已存在，正在覆盖"
            Application.DisplayAlerts = False
            mainWb.Sheets(subFolderName).Delete
            Application.DisplayAlerts = True
        End If
        
        ' 创建工作表（解除合并单元格）
        templateWS.Copy After:=mainWb.Sheets(mainWb.Sheets.count)
        Set subFolderSheet = mainWb.Sheets(mainWb.Sheets.count)
        subFolderSheet.name = subFolderName
        
        ' 解除合并单元格并清除内容
        With subFolderSheet
            .Rows("3:" & .Rows.count).UnMerge
            .Rows("3:" & .Rows.count).ClearContents
        End With
        
        ' 设置初始行
        Set dict = CreateObject("Scripting.Dictionary")
        lastRow = 2
        
        ' 处理文件夹中的文件
        For Each file In subFolder.files
            fileName = fso.GetBaseName(file.name)
            fileExt = LCase(fso.GetExtensionName(file.name))
            
            ' 只处理Excel文件
            If fileExt <> "xlsx" And fileExt <> "xlsm" Then GoTo NextFile
            
            ' 创建/更新文件条目
            If Not dict.Exists(fileName) Then
                subFolderSheet.Cells(lastRow, "A").value = fileName
                dict.Add fileName, lastRow
                lastRow = lastRow + 1
            End If
            Dim rowNum As Long: rowNum = dict(fileName)
            
            ' 处理xlsx文件（经度/纬度）
            If fileExt = "xlsx" Then
                Dim xlsWB As Workbook: Set xlsWB = Workbooks.Open(file.path, ReadOnly:=True)
                Dim i As Long: For i = 1 To xlsWB.Sheets(1).UsedRange.Rows.count
                    If xlsWB.Sheets(1).Cells(i, "B").value = "深泓点" Then
                        subFolderSheet.Cells(rowNum, "B").value = xlsWB.Sheets(1).Cells(i, "E").value ' 经度
                        subFolderSheet.Cells(rowNum, "C").value = xlsWB.Sheets(1).Cells(i, "F").value ' 纬度
                        Exit For
                    End If
                Next i
                xlsWB.Close False
            End If
            
            ' 处理xlsm文件（面积）
            If fileExt = "xlsm" Then
                Dim xlsmWB As Workbook: Set xlsmWB = Workbooks.Open(file.path, ReadOnly:=True)
                Dim areaSheet As Worksheet: Set areaSheet = Nothing
                On Error Resume Next
                Set areaSheet = xlsmWB.Sheets("面积计算")
                On Error GoTo 0
                If Not areaSheet Is Nothing Then
                    subFolderSheet.Cells(rowNum, "D").value = areaSheet.Range("B2").value ' 断面面积
                End If
                xlsmWB.Close False
            End If
            
NextFile:
        Next file
        
        ' 自动调整列宽
        subFolderSheet.Columns("A:D").AutoFit
    Next subFolder
    
    ' 保存并关闭文件
    mainWb.Save
    mainWb.Close False
    
    LogAction ">> 模块9执行完成"
    Exit Sub
    
ErrorHandler:
    LogAction "模块9错误: " & Err.Description & " (代码:" & Err.Number & ")"
    MsgBox "处理过程中发生错误: " & Err.Description, vbCritical
End Sub

' 日志记录函数
Sub LogAction(message As String)
    If Not logFile Is Nothing Then
        logFile.WriteLine Now & " - " & message
    End If
    Debug.Print Now & " - " & message
End Sub

