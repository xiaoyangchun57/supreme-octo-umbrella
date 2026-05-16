Attribute VB_Name = "A成图自动"
Option Explicit
' 需引用 Microsoft Scripting Runtime（工具->引用）
Dim fso As New FileSystemObject  ' 模块级变量，所有子程序可访问

Sub 横断面数据处理()
    Dim rootFolder As folder, logSheet As Worksheet
    Dim calcMode As Long, eventsStatus As Boolean
    Dim templateFile As file, templatePath As String
    Dim corrFilePath As String, wbCorr As Workbook, dictSheet1 As Object, dictSheet2 As Object
    Dim logRow As Long, lastLogRow As Long, destPath As String
    
    ' 优化Excel设置
    calcMode = Application.Calculation
    eventsStatus = Application.EnableEvents
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' 初始化日志表
    Set logSheet = InitLogSheet()
    
    ' 选择根目录
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .title = "请选择项目根目录"
        If .Show <> -1 Then GoTo Cleanup  ' 用户取消选择，终止
        Set rootFolder = fso.GetFolder(.SelectedItems(1))
    End With
    
    ' 检查根目录是否有"成图模板.xlsx"
    Set templateFile = Nothing
    On Error Resume Next
    Set templateFile = fso.GetFile(fso.BuildPath(rootFolder.path, "成图模板.xlsx"))
    On Error GoTo 0
    If templateFile Is Nothing Then
        MsgBox "根目录下未找到“成图模板.xlsx”文件，请添加后重试！", vbCritical, "错误"
        GoTo Cleanup
    End If
    templatePath = templateFile.path  ' 保存根目录模板路径
    
    ' 检查对应表
    corrFilePath = fso.BuildPath(rootFolder.path, "对应表.xlsx")
    Set dictSheet1 = Nothing
    Set dictSheet2 = Nothing
    
    If fso.FileExists(corrFilePath) Then
        ' 创建两个字典分别存储Sheet1和Sheet2的数据
        Set dictSheet1 = CreateObject("Scripting.Dictionary")
        Set dictSheet2 = CreateObject("Scripting.Dictionary")
        dictSheet1.CompareMode = vbTextCompare  ' 不区分大小写
        dictSheet2.CompareMode = vbTextCompare  ' 不区分大小写
        
        Set wbCorr = Workbooks.Open(corrFilePath, ReadOnly:=True)
        
        ' 填充Sheet1字典（K列→F列）
        FillDictionaryFromSheet wbCorr.Sheets(1), dictSheet1, "K"
        
        ' 填充Sheet2字典（H列→F列）
        FillDictionaryFromSheet wbCorr.Sheets(2), dictSheet2, "H"
        
        wbCorr.Close False
    End If
    
    ' 开始处理
    ProcessFolder rootFolder, logSheet, templatePath
    
    ' 处理完成后修改成图文件
    With logSheet
        lastLogRow = .Cells(.Rows.count, 1).End(xlUp).row
        For logRow = 2 To lastLogRow
            If .Cells(logRow, 5).value = "成功" Then
                destPath = .Cells(logRow, 4).value & "\" & .Cells(logRow, 2).value
                ModifyOutputFile destPath, dictSheet1, dictSheet2, .Cells(logRow, 7)
            End If
        Next logRow
    End With
    
Cleanup:
    ' 恢复Excel设置
    Application.Calculation = calcMode
    Application.EnableEvents = eventsStatus
    Application.ScreenUpdating = True
    
    If Not logSheet Is Nothing Then
        logSheet.Columns("A:G").AutoFit
        MsgBox "处理完成！共处理 " & logSheet.Range("A" & Rows.count).End(xlUp).row - 1 & " 条记录", vbInformation
    End If
    
    Set dictSheet1 = Nothing
    Set dictSheet2 = Nothing
End Sub

' 初始化日志表
Private Function InitLogSheet() As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("处理日志").Delete
    Application.DisplayAlerts = True
    Set InitLogSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    With InitLogSheet
        .name = "处理日志"
        .Range("A1:G1") = Array("源文件", "目标文件", "源路径", "目标路径", "处理状态", "时间戳", "详情")
        .Rows(1).Font.Bold = True
    End With
End Function

' 递归处理所有子文件夹
Private Sub ProcessFolder(currentFolder As folder, logSheet As Worksheet, templatePath As String)
    Dim subFolder As folder
    Dim measureFile As file
    Dim baseName As String, arrName() As String
    Dim destPath As String
    
    ' 处理当前文件夹内的源文件
    For Each measureFile In currentFolder.files
        If IsValidFile(measureFile) Then
            ' 解析文件名：获取最后一个下划线后的部分
            arrName = Split(measureFile.name, "_")
            If UBound(arrName) >= 1 Then
                baseName = arrName(UBound(arrName))
                
                ' 跳过模板文件
                If StrComp(measureFile.name, "成图模板.xlsx", vbTextCompare) = 0 Then
                    GoTo SkipCurrentFile
                End If
                
                ' 构造目标路径（源文件所在子文件夹 + 基础名）
                destPath = fso.BuildPath(currentFolder.path, baseName)
                
                ' 处理目标文件（删除旧文件→复制模板→传输数据）
                On Error Resume Next
                If fso.FileExists(destPath) Then Kill destPath  ' 删除旧文件
                fso.CopyFile templatePath, destPath  ' 复制根目录模板
                If Err.Number <> 0 Then
                    LogOperation logSheet, measureFile.path, destPath, "失败", "模板复制错误：" & Err.Description
                    Err.Clear
                    GoTo SkipCurrentFile  ' 复制失败，跳过后续处理
                End If
                On Error GoTo 0
                
                ' 传输数据（源文件→目标文件）
                TransferData measureFile.path, destPath, logSheet
            End If
        End If
SkipCurrentFile:
    Next measureFile
    
    ' 递归处理子文件夹
    For Each subFolder In currentFolder.subFolders
        ProcessFolder subFolder, logSheet, templatePath
    Next subFolder
End Sub

' 验证文件有效性
Private Function IsValidFile(fileObj As file) As Boolean
    IsValidFile = (LCase(Right(fileObj.name, 5)) = ".xlsx" Or _
                  LCase(Right(fileObj.name, 4)) = ".xls") And _
                  Not fileObj.name Like "~$*" And _
                  fileObj.Size > 0
End Function
Private Sub TransferData(srcPath As String, destPath As String, logSheet As Worksheet)
    Dim wbSrc As Workbook, wbDest As Workbook
    Dim srcSheet As Worksheet, destSheet As Worksheet
    Dim srcData As Variant
    Dim logStatus As String, logDetail As String
    Dim lastRow As Long
    Dim srcFileName As String
    Dim isProfile As Boolean  ' 纵断面标识
    
    On Error GoTo ErrorHandler
    
    ' 获取源文件名（不带路径）
    srcFileName = fso.GetFileName(srcPath)
    
    ' 打开源文件（只读）和目标文件（可写）
    Set wbSrc = Workbooks.Open(srcPath, ReadOnly:=True, UpdateLinks:=0)
    Set srcSheet = wbSrc.Sheets(1)
    Set wbDest = Workbooks.Open(destPath, ReadOnly:=False)
    Set destSheet = wbDest.Sheets(1)
    
    ' 判断是否为纵断面文件（文件名包含"纵断面"）
    isProfile = InStr(1, srcFileName, "纵断面", vbTextCompare) > 0
    
    ' 根据文件类型获取数据
    If isProfile Then
        ' 纵断面数据处理（C列和E列）
        lastRow = FindLastRow(srcSheet, "C,E")
        If lastRow < 11 Then
            Err.Raise 1001, , "纵断面源数据区域（C11或E11）为空"
        End If
        
        ' 获取两列数据
        Dim colC As Variant, colE As Variant
        colC = srcSheet.Range("C11:C" & lastRow).value
        colE = srcSheet.Range("E11:E" & lastRow).value
        
        ' ==== 新增：对C列数据进行累加处理 ====
        Dim cumulative As Double
        cumulative = 0
        
        ' 创建新数组（行数×2列）
        ReDim srcData(1 To UBound(colC, 1), 1 To 2)
        Dim i As Long
        For i = 1 To UBound(colC, 1)
            ' 累加C列值
            If IsNumeric(colC(i, 1)) Then
                cumulative = cumulative + CDbl(colC(i, 1))
            End If
            srcData(i, 1) = cumulative  ' 使用累加值
            srcData(i, 2) = colE(i, 1)  ' E列值不变
        Next i
        ' ==== 累加处理结束 ====
        
        logDetail = "纵断面迁移: C11:C" & lastRow & " & E11:E" & lastRow & " (C列已累加)"
    Else
        ' 横断面数据处理（C列和D列）
        lastRow = FindLastRow(srcSheet, "C,D")
        If lastRow < 13 Then
            Err.Raise 1001, , "横断面源数据区域（C13:D）为空"
        End If
        srcData = srcSheet.Range("C13:D" & lastRow).value
        logDetail = "横断面迁移: C13:D" & lastRow
    End If
    
    ' 写入目标文件
    With destSheet
        .Range("A4:B200").ClearContents
        .Range("A4").Resize(UBound(srcData, 1), UBound(srcData, 2)).value = srcData
        .Parent.Save
    End With
    
    ' 记录成功日志并添加迁移类型
    logStatus = "成功"
    logDetail = logDetail & " → A4:B" & (3 + UBound(srcData, 1))
    GoTo Cleanup
    
ErrorHandler:
    logStatus = "失败"
    logDetail = "错误描述：" & Err.Description & " | 文件类型: " & IIf(isProfile, "纵断面", "横断面")
    
Cleanup:
    LogOperation logSheet, srcPath, destPath, logStatus, logDetail
    
    If Not wbSrc Is Nothing Then wbSrc.Close False
    If Not wbDest Is Nothing Then wbDest.Close SaveChanges:=(logStatus = "成功")
    Set wbSrc = Nothing
    Set wbDest = Nothing
End Sub

' 增强的FindLastRow函数（支持多列）
Private Function FindLastRow(sht As Worksheet, cols As String) As Long
    Dim colArray() As String
    Dim col As Variant
    Dim lastRow As Long, tempRow As Long
    
    ' 拆分列标识（如"C,D,E"）
    colArray = Split(cols, ",")
    
    ' 遍历所有列找出最大行号
    For Each col In colArray
        col = Trim(col)
        If Not sht.Columns(col).Find("*", , , , xlByRows, xlPrevious) Is Nothing Then
            tempRow = sht.Columns(col).Find("*", , , , xlByRows, xlPrevious).row
            If tempRow > lastRow Then lastRow = tempRow
        End If
    Next col
    
    FindLastRow = lastRow
End Function

' 记录操作日志到"处理日志"工作表
Private Sub LogOperation(logSheet As Worksheet, srcPath As String, destPath As String, status As String, detail As String)
    With logSheet.Cells(Rows.count, 1).End(xlUp).Offset(1)
        .value = fso.GetFileName(srcPath)  ' 源文件名
        .Offset(0, 1).value = fso.GetFileName(destPath)  ' 目标文件名
        .Offset(0, 2).value = fso.GetParentFolderName(srcPath)  ' 源文件路径
        .Offset(0, 3).value = fso.GetParentFolderName(destPath)  ' 目标文件路径
        .Offset(0, 4).value = status  ' 处理状态（成功/失败）
        .Offset(0, 5).value = Now  ' 时间戳
        .Offset(0, 6).value = detail  ' 详情（成功行数/错误描述）
    End With
End Sub

' 填充字典（从指定列提取桩号映射到F列值）
Private Sub FillDictionaryFromSheet(sht As Worksheet, dict As Object, colLetter As String)
    Dim lastRow As Long, i As Long
    Dim cellValue As String, stake As String, parts() As String
    
    lastRow = sht.Cells(sht.Rows.count, colLetter).End(xlUp).row
    For i = 1 To lastRow
        cellValue = Trim(sht.Cells(i, colLetter).value)
        If cellValue <> "" Then
            parts = Split(cellValue, "_")
            If UBound(parts) >= 0 Then
                stake = parts(UBound(parts))  ' 取最后一个下划线后的内容
                dict(stake) = sht.Cells(i, "F").value
            End If
        End If
    Next i
End Sub

' 修改输出文件（新增Sheet2支持）
Private Sub ModifyOutputFile(filePath As String, dictSheet1 As Object, dictSheet2 As Object, logCell As Range)
    Dim wb As Workbook, sht As Worksheet
    Dim fileName As String
    Dim logDetail As String
    
    On Error Resume Next
    Set wb = Workbooks.Open(filePath)
    If Err.Number <> 0 Then
        logCell.value = logCell.value & " | 修改失败：无法打开文件"
        Exit Sub
    End If
    
    ' 获取文件名（不含扩展名）
    fileName = fso.GetBaseName(filePath)
    
    ' 修改工作表
    Set sht = wb.Sheets(1)
    With sht
        ' 将文件名写入B2
        .Range("B2").value = fileName
        
        ' 修改工作表名称（需要错误处理）
        On Error Resume Next
        .name = fileName
        If Err.Number <> 0 Then
            logDetail = " | 工作表更名失败：" & Err.Description
            Err.Clear
        Else
            logDetail = " | 工作表更名成功"
        End If
        On Error GoTo 0
    End With
    
    ' 从字典查找对应值并写入A2 - 优先Sheet1，其次Sheet2
    Dim valueToFill As Variant
    valueToFill = ""
    
    ' 先在Sheet1字典中查找
    If Not dictSheet1 Is Nothing Then
        If dictSheet1.Exists(fileName) Then
            valueToFill = dictSheet1(fileName)
            logDetail = logDetail & " | Sheet1找到对应值"
        End If
    End If
    
    ' 如果Sheet1没找到，再查找Sheet2
    If valueToFill = "" And Not dictSheet2 Is Nothing Then
        If dictSheet2.Exists(fileName) Then
            valueToFill = dictSheet2(fileName)
            logDetail = logDetail & " | Sheet2找到对应值"
        End If
    End If
    
    ' 如果找到值则填入A2，否则记录未找到
    If valueToFill <> "" Then
        sht.Range("A2").value = valueToFill
        logDetail = logDetail & " | 已填充A2"
    Else
        logDetail = logDetail & " | 未找到对应值"
    End If
    
    ' 保存修改
    wb.Close SaveChanges:=True
    logCell.value = logCell.value & logDetail
End Sub











