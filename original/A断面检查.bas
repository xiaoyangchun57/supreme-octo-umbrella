Attribute VB_Name = "A断面检查"
Sub CheckNonSDCellsOptimized()
    Dim fDialog As FileDialog
    Dim folderPath As String
    Dim logSheet As Worksheet
    Dim logRow As Long
    Dim processedFiles As Long
    Dim nonSDCount As Long
    Dim startTime As Double
    
    ' 记录开始时间
    startTime = Timer
    
    ' 创建日志工作表
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("SD检查日志").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set logSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    logSheet.name = "SD检查日志"
    ' 增加一列用于记录检查类型
    logSheet.Range("A1:F1") = Array("文件路径", "文件名", "工作表", "单元格地址", "检查内容", "检查类型")
    logSheet.Range("A1:F1").Font.Bold = True
    logRow = 2
    
    ' 选择文件夹
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    fDialog.title = "选择包含横断面文件的文件夹"
    
    If fDialog.Show = -1 Then
        folderPath = fDialog.SelectedItems(1)
        If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
        
        ' 开始处理
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        
        processedFiles = 0
        nonSDCount = 0
        
        ' 递归处理文件夹
        ProcessFolder folderPath, logSheet, logRow, processedFiles, nonSDCount
        
        ' 恢复设置
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        
        ' 格式化日志
        With logSheet
            .Columns("A:F").AutoFit
            If logRow > 2 Then
                .ListObjects.Add(xlSrcRange, .Range("A1:F" & logRow - 1), , xlYes).name = "NonSDLogTable"
            End If
        End With
        
        ' 计算耗时
        Dim elapsedTime As Double
        elapsedTime = Round(Timer - startTime, 2)
        
        ' 显示结果
        MsgBox "处理完成！" & vbCrLf & _
               "扫描文件夹: " & folderPath & vbCrLf & _
               "处理文件数: " & processedFiles & vbCrLf & _
               "发现问题数: " & nonSDCount & vbCrLf & _
               "耗时: " & elapsedTime & " 秒", _
               vbInformation, "横断面文件检查"
    Else
        MsgBox "未选择文件夹。操作已取消。", vbExclamation
    End If
End Sub

Sub ProcessFolder(folderPath As String, logSheet As Worksheet, ByRef logRow As Long, ByRef processedFiles As Long, ByRef nonSDCount As Long)
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fileName As String
    Dim fileExt As String
    Dim lastRow As Long
    Dim dataRange As Range
    Dim dataArray As Variant
    Dim i As Long, j As Long
    Dim sdStart As Long, sdEnd As Long
    Dim sdFound As Boolean
    Dim filePath As String
    Dim cellValue As String
    Dim leftDykeFound As Boolean, rightDykeFound As Boolean
    Dim leftDykeRow As Long, rightDykeRow As Long
    Dim validKeywords As Variant
    
    ' 定义有效的关键词
    validKeywords = Array("横断面", "B", "Q", "K", "D", "J")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' 处理当前文件夹中的文件
    For Each file In folder.files
        fileName = file.name
        filePath = file.path
        fileExt = LCase(fso.GetExtensionName(fileName))
        
        ' 检查是否为Excel文件且文件名包含有效关键词
        If fileExt = "xlsx" Then
            Dim containsKeyword As Boolean
            containsKeyword = False
            For Each KeyWord In validKeywords
                If InStr(1, fileName, KeyWord, vbTextCompare) > 0 Then
                    containsKeyword = True
                    Exit For
                End If
            Next KeyWord
            
            If Not containsKeyword Then GoTo SkipFile
            
            On Error Resume Next
            Set wb = Workbooks.Open(filePath, True, True) ' 仅读模式，不更新链接
            
            If Err.Number = 0 Then
                processedFiles = processedFiles + 1
                
                ' 处理每个工作表
                For Each ws In wb.Worksheets
                    ' 获取B列最后一行
                    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
                    If lastRow < 13 Then GoTo NextWorksheet ' 跳过不足13行的工作表
                    
                    ' 使用数组处理数据
                    Set dataRange = ws.Range("B13:D" & lastRow) ' 修改为包含D列
                    dataArray = dataRange.value
                    
                    ' 重置堤顶状态
                    leftDykeFound = False
                    rightDykeFound = False
                    leftDykeRow = 0
                    rightDykeRow = 0
                    
                    ' 查找SD段落的起始和结束位置 (包含SD、SD1、SD0等)
                    sdStart = 0
                    sdEnd = 0
                    sdFound = False
                    
                    ' 定义LSH、ZJ和YJ相关变量
                    Dim lshFound As Boolean
                    Dim lshRow As Long
                    Dim lshElevation As Double
                    Dim zjFound As Boolean
                    Dim zjElevation As Double
                    Dim yjFound As Boolean
                    Dim yjElevation As Double
                    
                    lshFound = False
                    zjFound = False
                    yjFound = False
                    
                    ' 查找堤顶位置、SD段落以及LSH/ZJ/YJ
                    For i = 1 To UBound(dataArray, 1)
                        If Not IsEmpty(dataArray(i, 1)) Then
                            cellValue = CStr(dataArray(i, 1))
                            
                            ' 查找左堤顶
                            If Not leftDykeFound And cellValue = "左堤顶" Then
                                leftDykeFound = True
                                leftDykeRow = i
                            End If
                            
                            ' 查找右堤顶
                            If Not rightDykeFound And cellValue = "右堤顶" Then
                                rightDykeFound = True
                                rightDykeRow = i
                            End If
                            
                            ' 查找SD开头的单元格
                            If Left(cellValue, 2) = "SD" Then
                                If Not sdFound Then
                                    sdStart = i
                                    sdFound = True
                                End If
                                sdEnd = i ' 更新最后一个SD位置
                            End If
                            
                            ' 查找LSH单元格（非深泓点）
                            If UCase(cellValue) = "LSH" Then
                                If IsNumeric(dataArray(i, 3)) Then
                                    lshFound = True
                                    lshRow = i
                                    lshElevation = CDbl(dataArray(i, 3))
                                End If
                            End If
                            
                            ' 查找ZJ单元格
                            If UCase(cellValue) = "ZJ" Then
                                If IsNumeric(dataArray(i, 3)) Then
                                    zjFound = True
                                    zjElevation = CDbl(dataArray(i, 3))
                                End If
                            End If
                            
                            ' 查找YJ单元格
                            If UCase(cellValue) = "YJ" Then
                                If IsNumeric(dataArray(i, 3)) Then
                                    yjFound = True
                                    yjElevation = CDbl(dataArray(i, 3))
                                End If
                            End If
                        End If
                    Next i
                    
                    ' 如果没有找到最后一个SD，但有SD存在，使用第一个SD位置
                    If sdFound And sdEnd = 0 Then sdEnd = sdStart
                    
                    ' 检查SD段落中的非SD单元格
                    If sdFound And sdEnd > 0 Then
                        For i = sdStart To sdEnd
                            If Not IsEmpty(dataArray(i, 1)) Then
                                cellValue = CStr(dataArray(i, 1))
                                ' 排除所有SD开头和深泓点的单元格
                                If Left(cellValue, 2) <> "SD" And cellValue <> "深泓点" Then
                                    ' 记录非SD内容
                                    logSheet.Cells(logRow, 1).value = filePath
                                    logSheet.Cells(logRow, 2).value = fileName
                                    logSheet.Cells(logRow, 3).value = ws.name
                                    logSheet.Cells(logRow, 4).value = "B" & (i + 12) ' 行号 = 13 + i - 1
                                    logSheet.Cells(logRow, 5).value = dataArray(i, 1)
                                    logSheet.Cells(logRow, 6).value = "非SD内容"
                                    logRow = logRow + 1
                                    nonSDCount = nonSDCount + 1
                                End If
                            End If
                        Next i
                    End If
                    
                    ' 检查左堤顶位置是否正确
                    If leftDykeFound Then
                        Dim foundLeftMark As Boolean
                        foundLeftMark = False
                        
                        ' 修正：检查左堤顶下方（之后的行）是否有正确标记
                        For j = leftDykeRow + 1 To UBound(dataArray, 1)
                            If Not IsEmpty(dataArray(j, 1)) Then
                                cellValue = CStr(dataArray(j, 1))
                                If UCase(cellValue) = "SB" Or UCase(cellValue) = "ZSB" Or _
                                   UCase(cellValue) = "DJ" Or UCase(cellValue) = "ZDJ" Then
                                    foundLeftMark = True
                                    Exit For
                                End If
                            End If
                        Next j
                        
                        ' 如果没有找到正确标记，记录错误
                        If Not foundLeftMark Then
                            logSheet.Cells(logRow, 1).value = filePath
                            logSheet.Cells(logRow, 2).value = fileName
                            logSheet.Cells(logRow, 3).value = ws.name
                            logSheet.Cells(logRow, 4).value = "B" & (leftDykeRow + 12)
                            logSheet.Cells(logRow, 5).value = "下方缺少SB/ZSB或DJ/ZDJ"
                            logSheet.Cells(logRow, 6).value = "左堤顶下方无标记"
                            logRow = logRow + 1
                            nonSDCount = nonSDCount + 1
                        End If
                    End If
                    
                    ' 检查右堤顶位置是否正确 - 修正：检查上方
                    If rightDykeFound Then
                        Dim foundRightMark As Boolean
                        foundRightMark = False
                        
                        ' 修正：检查右堤顶上方（之前的行）是否有正确标记
                        For j = rightDykeRow - 1 To 1 Step -1
                            If Not IsEmpty(dataArray(j, 1)) Then
                                cellValue = CStr(dataArray(j, 1))
                                If UCase(cellValue) = "SB" Or UCase(cellValue) = "YSB" Or _
                                   UCase(cellValue) = "DJ" Or UCase(cellValue) = "YDJ" Then
                                    foundRightMark = True
                                    Exit For
                                End If
                            End If
                        Next j
                        
                        ' 如果没有找到正确标记，记录错误
                        If Not foundRightMark Then
                            logSheet.Cells(logRow, 1).value = filePath
                            logSheet.Cells(logRow, 2).value = fileName
                            logSheet.Cells(logRow, 3).value = ws.name
                            logSheet.Cells(logRow, 4).value = "B" & (rightDykeRow + 12)
                            logSheet.Cells(logRow, 5).value = "上方缺少SB/YSB或DJ/YDJ"
                            logSheet.Cells(logRow, 6).value = "右堤顶上方无标记"
                            logRow = logRow + 1
                            nonSDCount = nonSDCount + 1
                        End If
                    End If
                    
                    ' 检查LSH高程是否符合要求
                    If lshFound Then
                        Dim lshError As String
                        lshError = ""
                        
                        ' 检查ZJ高程是否大于LSH高程0.5米以上
                        If zjFound Then
                            If zjElevation <= lshElevation + 0.5 Then
                                lshError = lshError & "ZJ高程(" & Format(zjElevation, "0.000") & _
                                          ")未超过LSH高程(" & Format(lshElevation, "0.000") & ")0.5米; "
                            End If
                        Else
                            lshError = lshError & "未找到ZJ; "
                        End If
                        
                        ' 检查YJ高程是否大于LSH高程0.5米以上
                        If yjFound Then
                            If yjElevation <= lshElevation + 0.5 Then
                                lshError = lshError & "YJ高程(" & Format(yjElevation, "0.000") & _
                                          ")未超过LSH高程(" & Format(lshElevation, "0.000") & ")0.5米; "
                            End If
                        Else
                            lshError = lshError & "未找到YJ; "
                        End If
                        
                        ' 如果发现错误，记录到日志
                        If lshError <> "" Then
                            logSheet.Cells(logRow, 1).value = filePath
                            logSheet.Cells(logRow, 2).value = fileName
                            logSheet.Cells(logRow, 3).value = ws.name
                            logSheet.Cells(logRow, 4).value = "B" & (lshRow + 12) & ", D" & (lshRow + 12)
                            logSheet.Cells(logRow, 5).value = lshError
                            logSheet.Cells(logRow, 6).value = "LSH高程不足"
                            logRow = logRow + 1
                            nonSDCount = nonSDCount + 1
                        End If
                    End If
                    
NextWorksheet:
                Next ws
                
                wb.Close False
            Else
                ' 记录无法打开的文件
                logSheet.Cells(logRow, 1).value = filePath
                logSheet.Cells(logRow, 2).value = fileName
                logSheet.Cells(logRow, 3).value = "无法打开文件"
                logSheet.Cells(logRow, 4).value = ""
                logSheet.Cells(logRow, 5).value = Err.Description
                logSheet.Cells(logRow, 6).value = "文件打开错误"
                logRow = logRow + 1
                Err.Clear
            End If
            On Error GoTo 0
        End If
SkipFile:
    Next file
    
    ' 递归处理子文件夹
    For Each subFolder In folder.subFolders
        ProcessFolder subFolder.path, logSheet, logRow, processedFiles, nonSDCount
    Next subFolder
    
    Set fso = Nothing
End Sub



