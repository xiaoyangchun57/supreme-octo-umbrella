Attribute VB_Name = "CSV转成果表一键横断面最终版"



Option Explicit

' ===== 进度条模块 =====
Public ProgressForm As UserForm1
Public logFilePath As String  ' 日志文件路径
Public templatePath As String ' 模板文件路径

Public Sub UpdateProgressBar(fileName As String, folderPath As String, current As Long, total As Long)
    If ProgressForm Is Nothing Then
        Set ProgressForm = New UserForm1
        ProgressForm.caption = "文件处理中"
        ProgressForm.Show vbModeless
    End If
    
    With ProgressForm
        .lblFileName.caption = "当前文件: " & fileName
        .lblFolder.caption = "文件夹: " & folderPath
        .lblCount.caption = "进度: " & current & " / " & total & " (" & Format(current / total, "0.0%") & ")"
        
        Dim progressWidth As Long
        progressWidth = (current / total) * (.Frame1.width - 4)
        progressWidth = Application.WorksheetFunction.Max(0, _
                    Application.WorksheetFunction.Min(progressWidth, .Frame1.width - 4))
        .Progress.width = progressWidth
        
        Static lastUpdate As Double, lastCount As Long
        Dim elapsedTime As Double
        If lastUpdate = 0 Then lastUpdate = Timer
        elapsedTime = Timer - lastUpdate
        
        If elapsedTime > 1 Then
            If lastCount > 0 Then
                Dim speed As Double
                speed = (current - lastCount) / elapsedTime
                .lblSpeed.caption = "速度: " & Format(speed, "0.0") & " 文件/秒"
                
                If speed > 0 Then
                    Dim remaining As Double
                    remaining = (total - current) / speed
                    .lblRemaining.caption = "剩余时间: " & FormatTime(remaining)
                End If
            End If
            lastUpdate = Timer
            lastCount = current
        End If
        
        .Repaint
    End With
End Sub

Private Function FormatTime(seconds As Double) As String
    If seconds <= 0 Then
        FormatTime = "00:00:00"
        Exit Function
    End If
    
    Dim minutes As Long, hours As Long
    minutes = Int(seconds / 60)
    seconds = seconds - minutes * 60
    hours = Int(minutes / 60)
    minutes = minutes - hours * 60
    
    FormatTime = Format(hours, "00") & ":" & Format(minutes, "00") & ":" & Format(Int(seconds), "00")
End Function

' ===== 日志写入函数 =====
Private Sub WriteToLog(msg As String)
    On Error Resume Next ' 防止日志写入失败中断程序
    Dim fileNo As Integer
    fileNo = FreeFile
    
    ' 使用追加模式写入日志
    Open logFilePath For Append As #fileNo
    Print #fileNo, Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & msg
    Close #fileNo
End Sub

' ===== 主处理程序 =====
Sub ProcessCSVToXLSX()
    Dim fso As Object
    Dim rootFolder As String
    Dim totalCount As Long, progressCount As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "请选择包含CSV文件的根文件夹"
        If .Show = -1 Then
            rootFolder = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ' 初始化日志文件
    logFilePath = rootFolder & "\error_log.txt"
    ' 创建新日志文件
    Open logFilePath For Output As #1
    Close #1
    
    ' 查找合适的模板文件
    templatePath = FindTemplateFile(rootFolder)
    If templatePath = "" Then
        MsgBox "找不到合适的模板文件! 请确保模板文件位于根文件夹中。", vbExclamation
        Exit Sub
    Else
        WriteToLog "使用模板文件: " & templatePath
    End If
    
    totalCount = CountCSVFiles(fso, rootFolder)
    If totalCount = 0 Then
        MsgBox "未找到CSV文件!", vbExclamation
        Exit Sub
    End If
    
    UpdateProgressBar "正在初始化...", rootFolder, 0, totalCount
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    progressCount = 0
    ProcessCSVFiles fso, rootFolder, totalCount, progressCount
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    If Not ProgressForm Is Nothing Then
        Unload ProgressForm
        Set ProgressForm = Nothing
    End If
    
    ' 检查日志文件是否有内容
    If FileLen(logFilePath) > 0 Then
        MsgBox "处理完成! 共处理 " & progressCount & " 个文件" & vbCrLf & _
               "部分文件处理失败，请查看日志: " & logFilePath, vbExclamation
    Else
        MsgBox "处理完成! 共处理 " & progressCount & " 个文件", vbInformation
    End If
End Sub

' 查找合适的模板文件
Function FindTemplateFile(rootFolder As String) As String
    Dim fso As Object, folder As Object, file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 先在根文件夹中查找模板文件
    Set folder = fso.GetFolder(rootFolder)
    For Each file In folder.files
        If InStr(1, file.name, "成果", vbTextCompare) > 0 And _
           LCase(fso.GetExtensionName(file.name)) = "xlsx" Then
            FindTemplateFile = file.path
            Exit Function
        End If
    Next
    
    ' 如果没找到，在子文件夹中查找
    For Each folder In folder.subFolders
        For Each file In folder.files
            If InStr(1, file.name, "成果", vbTextCompare) > 0 And _
               LCase(fso.GetExtensionName(file.name)) = "xlsx" Then
                FindTemplateFile = file.path
                Exit Function
            End If
        Next
    Next
    
    FindTemplateFile = ""
End Function



' ===== 坐标投影预处理函数 =====
Sub AdjustCSVPoints(ws As Worksheet)
    Dim lastRow As Long, zjRow As Long, yjRow As Long, newZJRow As Long, newYJRow As Long
    Dim zjX As Double, zjY As Double, yjX As Double, yjY As Double
    Dim lineVectorX As Double, lineVectorY As Double
    Dim vectorLength As Double, dotProduct As Double, t As Double
    Dim arrData As Variant, arrTemp As Variant
    Dim projPoints As Object
    Dim i As Long, j As Long
    Dim cellValue As String
    Dim hasYJ As Boolean
    
    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.count, "E").End(xlUp).row
    
    ' 寻找ZJ和YJ
    zjRow = 0
    yjRow = 0
    For i = 1 To lastRow
        cellValue = UCase(Trim(ws.Cells(i, "E").value))
        If cellValue = "ZJ" Then zjRow = i
        If cellValue = "YJ" Then yjRow = i
    Next i
    
    ' 检查是否找到必要的点
    If zjRow = 0 Or yjRow = 0 Then
        If zjRow = 0 Then WriteToLog "未找到ZJ点: " & ws.Parent.name
        If yjRow = 0 Then WriteToLog "未找到YJ点: " & ws.Parent.name
        Exit Sub
    End If
    
    ' 获取坐标
    zjY = ws.Cells(zjRow, "B").value
    zjX = ws.Cells(zjRow, "C").value
    yjY = ws.Cells(yjRow, "B").value
    yjX = ws.Cells(yjRow, "C").value
    
    ' 计算直线方向向量
    lineVectorX = yjX - zjX
    lineVectorY = yjY - zjY
    vectorLength = (lineVectorX ^ 2 + lineVectorY ^ 2)
    
    ' 检查有效性
    If vectorLength <= 0.001 Then  ' 避免浮点误差
        WriteToLog "错误: ZJ和YJ点位置过于接近，文件 " & ws.Parent.name
        Exit Sub
    End If
    
    ' 获取所有点的数据
    arrData = ws.Range("A1:E" & lastRow).value
    
    ' 存储投影参数
    Set projPoints = CreateObject("Scripting.Dictionary")
    For i = 1 To lastRow
        dotProduct = (arrData(i, 3) - zjX) * lineVectorX + (arrData(i, 2) - zjY) * lineVectorY
        projPoints(i) = dotProduct / vectorLength
    Next i
    
    ' 获取排序后的索引
    Dim keys() As Variant
    keys = projPoints.keys()
    Call BubbleSort(projPoints, keys)
    
    ' 创建新的数据数组
    ReDim arrTemp(1 To lastRow, 1 To 5)
    
    ' 按投影顺序重新排列点
    For i = LBound(keys) To UBound(keys)
        For j = 1 To 5
            arrTemp(i + 1, j) = arrData(keys(i), j)
        Next j
    Next i
    
    ' 将新数据写回工作表
    ws.Range("A1:E" & lastRow).value = arrTemp
    
    ' 查找新的ZJ和YJ位置
    newZJRow = 0
    newYJRow = 0
    For i = 1 To lastRow
        cellValue = UCase(Trim(ws.Cells(i, "E").value))
        If cellValue = "ZJ" Then newZJRow = i
        If cellValue = "YJ" Then newYJRow = i
    Next i
    
    ' 确保ZJ在第一行
    If newZJRow > 1 Then
        ws.Rows(newZJRow).Cut
        ws.Rows(1).Insert Shift:=xlDown
        newZJRow = 1  ' 更新为新的位置
    End If
    
    ' 确保YJ在最后一行
    If newYJRow < lastRow Then
        ws.Rows(newYJRow).Cut
        ws.Rows(lastRow).Insert Shift:=xlDown
        newYJRow = lastRow  ' 更新为新的位置
    End If
    
    ' 重新获取行数和数据
    lastRow = ws.Cells(ws.Rows.count, "E").End(xlUp).row
    arrData = ws.Range("A1:E" & lastRow).value
    
    ' 重新计算坐标
    zjX = ws.Cells(1, "C").value
    zjY = ws.Cells(1, "B").value
    yjX = ws.Cells(lastRow, "C").value
    yjY = ws.Cells(lastRow, "B").value
    lineVectorX = yjX - zjX
    lineVectorY = yjY - zjY
    vectorLength = (lineVectorX ^ 2 + lineVectorY ^ 2)
    
    ' 重新计算中间点坐标（投影到ZJ-YJ直线上）
    For i = 2 To lastRow - 1
        dotProduct = (arrData(i, 3) - zjX) * lineVectorX + (arrData(i, 2) - zjY) * lineVectorY
        t = dotProduct / vectorLength
        ws.Cells(i, "C").value = zjX + t * lineVectorX  ' 更新X坐标
        ws.Cells(i, "B").value = zjY + t * lineVectorY  ' 更新Y坐标
    Next i
    
    ' 确保YJ在最后一行
    hasYJ = False
    For i = 1 To lastRow
        If UCase(Trim(ws.Cells(i, "E").value)) = "YJ" Then
            hasYJ = True
            If i <> lastRow Then
                ws.Rows(i).Cut
                ws.Rows(lastRow).Insert Shift:=xlDown
            End If
            Exit For
        End If
    Next i
    
    If Not hasYJ Then
        WriteToLog "错误: 文件 " & ws.Parent.name & " 中未找到YJ点"
    End If
End Sub

' ===== 高效冒泡排序实现 =====
Sub BubbleSort(dict As Object, arr() As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim swapped As Boolean
    
    For i = LBound(arr) To UBound(arr) - 1
        swapped = False
        For j = LBound(arr) To UBound(arr) - i - 1
            If dict(arr(j)) > dict(arr(j + 1)) Then
                ' 交换元素
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
                swapped = True
            End If
        Next j
        
        ' 如果本轮没有交换，说明已经有序
        If Not swapped Then Exit For
    Next i
End Sub

' ===== CSV文件处理 =====
Sub ProcessCSVFiles(fso As Object, folderPath As String, totalCount As Long, ByRef progressCount As Long)
    Dim folder As Object, subFolder As Object, file As Object
    Dim csvPath As String, xlsxPath As String
    Dim wbCSV As Workbook, wbXLSX As Workbook
    Dim wsCSV As Worksheet, wsXLSX As Worksheet
    Dim lastRow As Long, i As Long
    Dim zjRow As Long, zjX As Double, zjY As Double
    Dim cellValue As String
    Dim startDistance As Double
    Dim Longitude As Double, Latitude As Double
    
    Set folder = fso.GetFolder(folderPath)
    
    For Each file In folder.files
 ' === 新增：跳过文件名包含"Z"的文件 ===
        If InStr(1, file.name, "Z", vbTextCompare) > 0 Then
            GoTo NextFile
        End If
        If LCase(fso.GetExtensionName(file.name)) = "csv" Then
            csvPath = file.path
            xlsxPath = fso.BuildPath(folderPath, fso.GetBaseName(csvPath) & ".xlsx")
            
            progressCount = progressCount + 1
            UpdateProgressBar file.name, folderPath, progressCount, totalCount
            
            On Error Resume Next
            Set wbCSV = Workbooks.Open(csvPath)
            If Err.Number <> 0 Then
                WriteToLog "打开CSV失败: " & file.name & " (错误: " & Err.Description & ")"
                Err.Clear
                GoTo NextFile
            End If
            
            Set wsCSV = wbCSV.Sheets(1)
            
            ' 添加点预处理步骤
            AdjustCSVPoints wsCSV
            wbCSV.Save ' 保存预处理后的CSV
            
            ' 创建目标文件
            If Not fso.FileExists(xlsxPath) Then
                WriteToLog "创建XLSX文件: " & file.name
                fso.CopyFile templatePath, xlsxPath
                If Err.Number <> 0 Then
                    WriteToLog "创建XLSX文件失败: " & file.name & " (错误: " & Err.Description & ")"
                    Err.Clear
                    GoTo NextFile
                End If
            End If
            
            Set wbXLSX = Workbooks.Open(xlsxPath)
            If Err.Number <> 0 Then
                WriteToLog "打开XLSX失败: " & file.name & " (错误: " & Err.Description & ")"
                Err.Clear
                wbCSV.Close SaveChanges:=False
                GoTo NextFile
            End If
            Set wsXLSX = wbXLSX.Sheets(1)
            On Error GoTo 0
            
            lastRow = wsCSV.Cells(wsCSV.Rows.count, "E").End(xlUp).row
            
            zjRow = 0
            For i = 1 To lastRow
                cellValue = UCase(Trim(wsCSV.Cells(i, "E").value))
                If cellValue = "ZJ" Then
                    zjRow = i
                    Exit For
                End If
            Next i
            
            If zjRow = 0 Then
                WriteToLog "缺少ZJ记录: " & file.name
                wbCSV.Close SaveChanges:=False
                wbXLSX.Close SaveChanges:=False
                GoTo NextFile
            End If
            
            zjY = wsCSV.Cells(zjRow, "B").value
            zjX = wsCSV.Cells(zjRow, "C").value
            
            ' 获取YJ坐标（最后一行）
            Dim yjX As Double, yjY As Double
            yjY = wsCSV.Cells(lastRow, "B").value
            yjX = wsCSV.Cells(lastRow, "C").value
            
            ' 存储所有堤顶点信息（行号、起点距、原始标记）
            Dim leveePoints As Collection
            Set leveePoints = New Collection
            
            ' 第一阶段：收集所有堤顶点信息
            For i = 1 To lastRow
                If IsEmpty(wsCSV.Cells(i, "A").value) Then GoTo NextRow1
                
                startDistance = Sqr((wsCSV.Cells(i, "C").value - zjX) ^ 2 + _
                                    (wsCSV.Cells(i, "B").value - zjY) ^ 2)
                
                Dim originalMarker As String
                originalMarker = UCase(Trim(wsCSV.Cells(i, "E").value))
                
                ' 检查是否为堤顶点
                If IsLeveeMarker(originalMarker) Then
                    ' 存储行号、起点距和原始标记
                    leveePoints.Add Array(i, startDistance, originalMarker)
                End If
NextRow1:
            Next i
            
            ' 确定左右堤顶
            Dim minLeveeDist As Double, maxLeveeDist As Double
            Dim minLeveeRow As Long, maxLeveeRow As Long
            Dim minMarker As String, maxMarker As String
            
            If leveePoints.count > 0 Then
                minLeveeDist = leveePoints(1)(1)
                maxLeveeDist = minLeveeDist
                minLeveeRow = leveePoints(1)(0)
                maxLeveeRow = minLeveeRow
                minMarker = leveePoints(1)(2)
                maxMarker = minMarker
                
                ' 找出最小和最大起点距的堤顶点
                For i = 2 To leveePoints.count
                    If leveePoints(i)(1) < minLeveeDist Then
                        minLeveeDist = leveePoints(i)(1)
                        minLeveeRow = leveePoints(i)(0)
                        minMarker = leveePoints(i)(2)
                    End If
                    If leveePoints(i)(1) > maxLeveeDist Then
                        maxLeveeDist = leveePoints(i)(1)
                        maxLeveeRow = leveePoints(i)(0)
                        maxMarker = leveePoints(i)(2)
                    End If
                Next i
            End If
            
            Dim xlsxRow As Long
            xlsxRow = 13
            
            ' 第二阶段：处理所有点
            For i = 1 To lastRow
                If IsEmpty(wsCSV.Cells(i, "A").value) Then GoTo NextRow2
                
                startDistance = Sqr((wsCSV.Cells(i, "C").value - zjX) ^ 2 + _
                                    (wsCSV.Cells(i, "B").value - zjY) ^ 2)
                
                originalMarker = UCase(Trim(wsCSV.Cells(i, "E").value))
                
                Dim finalMarker As String
                finalMarker = originalMarker   ' 默认保留原值
                
                ' 处理堤顶点标记
                If IsLeveeMarker(originalMarker) Then
                    If i = minLeveeRow Then
                        finalMarker = "左堤顶"
                    ElseIf i = maxLeveeRow Then
                        finalMarker = "右堤顶"
                    Else
                        ' 中间的堤顶点保留原始标记
                        finalMarker = originalMarker
                    End If
                End If
                
                With wsXLSX
                    .Cells(xlsxRow, "C").value = Round(startDistance, 1)
                    .Cells(xlsxRow, "B").value = finalMarker
                    .Cells(xlsxRow, "D").value = Round(wsCSV.Cells(i, "D").value, 3)
                    .Cells(xlsxRow, "G").value = wsCSV.Cells(i, "F").value
                    
                    On Error Resume Next
                     GaussProjInv wsCSV.Cells(i, "B").value, wsCSV.Cells(i, "C").value, 117#, Latitude, Longitude
            If Err.Number <> 0 Then
                WriteToLog "坐标转换失败: " & file.name & " (行号: " & i & ")"
                Err.Clear
                Longitude = 0
                Latitude = 0
            End If
            On Error GoTo 0
            
            .Cells(xlsxRow, "E").value = Round(Longitude, 6)
            .Cells(xlsxRow, "F").value = Round(Latitude, 6)
            .Cells(xlsxRow, "A").value = xlsxRow - 12
        End With
        
        xlsxRow = xlsxRow + 1
NextRow2:
    Next i
            
            CalculateAzimuth wsXLSX, xlsxRow
            ProcessSpecialMarkers wsXLSX, xlsxRow
            
            ' 设置格式处理
            FormatWorksheet wsXLSX, xlsxRow - 1
            
            ' === 新增：标记深泓点 ===
            MarkDeepestPoint wsXLSX, 14, xlsxRow - 1 ' 从第14行到最后一行
            
            wbXLSX.Close SaveChanges:=True
            wbCSV.Close SaveChanges:=False
        End If
NextFile:
    Next file
    
    For Each subFolder In folder.subFolders
        ProcessCSVFiles fso, subFolder.path, totalCount, progressCount
    Next subFolder
End Sub

' 判断是否为堤顶点标记
Function IsLeveeMarker(marker As String) As Boolean
    marker = UCase(Trim(marker))
    If marker = "KS" Or marker = "ZKS" Or marker = "YKS" Or _
       marker = "DD" Or marker = "ZDD" Or marker = "YDD" Then
        IsLeveeMarker = True
    Else
        IsLeveeMarker = False
    End If
End Function

' ===== 新增：标记深泓点 =====
Sub MarkDeepestPoint(ws As Worksheet, startRow As Long, endRow As Long)
    Dim minElevation As Double
    Dim minRow As Long
    Dim i As Long
    Dim elevation As Double
    
    ' 初始化最小高程值和行号
    minElevation = 9999999
    minRow = 0
    
    ' 查找最小高程值及其行号
    For i = startRow To endRow
        If IsNumeric(ws.Cells(i, "D").value) Then
            elevation = ws.Cells(i, "D").value
            If elevation < minElevation Then
                minElevation = elevation
                minRow = i
            End If
        End If
    Next i
    
    ' 如果找到最小高程点，将其标记为"深泓点"
    If minRow > 0 Then
        ws.Cells(minRow, "B").value = "深泓点"
    End If
End Sub

' ===== 文件计数 =====
Function CountCSVFiles(fso As Object, folderPath As String) As Long
    Dim folder As Object, subFolder As Object, file As Object
    Dim count As Long
    
    Set folder = fso.GetFolder(folderPath)
    
    For Each file In folder.files
        If LCase(fso.GetExtensionName(file.name)) = "csv" Then
            count = count + 1
        End If
    Next file
    
    For Each subFolder In folder.subFolders
        count = count + CountCSVFiles(fso, subFolder.path)
    Next subFolder
    
    CountCSVFiles = count
End Function
Sub GaussProjInv(x As Double, y As Double, centralMeridian As Double, B As Double, L As Double)
    ' 椭球参数 (WGS84)
    Const a As Double = 6378137#             ' 长半轴
    Const f As Double = 1# / 298.257223563   ' 扁率
    Const e2 As Double = 2 * f - f * f       ' 第一偏心率的平方
    Const e12 As Double = e2 / (1 - e2)      ' 第二偏心率的平方
    
    ' 将中央经线转换为弧度
    Dim centralMeridianRad As Double
    centralMeridianRad = centralMeridian * 3.14159265358979 / 180#
    
    ' 去掉东偏移500000米
    Dim y0 As Double
    y0 = y - 500000
    
    ' 初始化底点纬度
    Dim Bf_rad As Double, Bf0_rad As Double
    Dim M As Double
    Dim sinBf As Double, cosBf As Double
    Dim tf As Double, tf2 As Double, nf2 As Double
    Dim i As Integer
    Dim delta_l_rad As Double
    Dim Nf As Double, Mf As Double  ' 增加Mf计算
    Dim B_rad As Double  ' 修复：定义B_rad变量
    
    ' 初始值
    Bf_rad = x / (a * (1 - e2 / 4 - 3 * e2 * e2 / 64 - 5 * e2 * e2 * e2 / 256))
    
    ' 迭代计算底点纬度
    For i = 0 To 4
        sinBf = Sin(Bf_rad)
        cosBf = Cos(Bf_rad)
        tf = sinBf / cosBf
        
        ' 计算子午线弧长M
        M = a * ((1 - e2 / 4 - 3 * e2 * e2 / 64 - 5 * e2 * e2 * e2 / 256) * Bf_rad _
            - (3 * e2 / 8 + 3 * e2 * e2 / 32 + 45 * e2 * e2 * e2 / 1024) * Sin(2 * Bf_rad) _
            + (15 * e2 * e2 / 256 + 45 * e2 * e2 * e2 / 1024) * Sin(4 * Bf_rad) _
            - (35 * e2 * e2 * e2 / 3072) * Sin(6 * Bf_rad))
        
        ' 更新Bf_rad
        Bf0_rad = Bf_rad
        Bf_rad = (x - M) / (a * (1 - e2)) + Bf_rad
        If Abs(Bf_rad - Bf0_rad) < 0.0000000001 Then Exit For
    Next i
    
    sinBf = Sin(Bf_rad)
    cosBf = Cos(Bf_rad)
    tf = sinBf / cosBf
    tf2 = tf * tf
    nf2 = e12 * cosBf * cosBf
    Nf = a / Sqr(1 - e2 * sinBf * sinBf)
    
    ' === 关键修正：计算Mf（子午圈曲率半径） ===
    Mf = a * (1 - e2) / ((1 - e2 * sinBf * sinBf) ^ 1.5)
    
    ' 计算经差
    delta_l_rad = y0 / Nf
    
    ' === 使用正确的纬度计算公式 ===
    B_rad = Bf_rad - (y0 * y0 * tf) / (2 * Mf * Nf) _
            + (y0 * y0 * y0 * y0 * tf) / (24 * Mf * Nf * Nf * Nf) * (5 + 3 * tf2 + nf2 - 9 * nf2 * tf2) _
            - (y0 * y0 * y0 * y0 * y0 * y0 * tf) / (720 * Mf * Nf * Nf * Nf * Nf * Nf) * (61 + 90 * tf2 + 45 * tf2 * tf2)
    
    ' 计算经度L（弧度）
    Dim L_rad As Double
    L_rad = delta_l_rad / cosBf _
            - (1 + 2 * tf2 + nf2) * (delta_l_rad * delta_l_rad * delta_l_rad) / (6 * cosBf) _
            + (5 + 28 * tf2 + 24 * tf2 * tf2 + 6 * nf2 + 8 * nf2 * tf2) * (delta_l_rad * delta_l_rad * delta_l_rad * delta_l_rad * delta_l_rad) / (120 * cosBf)
    
    ' 转换为度
    B = B_rad * 180# / 3.14159265358979
    L = centralMeridian + L_rad * 180# / 3.14159265358979
End Sub

' ===== 修复后的方位角计算子过程 =====
Sub CalculateAzimuth(ws As Worksheet, lastRow As Long)
    If lastRow >= 14 Then
        Dim zjLon As Double, zjLat As Double
        Dim yjLon As Double, yjLat As Double
        Dim azimuthDeg As Double
        Dim yjFound As Boolean
        
        Const PI As Double = 3.14159265358979
        
        ' 初始化YJ查找状态
        yjFound = False
        
        On Error Resume Next ' 防止读取错误
        
        ' 获取ZJ点的经纬度（固定在第13行）
        zjLon = ws.Cells(13, "E").value
        zjLat = ws.Cells(13, "F").value
        
        ' 精确查找YJ位置（遍历所有行）
        Dim i As Long
        For i = 13 To lastRow
            If UCase(Trim(ws.Cells(i, "B").value)) = "YJ" Then
                yjLon = ws.Cells(i, "E").value
                yjLat = ws.Cells(i, "F").value
                yjFound = True
                Exit For ' 找到即退出
            End If
        Next i
        
        ' 若未找到YJ则使用最后一行（兼容旧逻辑）
        If Not yjFound Then
            yjLon = ws.Cells(lastRow, "E").value
            yjLat = ws.Cells(lastRow, "F").value
        End If
        
        ' 使用新公式计算方位角
        azimuthDeg = Mod360( _
            Degrees( _
                WorksheetFunction.Atan2( _
                    Cos(radians(zjLat)) * Sin(radians(yjLat)) - _
                    Sin(radians(zjLat)) * Cos(radians(yjLat)) * _
                    Cos(radians(yjLon) - radians(zjLon)), _
                    Cos(radians(yjLat)) * Sin(radians(yjLon) - radians(zjLon)) _
                ) _
            ) _
        )
        
        With ws.Range("B9")
            .value = Round(azimuthDeg, 4)
            .NumberFormat = "0.0000"
        End With
        On Error GoTo 0
    Else
        ws.Range("B9").ClearContents
    End If
End Sub

' ===== 辅助函数 =====
Function radians(Degrees As Double) As Double
    radians = Degrees * 3.14159265358979 / 180#
End Function

Function Degrees(radians As Double) As Double
    Degrees = radians * 180# / 3.14159265358979
End Function

Function Mod360(value As Double) As Double
    Mod360 = value - Int(value / 360) * 360
    If Mod360 < 0 Then Mod360 = Mod360 + 360
End Function

Sub ProcessSpecialMarkers(ws As Worksheet, lastRow As Long)
    Dim i As Long, markValue As String
    
    On Error Resume Next ' 捕获错误
    For i = 13 To lastRow
        markValue = UCase(Trim(ws.Cells(i, "B").value))
        
        Select Case markValue
            Case "ZJ"
                With ws
                    .Range("B8").value = .Cells(i, "E").value
                    .Range("E8").value = .Cells(i, "F").value
                    .Range("E7").value = .Cells(i, "D").value
                End With
            Case "ZZS"
                ws.Range("B10").value = ws.Cells(i, "D").value
            Case "LSH"
                ws.Range("E9").value = ws.Cells(i, "D").value
        End Select
    Next i
    On Error GoTo 0
End Sub

Sub FormatWorksheet(ws As Worksheet, lastDataRow As Long)
    On Error Resume Next ' 捕获错误
    With ws
        ' 设置全局对齐方式
        .UsedRange.HorizontalAlignment = xlCenter
        .UsedRange.VerticalAlignment = xlCenter
        
        ' 设置数据范围 (A13到G最后一行)
        If lastDataRow >= 13 Then
            Dim dataRange As Range
            Set dataRange = .Range("A13:G" & lastDataRow)
            
            ' 清除所有边框
            dataRange.Borders.LineStyle = xlNone
            
            ' 设置外边框
            dataRange.Borders.LineStyle = xlContinuous
            
            ' 清除空行的边框
            Dim rowIndex As Long
            For rowIndex = 13 To lastDataRow
                If IsEmpty(.Cells(rowIndex, "A").value) Then
                    .Rows(rowIndex).Borders.LineStyle = xlNone
                End If
            Next rowIndex
            
            ' 设置数字格式
            .Range("C13:C" & lastDataRow).NumberFormat = "0.0"
            .Range("D13:D" & lastDataRow).NumberFormat = "0.000"
            .Range("E13:F" & lastDataRow).NumberFormat = "0.000000"
        End If
        
        ' 特殊单元格格式
        .Range("B9").NumberFormat = "0.0000"
        .Range("B10").NumberFormat = "0.000"
        .Range("E7,E9").NumberFormat = "0.000"
        .Range("B8").NumberFormat = "0.000000"
        .Range("E8").NumberFormat = "0.000000"
    End With
    On Error GoTo 0
End Sub















