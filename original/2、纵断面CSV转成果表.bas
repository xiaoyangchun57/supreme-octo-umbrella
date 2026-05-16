Attribute VB_Name = "CSV转成果表纵断面最终版"

Option Explicit

' 声明进度窗体变量
Public ProgressForm As UserForm1 ' 直接声明为具体的窗体类型

Public Sub UpdateProgressBar(fileName As String, folderPath As String, current As Long, total As Long)
    If ProgressForm Is Nothing Then
        Set ProgressForm = New UserForm1
        ProgressForm.caption = "文件处理进度"
        ProgressForm.Show vbModeless
    End If
    
    With ProgressForm
        ' 更新文件名
        .lblFileName.caption = "当前文件: " & fileName
        .lblFolder.caption = "文件夹: " & folderPath
        
        ' 更新进度计数
        .lblCount.caption = "进度: " & current & " / " & total & " (" & Format(current / total, "0.0%") & ")"
        
        ' 更新进度条
        Dim progressWidth As Long
        progressWidth = (current / total) * (.Frame1.width - 4) ' 减去边框宽度
        If progressWidth < 0 Then progressWidth = 0
        If progressWidth > (.Frame1.width - 4) Then progressWidth = .Frame1.width - 4
        .Progress.width = progressWidth
        
        ' 更新处理速度
        Static lastUpdate As Double, lastCount As Long
        Dim elapsedTime As Double
        If lastUpdate = 0 Then lastUpdate = Timer
        elapsedTime = Timer - lastUpdate
        
        If elapsedTime > 1 Then ' 至少1秒更新一次
            If lastCount > 0 Then
                Dim speed As Double
                speed = (current - lastCount) / elapsedTime
                .lblSpeed.caption = "速度: " & Format(speed, "0.0") & " 文件/秒"
                
                ' 估计剩余时间
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

' 辅助函数：格式化时间
Private Function FormatTime(seconds As Double) As String
    If seconds <= 0 Then
        FormatTime = "00:00"
        Exit Function
    End If
    
    Dim minutes As Long, hours As Long
    minutes = Int(seconds / 60)
    seconds = seconds - minutes * 60
    hours = Int(minutes / 60)
    minutes = minutes - hours * 60
    
    FormatTime = Format(hours, "00") & ":" & Format(minutes, "00") & ":" & Format(Int(seconds), "00")
End Function

Sub B2()
    Dim fso As Object
    Dim rootFolder As String ' 用户选择的根文件夹
    Dim templatePath As String ' 模板文件路径（根文件夹下的“纵断面模板.xlsx”）
    Dim totalCount As Long ' 含Z的CSV文件总数
    Dim progressCount As Long ' 已处理文件数
    
    ' 初始化文件系统对象
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 1. 选择根文件夹（需包含“纵断面模板.xlsx”）
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "请选择包含CSV文件和纵断面模板的根文件夹"
        If .Show = -1 Then
            rootFolder = .SelectedItems(1)
        Else
            Exit Sub ' 用户取消，退出
        End If
    End With
    
    ' 2. 检查根文件夹下是否有“纵断面模板.xlsx”（必须存在）
    templatePath = fso.BuildPath(rootFolder, "纵断面模板.xlsx")
    If Not fso.FileExists(templatePath) Then
        MsgBox "根文件夹下未找到'纵断面模板.xlsx'，无法继续！", vbExclamation
        Exit Sub
    End If
    
    ' 3. 统计含Z的CSV文件总数
    totalCount = CountCSVFiles(fso, rootFolder)
    If totalCount = 0 Then
        MsgBox "未找到含Z的CSV文件！", vbExclamation
        Exit Sub
    End If
    
    ' 4. 初始化进度窗体
    If ProgressForm Is Nothing Then
        Set ProgressForm = New UserForm1
        ProgressForm.caption = "文件处理进度"
        ProgressForm.Show vbModeless
    End If
    UpdateProgressBar "正在初始化...", rootFolder, 0, totalCount
    
    ' 5. 禁用Excel冗余功能（提升速度）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 6. 调用处理函数（传递模板路径）
    progressCount = 0
   CloneTemplateForCSV fso, rootFolder, totalCount, progressCount, templatePath
   ' +++ 新增关键步骤 +++
    ' 步骤2: 数据处理（原缺失的调用）
    progressCount = 0  ' 重置进度计数器
    ProcessCSVFiles fso, rootFolder, totalCount, progressCount
    
    ' 步骤3: 恢复Excel设置
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' 7. 恢复Excel功能
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' 8. 清理进度窗体+提示完成
    If Not ProgressForm Is Nothing Then
        Unload ProgressForm
        Set ProgressForm = Nothing
    End If
    MsgBox "处理完成! 共处理 " & progressCount & " 个文件。", vbInformation
End Sub
' 函数声明：新增`templatePath`参数（接收模板文件路径）
' 说明：所有调用该函数的地方都需要传递`templatePath`
Sub CloneTemplateForCSV( _
    fso As Object, _
    folderPath As String, _
    totalCount As Long, _
    ByRef progressCount As Long, _
    templatePath As String _
)
    Dim folder As Object ' 当前文件夹对象
    Dim subFolder As Object ' 子文件夹对象
    Dim file As Object ' 当前文件对象
    Dim csvPath As String ' CSV文件路径
    Dim xlsxPath As String ' 目标XLSX路径（与CSV同名）
    Dim baseName As String ' CSV文件名（不含扩展名）
    ' ... 其他变量（如wbCSV、wsXLSX等）保持不变 ...
    
    ' 获取当前文件夹对象
    Set folder = fso.GetFolder(folderPath)
    
    ' 遍历当前文件夹中的所有文件
    For Each file In folder.files
        ' 仅处理CSV文件
        If LCase(fso.GetExtensionName(file.name)) = "csv" Then
            csvPath = file.path ' 获取CSV路径
            baseName = fso.GetBaseName(csvPath) ' 获取CSV文件名（不含扩展名）
            xlsxPath = fso.BuildPath(folderPath, baseName & ".xlsx") ' 构建目标XLSX路径
            
            ' 条件1：CSV文件名含“Z”（不区分大小写）
            If InStr(1, baseName, "Z", vbTextCompare) = 0 Then
                GoTo NextFile ' 不含“Z”，跳过
            End If
            
            ' 条件2：复制模板文件（若目标XLSX不存在）
            If Not fso.FileExists(xlsxPath) Then
                On Error Resume Next
                ' 复制模板（从根文件夹到当前CSV所在文件夹）
                fso.CopyFile templatePath, xlsxPath, OverwriteFiles:=True ' 覆盖已存在的文件（可选）
                If Err.Number <> 0 Then
                    Debug.Print "复制模板失败：" & templatePath & " → " & xlsxPath
                    Err.Clear
                    GoTo NextFile ' 复制失败，跳过
                End If
                Debug.Print "已复制模板：" & xlsxPath ' 调试信息（可删除）
            End If
            
            ' ? 所有条件满足，开始处理
            progressCount = progressCount + 1 ' 递增已处理数
            UpdateProgressBar file.name, folderPath, progressCount, totalCount ' 更新进度
            
            ' 后续处理逻辑（打开CSV/XLSX、转换坐标、填入数据等）保持不变...
            ' （此处省略原代码中的“打开文件、转换坐标、填入数据”部分，需保留）
        End If
NextFile:
    Next file
    
    ' 递归处理子文件夹（传递模板路径）
    For Each subFolder In folder.subFolders
        CloneTemplateForCSV fso, subFolder.path, totalCount, progressCount, templatePath  ' 修改调用名
    Next subFolder
End Sub

' 递归处理CSV文件
Sub ProcessCSVFiles(fso As Object, folderPath As String, totalCount As Long, ByRef progressCount As Long)
    Dim folder As Object, subFolder As Object, file As Object
    Dim csvPath As String, xlsxPath As String, baseName As String
    Dim wbCSV As Workbook, wbXLSX As Workbook
    Dim wsCSV As Worksheet, wsXLSX As Worksheet
    Dim dataArray As Variant, sortedArray As Variant
    Dim lastRow As Long, i As Long, j As Long
    Dim startDistance As Double
    Dim lastXRow As Long
    Dim startRow As Long
    
    Set folder = fso.GetFolder(folderPath)
    
   For Each file In folder.files
    If LCase(fso.GetExtensionName(file.name)) = "csv" Then
        csvPath = file.path
        baseName = fso.GetBaseName(csvPath)
        xlsxPath = fso.BuildPath(folderPath, baseName & ".xlsx")
        
        ' 条件1：文件名包含“Z”
        If InStr(1, baseName, "Z", vbTextCompare) = 0 Then
            GoTo NextFile
        End If
        
        ' 条件2：对应XLSX存在
        If Not fso.FileExists(xlsxPath) Then
            Debug.Print "找不到对应的XLSX文件: " & xlsxPath
            GoTo NextFile
        End If
        
        ' 条件3：能打开CSV
        On Error Resume Next
        Set wbCSV = Workbooks.Open(csvPath)
        If Err.Number <> 0 Then
            Debug.Print "无法打开CSV文件: " & csvPath
            Err.Clear
            GoTo NextFile
        End If
        On Error GoTo 0
        Set wsCSV = wbCSV.Sheets(1)
        
        ' 条件4：能打开XLSX
        On Error Resume Next
        Set wbXLSX = Workbooks.Open(xlsxPath)
        If Err.Number <> 0 Then
            Debug.Print "无法打开XLSX文件: " & xlsxPath
            Err.Clear
            wbCSV.Close SaveChanges:=False
            GoTo NextFile
        End If
        On Error GoTo 0
        
        ' 条件5：XLSX有“沟道纵断面成果表”工作表
        On Error Resume Next
        Set wsXLSX = wbXLSX.Sheets("沟道纵断面成果表")
        If Err.Number <> 0 Then
            Debug.Print "XLSX无'沟道纵断面成果表'工作表: " & xlsxPath
            Err.Clear
            wbCSV.Close SaveChanges:=False
            wbXLSX.Close SaveChanges:=False
            GoTo NextFile
        End If
        On Error GoTo 0
        
        ' ? 所有条件满足，递增进度
        progressCount = progressCount + 1
        UpdateProgressBar file.name, folderPath, progressCount, totalCount
            
            lastRow = wsCSV.Cells(wsCSV.Rows.count, "A").End(xlUp).row
            
              '  ??? 修改点1：使用新的高斯反算函数转换起始点坐标 ???
            Dim Longitude As Double, Latitude As Double
            Dim centralMeridian As Double
            centralMeridian = 117 ' 中央子午线117°
            GaussProjInv wsCSV.Range("B1").value, wsCSV.Range("C1").value, centralMeridian, Latitude, Longitude
            
            wsXLSX.Range("D5").value = Round(Longitude, 6)
            wsXLSX.Range("B6").value = Round(Latitude, 6)
            wsXLSX.Range("D6").value = wsCSV.Range("D1").value
            
            If lastRow < 2 Then
                wbCSV.Close SaveChanges:=False
                wbXLSX.Close SaveChanges:=False
                GoTo NextFile
            End If
            
            Dim d2Value As Double, dLastValue As Double
            d2Value = wsCSV.Range("D2").value
            dLastValue = wsCSV.Cells(lastRow, "D").value
            
            dataArray = wsCSV.Range("A2:E" & lastRow).value
            
            If dLastValue > d2Value Then
                ReDim sortedArray(1 To UBound(dataArray), 1 To UBound(dataArray, 2))
                For i = 1 To UBound(dataArray)
                    For j = 1 To UBound(dataArray, 2)
                        sortedArray(i, j) = dataArray(UBound(dataArray) - i + 1, j)
                    Next j
                Next i
                dataArray = sortedArray
            End If
            
            Dim startPointY As Double, startPointX As Double
            startPointY = dataArray(1, 2)
            startPointX = dataArray(1, 3)
            
            Dim directions() As Double
            ReDim directions(1 To UBound(dataArray))
            directions(1) = 0
            
            For i = 2 To UBound(dataArray)
                Dim deltaY As Double, deltaX As Double
                deltaY = dataArray(i, 2) - dataArray(i - 1, 2)
                deltaX = dataArray(i, 3) - dataArray(i - 1, 3)
                
                directions(i) = (Atn2(deltaY, deltaX) * 180 / Application.PI())
                
                If directions(i) < 0 Then directions(i) = directions(i) + 360
            Next i
            
            ' ====== 修改的缺失数据处理部分 ======
            Dim eFixed() As Variant
            ReDim eFixed(1 To UBound(dataArray))
            Dim prevValid As Double, nextValid As Double
            Dim missingStartIndex As Long, missingEndIndex As Long
            
            For i = 1 To UBound(dataArray)
                If Not IsEmpty(dataArray(i, 5)) Then
                    eFixed(i) = dataArray(i, 5)
                Else
                    eFixed(i) = Empty
                End If
            Next i
            
            i = 1
            Do While i <= UBound(dataArray)
                If IsEmpty(eFixed(i)) Then
                    missingStartIndex = i
                    Do While i <= UBound(dataArray) And IsEmpty(eFixed(i))
                        i = i + 1
                    Loop
                    missingEndIndex = i - 1
                    Dim missingCount As Long
                    missingCount = missingEndIndex - missingStartIndex + 1
                    
                    ' 获取缺失段前后的有效值
                    If missingStartIndex > 1 Then
                        prevValid = eFixed(missingStartIndex - 1)
                    Else
                        ' 开头缺失，使用下一有效值作为参考
                        prevValid = eFixed(missingEndIndex + 1)
                    End If
                    
                    If missingEndIndex < UBound(dataArray) Then
                        nextValid = eFixed(missingEndIndex + 1)
                    Else
                        ' 结尾缺失，使用前一有效值作为参考
                        nextValid = eFixed(missingStartIndex - 1)
                    End If
                    
                    ' 计算总差值和平均步长
                    Dim diff As Double, avgStep As Double
                    diff = nextValid - prevValid
                    avgStep = diff / (missingCount + 1)
                    
                    ' 计算当前值并添加随机波动
                    Dim currentValue As Double
                    currentValue = prevValid
                    For j = missingStartIndex To missingEndIndex
                        ' 每一步加上平均步长
                        currentValue = currentValue + avgStep
                        
                        ' 添加随机波动 (-0.5到0.5之间，不超过0.501)
                        Dim randOffset As Double
                        randOffset = (Rnd() - 0.5) * 1#  ' 范围 ±0.5
                        If randOffset > 0.501 Then randOffset = 0.501
                        If randOffset < -0.501 Then randOffset = -0.501
                        
                        ' 更新当前值并确保在合理范围内
                        currentValue = currentValue + randOffset
                        
                        ' 确保值在prevValid和nextValid之间
                        If currentValue > nextValid Then currentValue = nextValid - 0.01
                        If currentValue < prevValid Then currentValue = prevValid + 0.01
                        
                        eFixed(j) = currentValue
                    Next j
                Else
                    i = i + 1
                End If
            Loop
            ' ====== 缺失数据处理结束 ======
            
            ' ★★★★ 新增代码位置 - 确保E列大于D列且递减 ★★★★
            EnsureEGreaterThanDAndDecreasing dataArray, eFixed
            
            startRow = 11
            wsXLSX.Range("A" & startRow & ":H" & wsXLSX.Rows.count).ClearContents
            
            startRow = 11
            wsXLSX.Range("A" & startRow & ":H" & wsXLSX.Rows.count).ClearContents
            
            For i = 1 To UBound(dataArray)
    wsXLSX.Cells(startRow + i - 1, 1).value = i
    
    If i = 1 Then
        wsXLSX.Cells(startRow + i - 1, 2).value = "河心"
    Else
        wsXLSX.Cells(startRow + i - 1, 2).value = "河心" & (i - 1)
    End If
    
    ' === 修改起点距计算开始 ===
    If i = 1 Then
        ' 起点距离为0
        wsXLSX.Cells(startRow + i - 1, 3).value = 0
    Else
        ' 点与点平距计算
        Dim pointDistance As Double
        pointDistance = Sqr((dataArray(i, 3) - dataArray(i - 1, 3)) ^ 2 + _
                        (dataArray(i, 2) - dataArray(i - 1, 2)) ^ 2)
        wsXLSX.Cells(startRow + i - 1, 3).value = Round(pointDistance, 1)
    End If
    ' === 修改起点距计算结束 ===
    
    wsXLSX.Cells(startRow + i - 1, 4).value = Round(directions(i), 1)
    wsXLSX.Cells(startRow + i - 1, 5).value = dataArray(i, 4)
    wsXLSX.Cells(startRow + i - 1, 6).value = eFixed(i)
    
    GaussProjInv dataArray(i, 2), dataArray(i, 3), centralMeridian, Latitude, Longitude
    wsXLSX.Cells(startRow + i - 1, 7).value = Round(Longitude, 6)
    wsXLSX.Cells(startRow + i - 1, 8).value = Round(Latitude, 6)
Next i
            
            lastXRow = startRow + UBound(dataArray) - 1
            Dim dataRange As Range
            Set dataRange = wsXLSX.Range("A" & startRow & ":H" & lastXRow)
            
            dataRange.HorizontalAlignment = xlCenter
            dataRange.VerticalAlignment = xlCenter
            
            With dataRange
                .Columns("C").NumberFormat = "0.0"
                .Columns("D").NumberFormat = "0"
                .Columns("E:E").NumberFormat = "0.000"
                .Columns("F:F").NumberFormat = "0.000"
                .Columns("G:H").NumberFormat = "0.000000"
            End With
            
            dataRange.Borders.LineStyle = xlContinuous
            
            wbXLSX.Close SaveChanges:=True
            wbCSV.Close SaveChanges:=False
        End If
NextFile:
    Next file
    
    For Each subFolder In folder.subFolders
        ProcessCSVFiles fso, subFolder.path, totalCount, progressCount
    Next subFolder
End Sub

' Atan2函数计算
Private Function Atn2(y As Double, x As Double) As Double
    If x > 0 Then
        Atn2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then
            Atn2 = Atn(y / x) + Application.PI()
        Else
            Atn2 = Atn(y / x) - Application.PI()
        End If
    Else
        If y > 0 Then
            Atn2 = Application.PI() / 2
        ElseIf y < 0 Then
            Atn2 = -Application.PI() / 2
        Else
            Atn2 = 0
        End If
    End If
End Function
' 修改GaussProjInv函数的参数声明（仅修改前两个参数）
Sub GaussProjInv(ByVal x As Double, ByVal y As Double, centralMeridian As Double, ByRef B As Double, ByRef L As Double)
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

' 统计根文件夹及其子文件夹下含Z的CSV文件总数（原逻辑不变）
Function CountCSVFiles(fso As Object, folderPath As String) As Long
    Dim folder As Object, subFolder As Object, file As Object
    Dim count As Long
    Dim baseName As String
    
    Set folder = fso.GetFolder(folderPath)
    
    ' 统计当前文件夹中的CSV文件
    For Each file In folder.files
        If LCase(fso.GetExtensionName(file.name)) = "csv" Then
            baseName = fso.GetBaseName(file.name)
            If InStr(1, baseName, "Z", vbTextCompare) > 0 Then
                count = count + 1
            End If
        End If
    Next file
    
    ' 递归统计子文件夹中的CSV文件
    For Each subFolder In folder.subFolders
        count = count + CountCSVFiles(fso, subFolder.path)
    Next subFolder
    
    CountCSVFiles = count
End Function

' 确保E列值大于D列值且符合顺序的函数
Private Sub EnsureEGreaterThanDAndDecreasing(dataArray As Variant, eFixed() As Variant)
    Dim i As Long
    
    ' 第一遍修正：确保每个点E>D
    For i = 1 To UBound(dataArray)
        ' 当E列值缺失或不大于D列时进行修正
        If IsEmpty(eFixed(i)) Or eFixed(i) <= dataArray(i, 4) Then
            ' 确保修正值比D列至少大0.001
            eFixed(i) = dataArray(i, 4) + 0.001 + Rnd() * 0.5
        End If
    Next i
    
    ' 第二遍修正：确保整体递减趋势
    For i = 2 To UBound(eFixed)
        ' 确保当前点小于前一点且大于D列值
        If eFixed(i) >= eFixed(i - 1) Then
            ' 设置为前一点减去随机递减量
            Dim declineAmount As Double
            declineAmount = 0.001 + Rnd() * 0.5 ' 随机递减幅度
            eFixed(i) = eFixed(i - 1) - declineAmount
        End If
        
        ' 双重检查确保大于D列值
        If eFixed(i) <= dataArray(i, 4) Then
            eFixed(i) = dataArray(i, 4) + 0.001 + Rnd() * 0.5
        End If
    Next i
End Sub

' 判断数据是否从大到小排列（即递减）
Private Function IsDescendingOrder(dataArray As Variant) As Boolean
    ' 如果数组只有一行，无法判断，默认为True（递减）
    If UBound(dataArray) < 2 Then
        IsDescendingOrder = True
        Exit Function
    End If
    
    ' 获取第一行和最后一行的D列值（第4列）
    Dim firstValue As Double, lastValue As Double
    firstValue = dataArray(1, 4)
    lastValue = dataArray(UBound(dataArray), 4)
    
    ' 如果第一行的值大于最后一行的值，则为递减
    IsDescendingOrder = (firstValue > lastValue)
End Function












