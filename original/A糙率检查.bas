Attribute VB_Name = "A糙率检查"
Sub CheckEmptyG13Recursive()
    Dim fso As Object
    Dim fDialog As FileDialog
    Dim folderPath As String
    Dim logSheet As Worksheet
    Dim logRow As Long
    Dim fileCount As Long
    Dim emptyCount As Long
    Dim startTime As Double
    
    ' 记录开始时间
    startTime = Timer
    
    ' 设置文件对话框
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    fDialog.title = "选择包含Excel文件的文件夹"
    
    ' 显示对话框并获取文件夹路径
    If fDialog.Show = -1 Then
        folderPath = fDialog.SelectedItems(1)
        If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    Else
        MsgBox "未选择文件夹。操作已取消。", vbExclamation
        Exit Sub
    End If
    
    ' 创建日志工作表
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("日志").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set logSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    logSheet.name = "日志"
    logSheet.Range("A1").value = "文件路径"
    logSheet.Range("B1").value = "文件名"
    logSheet.Range("C1").value = "状态"
    logSheet.Range("A1:C1").Font.Bold = True
    
    logRow = 2
    fileCount = 0
    emptyCount = 0
    
    ' 创建文件系统对象
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 禁用Excel功能提高性能
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 递归处理文件夹
    ProcessFolder fso.GetFolder(folderPath), logSheet, logRow, fileCount, emptyCount
    
    ' 恢复Excel设置
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' 自动调整列宽
    logSheet.Columns("A:C").AutoFit
    
    ' 添加边框
    With logSheet.Range("A1:C" & logRow - 1)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' 添加筛选
    logSheet.Range("A1:C1").AutoFilter
    
    ' 显示结果
    Dim elapsedTime As Double
    elapsedTime = Round(Timer - startTime, 2)
    
    MsgBox "处理完成！" & vbCrLf & _
           "扫描文件夹: " & folderPath & vbCrLf & _
           "处理文件数: " & fileCount & vbCrLf & _
           "发现空G13文件数: " & emptyCount & vbCrLf & _
           "耗时: " & elapsedTime & " 秒", _
           vbInformation, "操作完成"
End Sub

' 递归处理文件夹
Private Sub ProcessFolder(folder As Object, logSheet As Worksheet, ByRef logRow As Long, ByRef fileCount As Long, ByRef emptyCount As Long)
    Dim file As Object
    Dim subFolder As Object
    Dim wb As Workbook
    Dim filePath As String
    
    ' 处理当前文件夹中的文件
    For Each file In folder.files
        ' 只处理Excel文件
        If LCase(Right(file.name, 5)) = ".xlsx" Or LCase(Right(file.name, 4)) = ".xls" Then
            fileCount = fileCount + 1
            filePath = file.path
            
            On Error Resume Next
            Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0)
            
            If Err.Number = 0 Then
                ' 检查G13是否为空
                If IsEmpty(wb.Sheets(1).Range("G13").value) Then
                    logSheet.Cells(logRow, 1).value = filePath
                    logSheet.Cells(logRow, 2).value = file.name
                    logSheet.Cells(logRow, 3).value = "G13为空"
                    logRow = logRow + 1
                    emptyCount = emptyCount + 1
                End If
                wb.Close False
            Else
                logSheet.Cells(logRow, 1).value = filePath
                logSheet.Cells(logRow, 2).value = file.name
                logSheet.Cells(logRow, 3).value = "无法打开文件"
                logRow = logRow + 1
                Err.Clear
            End If
            
            ' 每处理100个文件更新一次状态栏
            If fileCount Mod 100 = 0 Then
                Application.StatusBar = "已处理文件: " & fileCount & " | 发现空G13: " & emptyCount
            End If
        End If
    Next file
    
    ' 递归处理子文件夹
    For Each subFolder In folder.subFolders
        ProcessFolder subFolder, logSheet, logRow, fileCount, emptyCount
    Next subFolder
End Sub



