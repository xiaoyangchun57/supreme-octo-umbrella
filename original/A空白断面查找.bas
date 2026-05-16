Attribute VB_Name = "A空白断面查找"
Sub ProcessFiles()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 创建新工作簿
    Dim resultWb As Workbook
    Dim resultWs As Worksheet
    Set resultWb = Workbooks.Add
    Set resultWs = resultWb.Sheets(1)
    
    ' 设置标题
    resultWs.Range("A1").value = "文件名"
    resultWs.Range("B1").value = "数据1"
    resultWs.Range("C1").value = "数据2"
    Dim outputRow As Long
    outputRow = 2
    
    ' 选择文件夹
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "选择要处理的文件夹"
        If .Show = -1 Then folderPath = .SelectedItems(1) Else Exit Sub
    End With
    
    ' 递归处理文件夹
    ProcessFolder folderPath, resultWs, outputRow
    
    ' 自动调整列宽
    resultWs.Columns("A:C").AutoFit
    
    ' 保存结果
    Dim savePath As String
    savePath = folderPath & "\汇总结果_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
    resultWb.SaveAs savePath
    resultWb.Close
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "处理完成！结果已保存至：" & vbCrLf & savePath, vbInformation
End Sub

Sub ProcessFolder(folderPath As String, resultWs As Worksheet, ByRef outputRow As Long)
    Dim fso As Object, folder As Object, subFolder As Object
    Dim file As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' 处理当前文件夹中的文件
    For Each file In folder.files
        If InStr(file.name, ".") > 0 Then  ' 确保是文件
            Dim ext As String
            ext = LCase(Right(file.name, Len(file.name) - InStrRev(file.name, ".")))
            
            ' 检查是否为Excel文件
            If ext = "xlsx" Or ext = "xls" Or ext = "xlsm" Then
                ' 检查文件名关键字
                Dim fileName As String
                fileName = LCase(file.name)
                
                ' 处理横断面文件 (B/Q/K/D/横断面)
                If InStr(fileName, "b") > 0 Or InStr(fileName, "q") > 0 Or _
                   InStr(fileName, "k") > 0 Or InStr(fileName, "d") > 0 Or _
                   InStr(fileName, "横断面") > 0 Then
                    Set wb = Workbooks.Open(file.path, ReadOnly:=True)
                    Set ws = wb.Sheets(1)
                    CopyData ws, "E13", "F", resultWs, outputRow, file.name
                    wb.Close False
                
                ' 处理纵断面文件 (Z/纵断面)
                ElseIf InStr(fileName, "z") > 0 Or InStr(fileName, "纵断面") > 0 Then
                    Set wb = Workbooks.Open(file.path, ReadOnly:=True)
                    Set ws = wb.Sheets(1)
                    CopyData ws, "G11", "H", resultWs, outputRow, file.name
                    wb.Close False
                End If
            End If
        End If
    Next file
    
    ' 递归处理子文件夹
    For Each subFolder In folder.subFolders
        ProcessFolder subFolder.path, resultWs, outputRow
    Next subFolder
End Sub

Sub CopyData(sourceWs As Worksheet, startCell As String, endCol As String, _
             resultWs As Worksheet, ByRef outputRow As Long, fileName As String)
    On Error Resume Next
    Dim lastRow As Long
    Dim sourceRng As Range
    Dim fileBaseName As String
    
    ' 获取文件名（不含扩展名）
    fileBaseName = Left(fileName, InStrRev(fileName, ".") - 1)
    
    ' 查找数据结束行
    With sourceWs
        lastRow = .Range(startCell).Offset(0, 1).End(xlDown).row
        If lastRow < .Range(startCell).row Then lastRow = .Cells(.Rows.count, endCol).End(xlUp).row
    End With
    
    ' 检查有效数据范围
    If lastRow >= sourceWs.Range(startCell).row Then
        ' 设置复制范围
        Set sourceRng = sourceWs.Range( _
            startCell & ":" & endCol & lastRow)
        
        ' 复制数据到结果表
        resultWs.Cells(outputRow, "A").value = fileBaseName
        sourceRng.Copy
        resultWs.Cells(outputRow, "B").PasteSpecial Paste:=xlPasteValues
        
        ' 更新输出行位置（添加空行分隔）
        outputRow = outputRow + sourceRng.Rows.count + 1
    End If
    
    Application.CutCopyMode = False
End Sub

