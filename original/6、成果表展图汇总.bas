Attribute VB_Name = "A成果表展图汇总"
Sub ProcessFilesToTXT()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 选择文件夹
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "选择要处理的文件夹"
        If .Show = -1 Then folderPath = .SelectedItems(1) Else Exit Sub
    End With
    
    ' 创建输出文本文件路径
    Dim outputPath As String
    outputPath = folderPath & "\汇总结果_" & Format(Now, "yyyymmdd_hhmmss") & ".txt"
    
    ' 使用 ADODB.Stream 创建 UTF-8 文件
    Dim utfStream As Object
    Set utfStream = CreateObject("ADODB.Stream")
    utfStream.Type = 2 ' 文本类型
    utfStream.Charset = "utf-8"
    utfStream.Open
    
    ' 递归处理文件夹
    ProcessFolderToTXT folderPath, utfStream
    
    ' 保存 UTF-8 文件
    utfStream.SaveToFile outputPath, 2 ' 2 = 覆盖现有文件
    utfStream.Close
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "处理完成！结果已保存至：" & vbCrLf & outputPath, vbInformation
End Sub

Sub ProcessFolderToTXT(folderPath As String, outputStream As Object)
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
                Dim startCell As String
                Dim endCol As String
                Dim isProcessed As Boolean
                isProcessed = False
                
                ' 首先检查中文关键词（优先处理）
                If InStr(fileName, "沟道横断面成果表") > 0 Then
                    startCell = "E13"
                    endCol = "F"  ' 横断面处理
                    isProcessed = True
                ElseIf InStr(fileName, "沟道纵断面成果表") > 0 Then
                    startCell = "G11"
                    endCol = "H"  ' 纵断面处理
                    isProcessed = True
                ' 然后检查字母标识（不含中文关键词时使用）
                ElseIf InStr(fileName, "b") > 0 Or InStr(fileName, "q") > 0 Or _
                       InStr(fileName, "k") > 0 Or InStr(fileName, "d") > 0 Then
                    startCell = "E13"
                    endCol = "F"  ' 横断面处理
                    isProcessed = True
                ElseIf InStr(fileName, "z") > 0 Then
                    startCell = "G11"
                    endCol = "H"  ' 纵断面处理
                    isProcessed = True
                End If
                
                If isProcessed Then
                    On Error Resume Next
                    Set wb = Workbooks.Open(file.path, ReadOnly:=True, UpdateLinks:=0)
                    If Err.Number <> 0 Then
                        Debug.Print "无法打开文件: " & file.name
                    Else
                        Set ws = wb.Sheets(1)   ' 假设数据在第一个工作表
                        CopyDataToTXT ws, startCell, endCol, outputStream, file.name
                        wb.Close False
                        ' 添加空行分隔不同文件
                        outputStream.WriteText vbCrLf
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
    Next file
    
    ' 递归处理子文件夹
    For Each subFolder In folder.subFolders
        ProcessFolderToTXT subFolder.path, outputStream
    Next subFolder
End Sub

Sub CopyDataToTXT(sourceWs As Worksheet, startCell As String, endCol As String, _
             outputStream As Object, fileName As String)
    On Error Resume Next
    Dim lastRow As Long
    Dim fileBaseName As String
    Dim rowData As String
    Dim i As Long, j As Long
    
    ' 获取文件名（不含扩展名）
    fileBaseName = Left(fileName, InStrRev(fileName, ".") - 1)
    
    ' 查找数据结束行
    With sourceWs
        ' 确定起始行
        Dim startRow As Long
        startRow = .Range(startCell).row
        
        ' 结束行：从起始单元格所在列向下找最后一个非空行
        lastRow = .Cells(.Rows.count, endCol).End(xlUp).row
        
        ' 确保lastRow不小于起始行
        If lastRow < startRow Then lastRow = startRow
    End With
    
    ' 检查有效数据范围
    If lastRow >= sourceWs.Range(startCell).row Then
        ' 确定起始列和结束列
        Dim startCol As Long
        startCol = sourceWs.Range(startCell).Column
        Dim endColIndex As Long
        endColIndex = sourceWs.Range(endCol & "1").Column
        
        ' 遍历每行数据
        For i = sourceWs.Range(startCell).row To lastRow
            ' 检查整行是否为空
            Dim isEmptyRow As Boolean
            isEmptyRow = True
            
            ' 创建行数据（文件名开头）
            rowData = fileBaseName
            
            ' 遍历当前行的列
            For j = startCol To endColIndex
                Dim cellValue As String
                cellValue = Trim(CStr(sourceWs.Cells(i, j).value))
                
                ' 仅当单元格有值时添加
                If cellValue <> "" Then
                    rowData = rowData & "," & cellValue
                    isEmptyRow = False
                End If
            Next j
            
            ' 如果行不为空，写入数据
            If Not isEmptyRow Then
                outputStream.WriteText rowData & vbCrLf
            End If
        Next i
    End If
End Sub



