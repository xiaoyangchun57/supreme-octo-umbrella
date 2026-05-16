Attribute VB_Name = "A成果表转85"
Sub ProcessFiles()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' 1. 让用户选择根文件夹
    Dim rootFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "请选择包含数据的根文件夹"
        If .Show = -1 Then rootFolder = .SelectedItems(1) Else Exit Sub
    End With
    
    ' 2. 查找"对应表.xlsx"文件
    Dim refBook As Workbook
    Dim refPath As String
    refPath = FindRefFile(rootFolder, "对应表.xlsx")
    
    If refPath = "" Then
        MsgBox "未找到'对应表.xlsx'文件", vbExclamation
        Exit Sub
    End If
    
    ' 3. 打开对应表并读取Sheet3数据
    Set refBook = Workbooks.Open(refPath)
    Dim refSheet As Worksheet
    Set refSheet = refBook.Sheets("Sheet3")
    
    Dim refData() As Variant
    refData = refSheet.UsedRange.value
    
    ' 4. 递归处理所有子文件夹
    ProcessFolder rootFolder, refData, refBook.name
    
    ' 5. 清理
    refBook.Close SaveChanges:=False
    MsgBox "处理完成!", vbInformation

ExitHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandler:
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

Function FindRefFile(ByVal folderPath As String, ByVal fileName As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object, file As Object
    
    ' 检查当前文件夹
    For Each file In fso.GetFolder(folderPath).files
        If LCase(file.name) = LCase(fileName) Then
            FindRefFile = file.path
            Exit Function
        End If
    Next
    
    ' 递归检查子文件夹
    For Each folder In fso.GetFolder(folderPath).subFolders
        Dim result As String
        result = FindRefFile(folder.path, fileName)
        If result <> "" Then
            FindRefFile = result
            Exit Function
        End If
    Next
End Function

Sub ProcessFolder(ByVal folderPath As String, refData() As Variant, refBookName As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object, folder As Object
    Dim targetBook As Workbook
    
    ' 处理当前文件夹中的文件
    For Each file In fso.GetFolder(folderPath).files
        If LCase(Right(file.name, 5)) = ".xlsx" And LCase(file.name) <> LCase(refBookName) Then
            Dim searchName As String
            searchName = Left(file.name, InStrRev(file.name, ".") - 1)  ' 移除扩展名
            
            ' 判断文件名特征
            Dim hasHeng As Boolean
            Dim hasZong As Boolean
            hasHeng = (InStr(1, searchName, "横", vbTextCompare) > 0)
            hasZong = (InStr(1, searchName, "纵", vbTextCompare) > 0)
            
            ' 确定处理类型
            Dim processType As String
            If hasHeng Then
                processType = "不含Z"  ' 横字优先于纵字
            ElseIf hasZong Then
                processType = "含Z"
            Else
                ' 无汉字时按原Z标志判断
                processType = IIf(InStr(1, searchName, "Z", vbTextCompare) > 0, "含Z", "不含Z")
            End If
            
            ' 在Sheet3中查找匹配项
            Dim iRow As Long, iCol As Long
            Dim found As Boolean: found = False
            Dim subtractValue As Variant
            
            For iRow = 2 To UBound(refData, 1)
                For iCol = 1 To UBound(refData, 2)
                    If Not IsEmpty(refData(iRow, iCol)) Then
                        If refData(iRow, iCol) = searchName Then
                            found = True
                            subtractValue = refData(1, iCol)
                            Exit For
                        End If
                    End If
                Next
                If found Then Exit For
            Next
            
            ' 找到匹配文件则处理
            If found Then
                Set targetBook = Workbooks.Open(file.path)
                Dim targetSheet As Worksheet
                Set targetSheet = targetBook.Sheets(1)  ' 假设操作第一个工作表
                
                ' 根据处理类型执行操作
                Select Case processType
                    Case "含Z"
                        ' 含Z/纵字的处理
                        ProcessColumn targetSheet, "E", 11, subtractValue
                        ProcessColumn targetSheet, "F", 11, subtractValue
                        ProcessSingleCell targetSheet, "D6", subtractValue
                    Case "不含Z"
                        ' 不含Z/横字的处理
                        ProcessColumn targetSheet, "D", 13, subtractValue
                        ProcessSingleCellIfValue targetSheet, "E7", subtractValue
                        ProcessSingleCellIfValue targetSheet, "E9", subtractValue
                        ProcessSingleCellIfValue targetSheet, "B10", subtractValue
                End Select
                
                targetBook.Close SaveChanges:=True
            End If
        End If
    Next
    
    ' 递归处理子文件夹
    For Each folder In fso.GetFolder(folderPath).subFolders
        ProcessFolder folder.path, refData, refBookName
    Next
End Sub

Sub ProcessColumn(ws As Worksheet, colLetter As String, startRow As Long, subtractValue As Variant)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, colLetter).End(xlUp).row
    
    ' 确保有数据需要处理
    If lastRow >= startRow Then
        Dim rng As Range
        Set rng = ws.Range(colLetter & startRow & ":" & colLetter & lastRow)
        
        ' 执行减法操作
        Dim cell As Range
        For Each cell In rng
            If IsNumeric(cell.value) And IsNumeric(subtractValue) Then
                cell.value = cell.value - subtractValue
            End If
        Next
    End If
End Sub

' 处理单个单元格的减法操作
Sub ProcessSingleCell(ws As Worksheet, cellAddress As String, subtractValue As Variant)
    Dim cell As Range
    Set cell = ws.Range(cellAddress)
    
    If Not cell Is Nothing Then
        If IsNumeric(cell.value) And IsNumeric(subtractValue) Then
            cell.value = cell.value - subtractValue
        End If
    End If
End Sub

' 仅当单元格有值时才执行减法操作
Sub ProcessSingleCellIfValue(ws As Worksheet, cellAddress As String, subtractValue As Variant)
    Dim cell As Range
    Set cell = ws.Range(cellAddress)
    
    If Not cell Is Nothing Then
        ' 检查单元格是否有值（非空且为数值）
        If Not IsEmpty(cell.value) And IsNumeric(cell.value) And IsNumeric(subtractValue) Then
            cell.value = cell.value - subtractValue
        End If
    End If
End Sub
