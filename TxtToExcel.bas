Sub 文字檔轉Excel()
    ' 宣告變數
    Dim FileNum As Integer
    Dim TextLine As String
    Dim CurrentRow As Long
    Dim hostName As String
    Dim macAddress As String
    Dim ipAddress As String
    Dim filePath As String
    Dim startPos As Long, endPos As Long
    Dim TextLines() As String
    Dim i As Long
    
    ' 使用檔案選擇對話框
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "請選擇要讀取的文字檔"
        .Filters.Clear
        .Filters.Add "文字檔", "*.txt"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            MsgBox "未選擇檔案，程式結束"
            Exit Sub
        End If
    End With
    
    ' 清除現有資料
    Cells.Clear
    
    ' 設定起始行
    CurrentRow = 2  ' 第一行保留給標題
    
    ' 讀取整個文字檔
    Open filePath For Input As #1
    TextLine = Input$(LOF(1), #1)
    Close #1
    
    ' 分割文字成陣列
    TextLines = Split(TextLine, vbLf)
    
    ' 設定標題
    Cells(1, 1) = "主機名稱"
    Cells(1, 2) = "MAC地址"
    Cells(1, 3) = "IP地址"
    
    ' 處理每一行
    For i = 0 To UBound(TextLines)
        TextLine = Trim(TextLines(i))
        
        If InStr(TextLine, "host ") > 0 Then
            ' 提取主機名稱
            startPos = InStr(TextLine, "host ") + 5
            endPos = InStr(startPos, TextLine, " {") - 1
            hostName = Mid(TextLine, startPos, endPos - startPos + 1)
            
            ' 提取MAC地址
            startPos = InStr(TextLine, "ethernet ") + 9
            endPos = InStr(startPos, TextLine, ";") - 1
            macAddress = Mid(TextLine, startPos, endPos - startPos + 1)
            
            ' 提取IP地址
            startPos = InStr(TextLine, "fixed-address ") + 13
            endPos = InStr(startPos, TextLine, ";") - 1
            ipAddress = Mid(TextLine, startPos, endPos - startPos + 1)
            
            ' 寫入Excel
            Cells(CurrentRow, 1) = Trim(hostName)
            Cells(CurrentRow, 2) = Trim(macAddress)
            Cells(CurrentRow, 3) = Trim(ipAddress)
            
            ' 移至下一行
            CurrentRow = CurrentRow + 1
        End If
    Next i
    
    ' 自動調整欄寬
    Columns("A:C").AutoFit
    
    ' 設定表格樣式
    Range("A1:C1").Font.Bold = True
    Range("A1:C" & CurrentRow - 1).Borders.LineStyle = xlContinuous
    Range("A1:C" & CurrentRow - 1).HorizontalAlignment = xlLeft
    
    MsgBox "轉換完成！共處理 " & (CurrentRow - 2) & " 筆資料"
End Sub 