Attribute VB_Name = "模块1"
Sub InsertPerfectSizedPictures()
    ' 参数设置
    Const TARGET_COL As String = "I"    '
    Const START_ROW As Long = 3         ' 起始行
    Const MAX_IMAGE_HEIGHT As Long = 80 ' 最大图片高度（磅）
    Const PADDING_PERCENT As Double = 0.05 ' 边距百分比
    
    Dim fd As FileDialog
    Dim fileItem As Variant
    Dim selectedFiles As Collection
    Dim pic As Picture
    Dim i As Long
    Dim cell As Range
    Dim aspectRatio As Double
    Dim targetWidth As Double
    Dim targetHeight As Double
    
    ' 创建文件选择对话框
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Set selectedFiles = New Collection
    
    With fd
        .AllowMultiSelect = True
        .Title = "选择多张图片（按住Ctrl多选）"
        .Filters.Clear
        .Filters.Add "图片文件", "*.jpg;*.jpeg;*.png;*.bmp;*.gif"
        
        ' 显示对话框并收集文件
        If .Show = -1 Then
            For Each fileItem In .SelectedItems
                selectedFiles.Add fileItem
            Next fileItem
        Else
            MsgBox "操作已取消", vbInformation
            Exit Sub
        End If
    End With
    
    ' 检查是否选择了文件
    If selectedFiles.Count = 0 Then
        MsgBox "未选择任何图片", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False ' 关闭屏幕刷新
    
    ' 设置目标列宽（自动调整为最大宽度）
    Columns(TARGET_COL).ColumnWidth = 30 ' 初始宽度
    
    ' 循环插入每张图片
    For i = 1 To selectedFiles.Count
        Set cell = Cells(START_ROW + i - 1, TARGET_COL)
        
        ' 插入图片
        Set pic = ActiveSheet.Pictures.Insert(selectedFiles(i))
        
        With pic
            ' 获取原始宽高比
            aspectRatio = .Width / .Height
            
            ' 计算目标尺寸
            targetHeight = MAX_IMAGE_HEIGHT
            targetWidth = targetHeight * aspectRatio
            
            ' 设置图片位置和尺寸
            .Top = cell.Top + (cell.Height * PADDING_PERCENT)
            .Left = cell.Left + (cell.Width * PADDING_PERCENT)
            .Width = cell.Width * (1 - 2 * PADDING_PERCENT)
            .Height = .Width / aspectRatio ' 保持比例
            
            ' 如果高度超过限制，重新计算
            If .Height > MAX_IMAGE_HEIGHT Then
                .Height = MAX_IMAGE_HEIGHT
                .Width = MAX_IMAGE_HEIGHT * aspectRatio
                ' 水平居中
                .Left = cell.Left + (cell.Width - .Width) / 2
            End If
            
            ' 设置图片随单元格移动
            .Placement = xlMoveAndSize
            
            ' 调整行高以适应图片
            cell.RowHeight = .Height + (cell.Height * PADDING_PERCENT * 2)
        End With
    Next i
    
    ' 自动调整列宽（基于最大图片宽度）
    Dim maxWidth As Double
    maxWidth = 0
    For Each pic In ActiveSheet.Pictures
        If pic.Width > maxWidth Then maxWidth = pic.Width
    Next pic
    
    Columns(TARGET_COL).ColumnWidth = (maxWidth * (1 + 2 * PADDING_PERCENT)) * 0.75
    
    Application.ScreenUpdating = True ' 恢复屏幕刷新
    MsgBox "成功插入 " & selectedFiles.Count & " 张图片，已自动调整尺寸", vbInformation
End Sub
