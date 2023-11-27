# 图片部分

## 所有图片居中

```vb
Sub PicCenter()
    Dim myS As InlineShape
    Application.ScreenUpdating = False
    For Each myS In ActiveDocument.InlineShapes
        If Len(myS.Range.Paragraphs(1).Range.Text) = 2 Then
            myS.Range.Paragraphs.Alignment = wdAlignParagraphCenter
        End If
    Next
    Application.ScreenUpdating = True
End Sub
```
当然，让我们逐句分析一下上述 VBA 代码：

1. `Sub PicCenter()`
    - 这是一个 VBA 宏的开始。`Sub` 表示子过程（或子程序），`PicCenter` 是宏的名称。您可以将其视为一个功能，它将执行一系列操作。
2. `Dim myS As InlineShape`
    - 这一行声明了一个名为 `myS` 的变量，类型为 `InlineShape`。`InlineShape` 是 Word 文档中的内联形状，通常用于表示图片、图表等。
3. `Application.ScreenUpdating = False`
    - 这一行关闭了屏幕更新，以提高宏的执行速度。在执行宏期间，屏幕不会刷新。
4. `For Each myS In ActiveDocument.InlineShapes`
    - 这是一个循环语句，用于遍历文档中的所有内联形状（即图片）。`myS` 将依次引用每个内联形状。
5. `If Len(myS.Range.Paragraphs(1).Range.Text) = 2 Then`
    - 这是一个条件语句。它检查当前内联形状的第一个段落的文本长度是否等于 2。如果是，说明这是一个图片（通常图片的文本长度很短）。
6. `myS.Range.Paragraphs.Alignment = wdAlignParagraphCenter`
    - 这一行将当前内联形状所在段落的对齐方式设置为居中对齐。
7. `Next`
    - 这表示循环的结束。它会继续遍历下一个内联形状，直到所有形状都被处理完。
8. `Application.ScreenUpdating = True`
    - 这一行重新打开屏幕更新，以便用户可以看到执行结果。
9. `End Sub`
    - 这是宏的结束。

总之，这段代码的作用是遍历文档中的所有图片，并将它们的对齐方式设置为居中。如果您有其他问题或需要进一步解释，请随时提问！🙂


## 图例居中：图...
```vb
Sub CenterSingleSentence()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If Left(para.Range.Text, 1) = "图" And para.Range.Sentences.Count = 1 Then
            para.Alignment = wdAlignParagraphCenter
        End If
    Next para
End Sub

```
当然，让我们逐句分析一下这段 VBA 代码：

```vb
Sub CenterSingleSentence()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If Left(para.Range.Text, 3) = "图 " And para.Range.Sentences.Count = 1 Then
            para.Alignment = wdAlignParagraphCenter
        End If
    Next para
End Sub
```

1. `Sub CenterSingleSentence()`
    - 这是一个 VBA 宏的开始。`Sub` 表示子过程（或子程序），`CenterSingleSentence` 是宏的名称。
2. `Dim para As Paragraph`
    - 这一行声明了一个名为 `para` 的变量，类型为 `Paragraph`。`Paragraph` 表示 Word 文档中的段落。
3. `For Each para In ActiveDocument.Paragraphs`
    - 这是一个循环语句，用于遍历文档中的所有段落。`para` 将依次引用每个段落。
4. `If Left(para.Range.Text, 3) = "图 " And para.Range.Sentences.Count = 1 Then`
    - 这是一个条件语句。它检查当前段落的文本是否以“图 ”开头，并且该段落只有一句话。
5. `para.Alignment = wdAlignParagraphCenter`
    - 这一行将当前段落的对齐方式设置为居中对齐。
6. `Next para`
    - 这表示循环的结束。它会继续遍历下一个段落，直到所有段落都被处理完。
7. `End Sub`
    - 这是宏的结束。

总之，这段代码的作用是遍历文档中的所有段落，如果某个段落以“图 ”开头并且只有一句话，就将其对齐方式设置为居中。如果您有其他问题或需要进一步解释，请随时提问！🙂

## 所有图片下方插入题注
（会把所有mathtype认为图片，有待改进）

```vb
Sub AddCaptionToAllImages()
    Dim oShp As InlineShape
    Dim i As Integer

    ' 遍历文档中的所有图片
    For Each oShp In ActiveDocument.InlineShapes
        oShp.Select
		Selection.InsertCaption Label:="图", Title:=" ", Position:=wdCaptionPositionBelow, ExcludeLabel:=0
    Next oShp
End Sub
```

# 标题部分
## VBA 设置标题字体

```vb
Sub SetHeadingStyleFont()
    With ActiveDocument.Styles(wdStyleHeading1).Font
        .Color = wdColorBlack
        .Bold = False
        .Size = 12
        .Name = "黑体"
    End With
    With ActiveDocument.Styles(wdStyleHeading2).Font
        .Color = wdColorBlack
        .Bold = False
        .Size = 12
        .Name = "黑体"
    End With
    With ActiveDocument.Styles(wdStyleHeading3).Font
        .Color = wdColorBlack
        .Bold = False
        .Size = 12
        .Name = "黑体"
    End With
    With ActiveDocument.Styles(wdStyleHeading4).Font
        .Color = wdColorBlack
        .Bold = False
        .Size = 12
        .Name = "黑体"
    End With
    With ActiveDocument.Styles(wdStyleHeading5).Font
        .Color = wdColorBlack
        .Bold = False
        .Size = 12
        .Name = "黑体"
    End With
    With ActiveDocument.Styles(wdStyleHeading6).Font
        .Color = wdColorBlack
        .Bold = False
        .Size = 12
        .Name = "黑体"
    End With
    With ActiveDocument.Styles(wdStyleHeading1).ParagraphFormat
        .LineSpacingRule = wdLineSpace1pt5
    End With
End Sub

```


## VBA 设置标题退为次级标题

```vb
Sub SetSubHeadingStyles()
    Dim para As Paragraph
    Dim level As Long
    
    ' 遍历文档中的每个段落
    For Each para In ActiveDocument.Paragraphs
        ' 获取段落的大纲级别
        level = para.OutlineLevel
        
        ' 如果是标题样式（从标题1到标题9），则将大纲级别减1
        If level >= wdOutlineLevel1 And level <= wdOutlineLevel9 Then
            para.OutlineLevel = level - 1
        End If
    Next para
End Sub

```


## VBA 设置每一个段落大纲级别

```vb
Set myParas = ActiveDocument.Paragraphs
ActiveDocument.Paragraphs.Style = wdStyleNormal
For x = 1 To myParas.Count
    If x Mod 3 = 1 Then
        myParas(x).OutlineLevel = wdOutlineLevel1
    ElseIf x Mod 3 = 2 Then
        myParas(x).OutlineLevel = wdOutlineLevel2
    Else
        myParas(x).OutlineLevel = wdOutlineLevel3
    End If
Next x
ActiveDocument.ActiveWindow.View.Type = wdOutlineView

```

## vba在word中将标题1改为标题2

```vb
Sub ChangeHeading1ToHeading2()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Range.Style = ActiveDocument.Styles(wdStyleHeading1) Then
            para.Range.Style = ActiveDocument.Styles(wdStyleHeading2)
            ElseIf para.Range.Style = ActiveDocument.Styles(wdStyleHeading2) Then
                para.Range.Style = ActiveDocument.Styles(wdStyleHeading3)
            ElseIf para.Range.Style = ActiveDocument.Styles(wdStyleHeading3) Then
                para.Range.Style = ActiveDocument.Styles(wdStyleHeading4)
            ElseIf para.Range.Style = ActiveDocument.Styles(wdStyleHeading4) Then
                para.Range.Style = ActiveDocument.Styles(wdStyleHeading5)
            ElseIf para.Range.Style = ActiveDocument.Styles(wdStyleHeading5) Then
                para.Range.Style = ActiveDocument.Styles(wdStyleHeading6)
            ElseIf para.Range.Style = ActiveDocument.Styles(wdStyleHeading6) Then
                para.Range.Style = ActiveDocument.Styles(wdStyleHeading7)
            ElseIf para.Range.Style = ActiveDocument.Styles(wdStyleHeading7) Then
                para.Range.Style = ActiveDocument.Styles(wdStyleHeading8)
            ElseIf para.Range.Style = ActiveDocument.Styles(wdStyleHeading8) Then
                para.Range.Style = ActiveDocument.Styles(wdStyleHeading9)
        End If
    Next para
End Sub
```


## 删除题1开头的1或者1.1或者1.1.1数字
只能删除自动编号的数字
```vb
Sub RemoveNumberingFromHeading1()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Range.Style = ActiveDocument.Styles(wdStyleHeading1) Then
            para.Range.ListFormat.RemoveNumbers
        End If
    Next para
End Sub

```


### 删除以1或者1.1或者1.1.1

```vb
Sub RemoveNumberingFromHeading1()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Range.Style = ActiveDocument.Styles(wdStyleHeading1) Then
            If Left(para.Range.Text, 1) Like "[0-9]" Then
                para.Range.Text = Mid(para.Range.Text, InStr(para.Range.Text, " ") + 1)
            End If
        End If
    Next para
End Sub

```


```vb
Sub AddCaptionToAllImages()
    Dim oShp As InlineShape
    Dim i As Integer

    ' 遍历文档中的所有图片
    For Each oShp In ActiveDocument.InlineShapes
        oShp.Select
        Selection.InsertCaption Label:="图", Title:=" ", Position:=wdCaptionPositionBelow, ExcludeLabel:=0
    Next oShp
End Sub

```

# 表格部分
## 所有表格添加题主
```vb
Sub AddCaptionToAlltables()
    Dim oShp As Table
    Dim i As Integer

    ' 遍历文档中的所有表格
    For Each oShp In ActiveDocument.Tables
        oShp.Select
        Selection.InsertCaption Label:="表", Title:=" ", Position:=wdCaptionPositionAbove, ExcludeLabel:=0
    Next oShp
End Sub

```

# 域代码部分

## 替换域代码
```vb
Sub ChangeFieldCode()
  Dim oField As Field
  For Each oField In ActiveDocument.Fields
  ' MsgBox (oField.Code.Text)
    If oField.Code.Text = " STYLEREF 1 \s " Then
    ' MsgBox "yes"
      oField.Code.Text = "SEQ 图 \* ARABIC"
    End If
  Next
End Sub
```

## 删除相应域代码
```vb
Sub RemoveField()
  Dim oField As Field
  For Each oField In ActiveDocument.Fields
  ' MsgBox (oField.Code.Text)
    If oField.Code.Text = " STYLEREF 1 \s " Then
    ' MsgBox "yes"
      oField.Delete
    End If
  Next
End Sub
```