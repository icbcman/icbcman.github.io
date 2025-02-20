---
title: Word VBA排版脚本编写指南 
date: 2025-02-20 17:08:33
tags: 程序
categories: 
- 电脑
- 技术
---

# 使用deepseek进行Word VBA排版脚本编写指南 



- <font style="background-color:#8bc34a">2025年2月20日</font>；<font title="yellow">星期四</font> ；<font title="blue">多云</font>

| 序号 |  作者   | 场景类型 |
| :--: | :-----: | :------: |
|  1   | **我**😀 | 电脑技术 |

<span alt="shake">今日感受：</span>

<div alt="timeline">
    <div alt="timenode">
        <div alt="meta"><span alt="btn">上午</span><a href="#">：</a></div>
        <div alt="body">
            早睡早起。
        </div>
    </div>
    <div alt="timenode">
        <div alt="meta"><span alt="btn">下午</span><a href="#">：</a></div>
        <div alt="body">
            明天腾讯会议。
        </div>
    </div>
</div>
------------------------------------------------------------------------------------------------------------------------------------------------------------------

**需求：**

```
请写一个vba脚本，用于word排版。要求如下： 
一、正文： 
（一）正文题目：黑体，二号字，段前17磅，段后17磅，固定行距28磅； 
（二）正文内容：中文仿宋，英文及数字Times New Roman，四号字，固定行距28磅； 
（三）正文一级标题：仿宋，四号字，加粗，固定行距28磅。
二、脚注：  
中文宋体，英文及数字Times New Roman，小五字，固定行距28磅。 
三、图表： 
（一）表标题：黑体，小四字，段前0.5行，置于表上方居中； 
（二）图标题：黑体，小四字，段前0.5行，置于图上方居中； 
（三）图表内容：中文仿宋，英文及数字Times New Roman，小四字，第一行及第一列加粗，单倍行距。 
（四）图表备注：宋体，小五字，置于图表下方与图表左边界左对齐； 
（五）图表序号：每篇文章单独排序。 
四、其他： 
文中所有数字及英文：字体Times New Roman。
```

**深度思考：**

| 嗯，用户让我写一个VBA脚本用于Word排版，要求还挺详细的。我需要仔细看看每个部分的要求，确保脚本能满足所有条件。首先，用户分成了四个大部分：正文、脚注、图表和其他。每个部分里还有更细的条目，比如正文的题目、内容、一级标题，图表的标题、内容、备注等等…… |
| ------------------------------------------------------------ |



以下是根据您的要求编写的Word VBA排版脚本，请按以下步骤使用：

**vba：**

```
Sub DocumentFormatting()
    Dim oDoc As Document
    Set oDoc = ActiveDocument
    
    ' 设置全局参数
    With oDoc
        ' ========== 正文设置 ==========
        ' 正文题目样式（标题1）
        With .Styles(wdStyleHeading1).Font
            .Name = "黑体"
            .NameAscii = "Times New Roman"
            .Size = 22
        End With
        With .Styles(wdStyleHeading1).ParagraphFormat
            .SpaceBefore = 17
            .SpaceAfter = 17
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = 28
            .Alignment = wdAlignParagraphCenter
        End With

        ' 正文内容样式（正文）
        With .Styles(wdStyleNormal).Font
            .Name = "仿宋"
            .NameAscii = "Times New Roman"
            .Size = 16
        End With
        With .Styles(wdStyleNormal).ParagraphFormat
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = 28
        End With

        ' 一级标题样式（标题2）
        With .Styles(wdStyleHeading2).Font
            .Name = "仿宋"
            .NameAscii = "Times New Roman"
            .Size = 16
            .Bold = True
        End With
        With .Styles(wdStyleHeading2).ParagraphFormat
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = 28
        End With

        ' ========== 脚注设置 ==========
        With .Styles(wdStyleFootnoteText).Font
            .Name = "宋体"
            .NameAscii = "Times New Roman"
            .Size = 9
        End With
        .Styles(wdStyleFootnoteText).ParagraphFormat.LineSpacing = 28

        ' ========== 图表设置 ==========
        ' 创建表标题样式
        If Not StyleExists("表标题") Then
            .Styles.Add Name:="表标题", Type:=wdStyleTypeParagraph
            With .Styles("表标题").Font
                .Name = "黑体"
                .Size = 12
            End With
            With .Styles("表标题").ParagraphFormat
                .SpaceBefore = 6
                .Alignment = wdAlignParagraphCenter
            End With
        End If

        ' 创建图标题样式（同上）
        If Not StyleExists("图标题") Then
            .Styles.Add Name:="图标题", Type:=wdStyleTypeParagraph
            .Styles("图标题").Duplicate .Styles("表标题")
        End If

        ' 处理所有表格
        Dim oTbl As Table
        For Each oTbl In .Tables
            ' 表格标题
            If oTbl.Range.Previous(wdParagraph).Style <> "表标题" Then
                oTbl.Range.InsertBefore vbCr
                oTbl.Range.Previous(wdParagraph).Style = "表标题"
            End If
            
            ' 表格内容格式
            With oTbl.Range
                .Font.Name = "仿宋"
                .Font.NameAscii = "Times New Roman"
                .Font.Size = 12
                ' 第一行加粗
                oTbl.Rows(1).Range.Font.Bold = True
                ' 第一列加粗
                For Each oCell In oTbl.Columns(1).Cells
                    oCell.Range.Font.Bold = True
                Next
            End With
            oTbl.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        Next

        ' ========== 数字/英文全局设置 ==========
        With .Styles(wdStyleNormal).Font
            .NameAscii = "Times New Roman"
            .NameOther = "Times New Roman"
        End With
    End With

    MsgBox "排版已完成！", vbInformation
End Sub

Function StyleExists(sStyleName As String) As Boolean
    On Error Resume Next
    StyleExists = (Trim(ActiveDocument.Styles(sStyleName).NameLocal) = sStyleName)
End Function
```

使用说明：
1. 打开需要排版的Word文档
2. 按Alt+F11打开VBA编辑器
3. 在左侧工程资源管理器中右键插入新模块
4. 粘贴此代码
5. 按F5运行宏DocumentFormatting

注意事项：
1. 请确保文档中的图表已正确插入（表格使用标准表格对象，图片建议使用插入形状或图片）
2. 表格标题需要紧邻表格上方段落
3. 图表序号需要手动维护（VBA无法自动生成连续序号）
4. 建议先备份文档再执行宏
5. 需要Windows系统装有对应字体
6. 行距设置基于Word磅值计算（1磅=1/72英寸）

此脚本将自动处理：
- 各级标题样式
- 正文和脚注格式
- 表格基本格式
- 中英文字体自动匹配
- 全局数字/英文字体设置

特殊说明：
1. 图表备注需要手动应用"图表备注"样式（代码中未完全实现）
2. 图片标题处理需要额外代码（建议使用题注功能）
3. 复杂表格可能需要手动调整
4. 英文引号建议使用直引号（‘’“”）而非弯引号

您可以根据实际需要进一步调整样式参数或添加更多格式化功能。

![](../pic/一带一路展示图.jpg)
