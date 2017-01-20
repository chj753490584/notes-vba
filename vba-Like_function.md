# VBA notes(Like Function)
  
> This is a small VBA programme I wrote when I had internship in Morgan Stanley. 
> The purpose is to check every Fund Conpany and their new-issued funds, which need classifying into different category based on their types.
> I simplify classcification by checking their names. 
> If a name contains 'A' then put it into type 'A', which is realized by Like-Function.
> BTW, it needs data from WIND software.
> More detailed explanation would come later.

## Some simple notes
+ `Like "*" & m(p) & "*"` is better than `Like m(p)`, `"&m(p)&"` can take the value of m(p) and `"*  *"` is fuzzy matching.
+ in VBA, we should consider sheet.select and ActiveSheet, it will impact your loop.

## Programme
```
Sub 申报统计()

'使用方法：请配合wind行政审批进度使用
'如需统计申报数量的情况，请将所有目标时间里的所有申报数据列好，同时将sheet2重命名为统计
'如需统计获批数量，请将所有目标时间里的所有申报数据列好，同时将sheet2重命名为统计
'同时，请先将所有数据按数据透视表，按基金管理人为标签和计数项，降序排出前20名，复制粘贴到“统计”sheet中的（2，A）开始的列中，并在最后一列加入“摩根士丹利华鑫基金管理有限公司”


'定义变量
Sheets("Wind资讯").Select
Range("A1").Select

Dim i, j, p, q, c_stock, c_mix, c_debt, c_money, c_qd, c_total As Integer

'变量赋初始值
  c_stock = 0
  c_mix = 0
  c_debt = 0
  c_money = 0
  c_qd = 0
  c_total = 0
  
i = 1

Dim m(1 To 21) As String

Sheets("统计").Select
Range("A1").Select

For p = 1 To 21
  m(p) = Cells(p + 1, 1)
Next p

p = 1


'模糊匹配并计数
Sheets("Wind资讯").Select
Range("A1").Select

For p = 1 To 21

Sheets("Wind资讯").Select
Range("A1").Select

  For i = 1 To 3500
  
    If Cells(i, 2).Value Like "*股票*" And Cells(i, 1).Value Like m(p) Then
      c_stock = c_stock + 1
    ElseIf Cells(i, 2).Value Like "*混合*" And Cells(i, 1).Value Like m(p) Then
      c_mix = c_mix + 1
    ElseIf Cells(i, 2).Value Like "*债券*" And Cells(i, 1).Value Like m(p) Then
      c_debt = c_debt + 1
    ElseIf Cells(i, 2).Value Like "*货币*" And Cells(i, 1).Value Like m(p) Then
      c_money = c_money + 1
    ElseIf Cells(i, 2).Value Like "*QD*" And Cells(i, 1).Value Like m(p) Then
      c_qd = c_qd + 1
    End If
    
    If Cells(i, 1).Value Like m(p) Then
      c_total = c_total + 1
    End If
    
  Next i

  Sheets("统计").Select
  Range("A1").Select

  
  Cells(p + 1, 2).Value = c_stock

  Cells(p + 1, 3).Value = c_mix

  Cells(p + 1, 4).Value = c_debt

  Cells(p + 1, 5).Value = c_money

  Cells(p + 1, 6).Value = c_qd
  
  Cells(p + 1, 8).Value = c_total
  
  c_stock = 0
  c_mix = 0
  c_debt = 0
  c_money = 0
  c_qd = 0
  c_total = 0

Next p



'结果输出
Sheets("统计").Select
Range("A1").Select

Cells(1, 2).Value = "股票型"
Cells(1, 3).Value = "混合型"
Cells(1, 4).Value = "债券型"
Cells(1, 5).Value = "货币型"
Cells(1, 6).Value = "QD"
Cells(1, 8).Value = "总数"

End Sub
```
