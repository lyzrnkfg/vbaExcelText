
# open source 开源

```ruby

Sub button1_Click()
   
   Set take = CreateObject("Scripting.Dictionary")
   
   For j = 1 To Sheet2.Range("A65536").End(xlUp).Row
   
        Dim key_test2 As String
        
        key_test2 = CommandButton1_Click(Sheet2.Cells(j, "D").Value, Sheet2.Cells(j, "P").Value, consumpation_check(Sheet2.Cells(j, "G").Value), Sheet2.Cells(j, "J").Value)
        
        take(key_test2) = j
   
   Next j
   
   Sheet4.Cells.Clear
   
   For i = 1 To Sheet1.Range("A65536").End(xlUp).Row
   
    Dim key_test1 As String
        
    key_test1 = CommandButton1_Click(Sheet1.Cells(i, "A").Value, Sheet1.Cells(i, "H").Value, Sheet1.Cells(i, "C").Value, Sheet1.Cells(i, "D").Value)
    
    If take(key_test1) <> "" Then
    
        Sheet2.Rows(take(key_test1)).Copy Sheet4.Cells(i, 1)
        
    Else
    
        Sheet4.Cells(i, 1) = key_test1
        
        Sheet4.Range("A" + CStr(i), "P" + CStr(i)).Interior.Color = RGB(255, 255, 0)
    
    End If
        
   Next i
   
   'For y = 0 To UBound(take.keys)
    'MsgBox take.keys()(y)
   'Next y
   
   Sheet4.Select
    
End Sub

Private Function CommandButton1_Click(name As String, money As String, consumpation As String, summary As String)
    
   Dim result As String
    
   '小写字母-大写字母
   summary = VBA.UCase(summary)
   'VBA将字符由半角-全角
   summary = StrConv(summary, vbNarrow)
   '生成key
   result = name + money + consumpation + summary
   '去半角空格
   result = Replace(result, " ", "")
   '去全角空格
   result = Replace(result, "　", "")
   
   CommandButton1_Click = result
   
End Function

Private Function consumpation_check(consumpation As String)
    
   Dim arr
    
   arr = Split(consumpation, "-")
   
   If ArrayLength(arr) = 3 Then
   
    consumpation_check = arr(2) + arr(1) + arr(0)
    
   Else
   
    consumpation_check = consumpation
    
   End If
   
End Function

Public Function ArrayLength(ByVal ary) As Integer

    ArrayLength = UBound(ary) - LBound(ary) + 1
    
End Function

```

open excel alt+F8 edit