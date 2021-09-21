' Range(영역)를 For 반복처리
For Each c In Worksheets("Sheet1").Range("B2:BS2")
    If c.Value <> "" Then
        Range("B3").Offset(idx, 0).Value = c.Value
        Range("B3").Offset(idx, 1).Value = c.Address
        idx = idx + 1
    End If
Next
