' Range(영역)를 For 반복처리
For Each c In Worksheets("Sheet1").Range("B2:BS2")
    If c.Value <> "" Then
        Range("B3").Offset(idx, 0).Value = c.Value
        Range("B3").Offset(idx, 1).Value = c.Address
        idx = idx + 1
    End If
Next

1. 영역 지정을 통한 방법	
Set Ws = Worksheets("레인지")	
Set Sel = Ws.Range("B2:D10")	
Set Sel = Ws.Range("B2").CurrentRegion	
Set Sel = Ws.Range("표1")	
	
Sel.Count	
Sel.Rows.Count	
Sel.Columns.Count	
	
Sel(i, j).Value	
	
2. 영역을 유동적으로 처리하는 방법	
Set Sel = Ws.Range(Cur.End(xlDown).Address, Cur.End(xlDown).End(xlDown).Address)	
	
For Each C In Sel.Cells	
        Debug.Print (C.Offset(, 1).Value)	
Next	
