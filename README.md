# Zoom-Attendance-Record
> Zoom視訊軟體有為使用者提供一份與會者的出席紀錄報告，在報告中可以看到每位使用者進出會議的時間。

## 痛點
報告中有詳細的資料，但並沒辦法一眼看出每位出席者的出席狀況，例如何時加入、合適退出、何時有再加入，若能以圖像化應該更好！

## 解決思路
想到以圖像化方式顯示，可能可以參考甘特圖，但只用excel內建的圖表模擬出甘特圖的樣式，卻不能完全符合使用需求，因為有可能會有使用者在會議中中途退出又加入，最後就想到直接用儲存格的背景顏色來呈現。  

解決過程中遇到另一個問題是，若使用者在參與會議的過程中有改名，那會沒辦法透過程式直接知道這兩個名字指的是同一個人，因此需要對與會者的名字做一些的處理。  
以下是針對名字處理的程式碼  
```
// VBA
Set wso = Worksheets("original")
Set wsn = Worksheets("output")

wso.Activate
Columns("A:A").Select
Selection.TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="(", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
Columns("J:L").Select
Selection.EntireColumn.Hidden = True
    
Range("L1").Select
ActiveCell.FormulaR1C1 = "=SUBSTITUTE(RC[-1], "")"", """")"
Selection.AutoFill Destination:=Range("L1:L139")
Range("L1:L139").Select
Range("H1").Select
ActiveCell.FormulaR1C1 = "fixed name"
Range("H2").Select
ActiveCell.FormulaR1C1 = "=XLOOKUP(RC[2], C[4], C[2], RC[2])"
Selection.AutoFill Destination:=Range("H2:H" & Application.CountA(Range("A:A")))
```

## 成果
![img](https://github.com/JT-427/Zoom-Attendance-Record/blob/master/zoom_demo.gif)