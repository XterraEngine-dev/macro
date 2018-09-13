
Sub CONSUMOINDIVIDUO()

'
' CONSUMOINDIVIDUOREP Macro
'

'
    LIMPIAR
    Range("A5").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A4").Select
    ActiveSheet.Paste
    Range("A5").Select
    Application.CutCopyMode = False
    Range("A6:A10").Select
    Selection.ClearContents
    Range("A12:A15").Select
    Selection.ClearContents
    Range("A17:A32").Select
    Selection.ClearContents
    Range("A34:A43").Select
    Selection.ClearContents
    Range("A45:A54").Select
    Selection.ClearContents
    Range("A46").Select
    Selection.End(xlDown).Select
    Range("A56:A64").Select
    Selection.ClearContents
    Range("A57").Select
    Selection.End(xlDown).Select
    Range("A66:A77").Select
    Selection.ClearContents
    Range("A65").Select
    Selection.AutoFill Destination:=Range("A65:A77")
    Range("A65:A77").Select
    ActiveWindow.SmallScroll Down:=-12
    Range("A55").Select
    Selection.AutoFill Destination:=Range("A55:A64")
    Range("A55:A64").Select
    ActiveWindow.SmallScroll Down:=-18
    Range("A44").Select
    Selection.AutoFill Destination:=Range("A44:A54")
    Range("A44:A54").Select
    ActiveWindow.SmallScroll Down:=-12
    Range("A33").Select
    Selection.AutoFill Destination:=Range("A33:A43")
    Range("A33:A43").Select
    ActiveWindow.SmallScroll Down:=-15
    Range("A16").Select
    Selection.AutoFill Destination:=Range("A16:A32")
    Range("A16:A32").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("A11").Select
    Selection.AutoFill Destination:=Range("A11:A15")
    Range("A11:A15").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("A5").Select
    Selection.AutoFill Destination:=Range("A5:A10")
    Range("A5:A10").Select
    Rows("5:5").Select
    Selection.Delete Shift:=xlUp
    Rows("10:10").Select
    Selection.Delete Shift:=xlUp
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Rows("13:13").Select
    Selection.Delete Shift:=xlUp
    Range("B14").Select
    ActiveWindow.SmallScroll Down:=9
    Rows("29:29").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=15
    Rows("39:39").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=15
    Rows("49:49").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=6
    Rows("58:58").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=6
    Rows("62:62").Select
    Selection.Delete Shift:=xlUp
    Rows("65:65").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=-87
    Range("C2:F2").Select
    Selection.Cut
    Range("C3").Select
    ActiveSheet.Paste
    Range("B1").Select
    Selection.Cut
    Range("A3").Select
    ActiveSheet.Paste
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("C3").Select
    Selection.EntireColumn.Insert
    Range("H3").Select
    Selection.EntireColumn.Insert
    Range("M3").Select
    Selection.EntireColumn.Insert
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("C4").Select
    Selection.AutoFill Destination:=Range("C4:C67")
    Range("C4:C67").Select
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("H4:H67")
    Range("H4:H67").Select
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("M4:M67")
    Range("M4:M67").Select
    Range("R3").Select
    Selection.EntireColumn.Insert
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("R4:R67")
    Range("R4:R67").Select
    Range("W4").Select
    Selection.EntireColumn.Insert
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("W4:W67")
    Range("W4:W67").Select
    Range("AB4").Select
    Selection.EntireColumn.Insert
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("AB4:AB67")
    Range("AB4:AB67").Select
    Range("AG4").Select
    Selection.EntireColumn.Insert
    Range("AH2").Select
    Selection.Cut
    Selection.End(xlDown).Select
    Range("AH67").Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AG4").Select
    Selection.AutoFill Destination:=Range("AG4:AG67")
    Range("AG4:AG67").Select
    Range("AL4").Select
    Selection.EntireColumn.Insert
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("AL4:AL67")
    Range("AL4:AL67").Select
    Range("AQ2").Select
    Selection.Cut
    Range("AQ2").Select
    Application.CutCopyMode = False
    Range("AQ4").Select
    Selection.EntireColumn.Insert
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Range("AQ4").Select
    Selection.AutoFill Destination:=Range("AQ4:AQ67")
    Range("AQ4:AQ67").Select
    Range("AV4").Select
    Selection.EntireColumn.Insert
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("AV4:AV67")
    Range("AV4:AV67").Select
    Range("BA4").Select
    Selection.EntireColumn.Insert
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("BA4:BA67")
    Range("BA4:BA67").Select
    Range("BF4").Select
    Selection.EntireColumn.Insert
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("BF4:BF67")
    Range("BF4:BF67").Select
    Range("BK4").Select
    Selection.EntireColumn.Insert
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Range("BK4").Select
    Selection.AutoFill Destination:=Range("BK4:BK67")
    Range("BK4:BK67").Select
    Range("BP4").Select
    Selection.EntireColumn.Insert
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("BP4:BP67")
    Range("BP4:BP67").Select
    Range("BU4").Select
    Selection.EntireColumn.Insert
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Range("BU4").Select
    Selection.AutoFill Destination:=Range("BU4:BU67")
    Range("BU4:BU67").Select
    Range("BZ2").Select
    Selection.Cut
    Range("BZ4").Select
    Range("BZ5").Select
    Application.CutCopyMode = False
    Range("BZ4").Select
    Selection.EntireColumn.Insert
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("BZ4:BZ67")
    Range("BZ4:BZ67").Select
    Range("CE4").Select
    Selection.EntireColumn.Insert
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Selection.End(xlToRight).Select
    Range("CE4").Select
    Selection.AutoFill Destination:=Range("CE4:CE67")
    Range("CE4:CE67").Select
    Range("CJ4").Select
    Selection.EntireColumn.Insert
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CJ4").Select
    Selection.AutoFill Destination:=Range("CJ4:CJ67")
    Range("CJ4:CJ67").Select
    Range("CO3").Select
    Selection.EntireColumn.Insert
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Selection.End(xlToRight).Select
    Range("CO4").Select
    Selection.AutoFill Destination:=Range("CO4:CO67")
    Range("CO4:CO67").Select
    Range("CT4").Select
    Selection.EntireColumn.Insert
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Range("CT4").Select
    Selection.AutoFill Destination:=Range("CT4:CT67")
    Range("CT4:CT67").Select
    Range("CY4").Select
    Selection.EntireColumn.Insert
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("CY4:CY67")
    Range("CY4:CY67").Select
    Range("CY4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT68").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO68").Select
    ActiveSheet.Paste
    Range("CO67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Cut
    Application.CutCopyMode = False
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ68").Select
    ActiveSheet.Paste
    Range("CJ67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE68").Select
    ActiveSheet.Paste
    Range("CE67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ68").Select
    ActiveSheet.Paste
    Range("BZ67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU68").Select
    ActiveSheet.Paste
    Range("BU67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP68").Select
    ActiveSheet.Paste
    Range("BP67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK68").Select
    ActiveSheet.Paste
    Range("BK67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF68").Select
    ActiveSheet.Paste
    Range("BF67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA68").Select
    ActiveSheet.Paste
    Range("BA67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV68").Select
    ActiveSheet.Paste
    Range("AV67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ68").Select
    ActiveSheet.Paste
    Range("AQ67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Selection.End(xlToLeft).Select
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL68").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG68").Select
    ActiveSheet.Paste
    Range("AG67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB68").Select
    ActiveSheet.Paste
    Range("AB67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("AB4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W68").Select
    ActiveSheet.Paste
    Range("W67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R68").Select
    ActiveSheet.Paste
    Range("R67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M68").Select
    ActiveSheet.Paste
    Range("M67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H68").Select
    ActiveSheet.Paste
    Range("H67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C68").Select
    ActiveSheet.Paste
    Range("H1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Delete Shift:=xlToLeft
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "CONSUMO INDIVIDUO"
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A1345")
    Range("A2:A1345").Select
    ActiveWindow.SmallScroll Down:=54
    Range("B63").Select
    Selection.End(xlUp).Select
    Range("B2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("B2:C3").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=18
    Selection.AutoFill Destination:=Range("B2:C1345")
    Range("B2:C1345").Select
    ActiveWindow.SmallScroll Down:=27
    Range("C131").Select
    Selection.End(xlDown).Select
    Range("D1332").Select
    Selection.End(xlUp).Select
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("D2").Select
End Sub


Sub CONSUMOINDIVIDUOMARCASPARTE1()

'
' CONSUMOINDIVIDUOMARCAS Macro
'
'

    LIMPIAR
    Range("A4").Select
    Selection.EntireColumn.Insert
    Range("B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A4").Select
    ActiveSheet.Paste
    Range("A4").Select
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("A5:A11").Select
    Selection.ClearContents
    Range("A13:A34").Select
    Selection.ClearContents
    Range("C35").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A36:A78").Select
    Range("A78").Activate
    Selection.Cut
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("C79").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A80:A82").Select
    Range("A82").Activate
    Selection.ClearContents
    Range("A84:A103").Select
    Selection.ClearContents
    Range("A85").Select
    Selection.End(xlDown).Select
    Range("A105:A106").Select
    Selection.ClearContents
    Range("A108:A125").Select
    Selection.ClearContents
    Range("A109").Select
    Selection.End(xlDown).Select
    Range("C126").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A127:A128").Select
    Range("A128").Activate
    Selection.ClearContents
    Range("C128").Select
    Selection.End(xlDown).Select
    Range("A130:A158").Select
    Selection.ClearContents
    Range("C130").Select
    Selection.End(xlDown).Select
    Range("A160:A203").Select
    Selection.ClearContents
    Range("C160").Select
    Selection.End(xlDown).Select
    Range("A205:A225").Select
    Selection.ClearContents
    Range("C205").Select
    Selection.End(xlDown).Select
    Range("C226").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A256:A268").Select
    Range("A268").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A227:A268").Select
    Range("A268").Activate
    Selection.ClearContents
    Range("C269").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A300").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A270:A300").Select
    Range("A300").Activate
    Selection.ClearContents
    Range("C301").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A323:A324").Select
    Range("A324").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A302:A324").Select
    Range("A324").Activate
    Selection.ClearContents
    Range("C325").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A326:A349").Select
    Range("A349").Activate
    Selection.ClearContents
    Range("C347").Select
    Selection.End(xlDown).Select
    Range("C350").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A351:A367").Select
    Range("A367").Activate
    Selection.ClearContents
    Range("C367").Select
    Selection.End(xlDown).Select
    Range("C368").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A413:A414").Select
    Range("A414").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A369:A414").Select
    Range("A414").Activate
    Selection.ClearContents
    Range("C415").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A416:A424").Select
    Range("A424").Activate
    Selection.ClearContents
    Range("C425").Select
    Selection.End(xlDown).Select
    Range("C427").Select
    Selection.End(xlDown).Select
    Range("A426:A435").Select
    Range("A435").Activate
    Selection.ClearContents
    Range("C435").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A437:A443").Select
    Range("A443").Activate
    Selection.ClearContents
    Range("C443").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A445:A457").Select
    Range("A457").Activate
    Range("A445:A457").Select
    Range("A457").Activate
    Selection.ClearContents
    Range("C458").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A459:A460").Select
    Range("A460").Activate
    Selection.ClearContents
    Range("C461").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A462:A468").Select
    Range("A468").Activate
    Selection.ClearContents
    Range("C468").Select
    Selection.End(xlDown).Select
    Range("C471").Select
    Selection.End(xlDown).Select
    Range("A470:A475").Select
    Range("A475").Activate
    Selection.ClearContents
    Range("C475").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A505").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A477:A505").Select
    Range("A505").Activate
    Selection.ClearContents
    Range("C506").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A526:A527").Select
    Range("A527").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A507:A527").Select
    Range("A527").Activate
    Selection.ClearContents
    Range("C527").Select
    Selection.End(xlDown).Select
    Range("C529").Select
    Selection.End(xlDown).Select
    Range("A563").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A529:A563").Select
    Range("A563").Activate
    Selection.ClearContents
    Range("C565").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A566").Select
    Selection.ClearContents
    Range("A565").Select
    Selection.ClearContents
    Range("C567").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A605").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A568:A605").Select
    Range("A605").Activate
    Selection.ClearContents
    Range("C606").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A637").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A607:A637").Select
    Range("A637").Activate
    Selection.ClearContents
    Range("C638").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A664").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A639:A664").Select
    Range("A664").Activate
    Selection.ClearContents
    Range("C665").Select
    Selection.End(xlDown).Select
    Range("C665").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A667:A668").Select
    Range("A668").Activate
    Selection.ClearContents
    Range("A666").Select
    Selection.ClearContents
    Range("C669").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A713").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A670:A713").Select
    Range("A713").Activate
    Selection.ClearContents
    Range("A715:A723").Select
    Selection.ClearContents
    Range("C724").Select
    Selection.End(xlDown).Select
    Range("C727").Select
    Selection.End(xlDown).Select
    Range("A726").Select
    Selection.ClearContents
    Range("A725").Select
    Selection.ClearContents
    Range("C727").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A785").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A728:A785").Select
    Range("A785").Activate
    Selection.ClearContents
    Range("A784").Select
    Selection.End(xlUp).Select
    Range("B727").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("A4").Select
    Selection.End(xlDown).Select
    Range("A13").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A463").Select
    Selection.End(xlDown).Select
    Range("A471").Select
    Selection.End(xlDown).Select
    Range("A477").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B567").Select
    Selection.End(xlDown).Select
    Range("A1589").Select
    Selection.End(xlUp).Select
    Range("A787").Select
    Selection.ClearContents
    Range("A788").Select
    Selection.ClearContents
    Range("A789").Select
    Selection.ClearContents
    Range("C790").Select
    Selection.End(xlDown).Select
    Range("A850").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A791:A850").Select
    Range("A850").Activate
    ActiveWindow.SmallScroll Down:=-9
    Range("A849").Select
    Selection.End(xlUp).Select
    Range("C786").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A850").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveWindow.SmallScroll Down:=-9
    Selection.ClearContents
    Range("A849").Select
    Selection.End(xlUp).Select
    Range("C786").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("C851").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A909").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A852:A909").Select
    Range("A909").Activate
    Selection.ClearContents
    Range("C910").Select
    Selection.End(xlDown).Select
    Range("A916").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A911:A916").Select
    Range("A916").Activate
    Range("C910").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A925:A927").Select
    Range("A927").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A911:A927").Select
    Range("A927").Activate
    Selection.ClearContents
    Range("C927").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A969").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A929:A969").Select
    Range("A969").Activate
    Selection.ClearContents
    Range("C970").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1016").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A971:A1016").Select
    Range("A1016").Activate
    Selection.ClearContents
    Range("C1017").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1029").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1018:A1029").Select
    Range("A1029").Activate
    Selection.ClearContents
    Range("C1028").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("C1031").Select
    Selection.End(xlDown).Select
    Range("A1070").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1031:A1070").Select
    Range("A1070").Activate
    Selection.ClearContents
    Range("C1071").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1085").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1072:A1085").Select
    Range("A1085").Activate
    Selection.ClearContents
    Range("C1085").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1106").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1087:A1106").Select
    Range("A1106").Activate
    Selection.ClearContents
    Range("C1106").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1142").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1108:A1142").Select
    Range("A1142").Activate
    Selection.ClearContents
    Range("B1142").Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Range("A1589").Select
    Selection.End(xlUp).Select
    Range("C1107").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("C1143").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1169").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1144:A1169").Select
    Range("A1169").Activate
    Selection.ClearContents
    Range("C1170").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1184").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1171:A1184").Select
    Range("A1184").Activate
    Selection.ClearContents
    Range("C1184").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1217").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1186:A1217").Select
    Range("A1217").Activate
    Selection.ClearContents
    Range("C1218").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1240").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1219:A1240").Select
    Range("A1240").Activate
    Selection.ClearContents
    Range("C1240").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("C1242").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1243").Select
    Selection.ClearContents
    Range("A1242").Select
    Selection.ClearContents
    Range("C1245").Select
    Selection.End(xlDown).Select
    Range("A1254").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1245:A1254").Select
    Range("A1254").Activate
    Selection.ClearContents
    Range("C1254").Select
    Selection.End(xlDown).Select
    Range("C1255").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1300").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1256:A1300").Select
    Range("A1300").Activate
    Selection.ClearContents
    Range("C1301").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1306").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1302:A1306").Select
    Range("A1306").Activate
    Selection.ClearContents
    Range("C1307").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1315").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1308:A1315").Select
    Range("A1315").Activate
    Selection.ClearContents
    Range("C1315").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1317:A1334").Select
    Range("A1334").Activate
    Selection.ClearContents
    Range("C1334").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1342").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1336:A1342").Select
    Range("A1342").Activate
    Selection.ClearContents
    Range("C1342").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1350").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1344:A1350").Select
    Range("A1350").Activate
    Selection.ClearContents
    Range("C1350").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1361").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1352:A1361").Select
    Range("A1361").Activate
    Selection.ClearContents
    Range("C1361").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1375").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1363:A1375").Select
    Range("A1375").Activate
    Selection.ClearContents
    Range("C1375").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1385").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1377:A1385").Select
    Range("A1385").Activate
    Selection.ClearContents
    Range("C1386").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1404").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1387:A1404").Select
    Range("A1404").Activate
    Selection.ClearContents
    Range("C1404").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1423").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1406:A1423").Select
    Range("A1423").Activate
    Selection.ClearContents
    Range("C1423").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1453").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1425:A1453").Select
    Range("A1453").Activate
    Selection.ClearContents
    Range("C1453").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1462").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1455:A1462").Select
    Range("A1462").Activate
    Selection.ClearContents
    Range("C1462").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1476").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1464:A1476").Select
    Range("A1476").Activate
    Selection.ClearContents
    Range("C1476").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1482:A1483").Select
    Range("A1483").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1478:A1483").Select
    Range("A1483").Activate
    Selection.ClearContents
    Range("C1483").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1530").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1485:A1530").Select
    Range("A1530").Activate
    Selection.ClearContents
    Range("C1530").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1556").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1532:A1556").Select
    Range("A1556").Activate
    Selection.ClearContents
    Range("C1557").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1568").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1558:A1568").Select
    Range("A1568").Activate
    Selection.ClearContents
    Range("C1569").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1575").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1570:A1575").Select
    Range("A1575").Activate
    Selection.ClearContents
    Range("C1575").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1589").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1577:A1589").Select
    Range("A1589").Activate
    Selection.ClearContents
    Range("C1589").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("C1508").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A4").Select
    Selection.Copy
    Range("A4:A11").Select
    ActiveSheet.Paste
    Range("A12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A12:A34").Select
    ActiveSheet.Paste
    Range("A13").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A36").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A36:A78").Select
    ActiveSheet.Paste
    Range("A37").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A79:A82").Select
    ActiveSheet.Paste
    Range("A82").Select
    Selection.End(xlDown).Select
    Range("A83").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A83").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A83:A103").Select
    ActiveSheet.Paste
    Range("A85").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A105:A106").Select
    ActiveSheet.Paste
    Range("A106").Select
    Application.CutCopyMode = False
    Range("A107").Select
    Selection.Copy
    Range("A108:A125").Select
    ActiveSheet.Paste
    Range("A109").Select
    Selection.End(xlDown).Select
    Range("A126").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A127:A128").Select
    ActiveSheet.Paste
    Range("A129").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A129:A158").Select
    ActiveSheet.Paste
    Range("A130").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A159").Select
    Selection.End(xlDown).Select
    Range("A203").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A204").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A225").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A226").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A227").Select
    Selection.End(xlDown).Select
    Range("A268").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A269").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A270").Select
    Selection.End(xlDown).Select
    Range("A300").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A301").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A302").Select
    Selection.End(xlDown).Select
    Range("A324").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A325").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A325").Select
    Selection.End(xlDown).Select
    Range("A349").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A350").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A351").Select
    Selection.End(xlDown).Select
    Range("A367").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A368").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Range("A414").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A415").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A416").Select
    Selection.End(xlDown).Select
    Range("A424").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A425").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A426").Select
    Selection.End(xlDown).Select
    Range("A435").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A436").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A437").Select
    Selection.End(xlDown).Select
    Range("A443").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A444").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Range("A457").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A458").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Range("A460").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A461").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A462").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A468").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A469").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A470").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A475").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A476").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A477").Select
    Selection.End(xlDown).Select
    Range("A505").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A506").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A507").Select
    Selection.End(xlDown).Select
    Range("A527").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A528").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A529").Select
    Selection.End(xlDown).Select
    Range("A563").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A564").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A566").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A567").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A568").Select
    Selection.End(xlDown).Select
    Range("A605").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A606").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A607").Select
    Selection.End(xlDown).Select
    Range("A637").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A638").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A639").Select
    Selection.End(xlDown).Select
    Range("A664").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A665").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A668").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A669").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A670").Select
    Selection.End(xlDown).Select
    Range("A713").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A714").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A715").Select
    Selection.End(xlDown).Select
    Range("A723").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A724").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A725").Select
    Selection.End(xlDown).Select
    Range("A726").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A727").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A728").Select
    Selection.End(xlDown).Select
    Range("A785").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A786").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A787").Select
    Selection.End(xlDown).Select
    Range("A850").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A851").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A852").Select
    Selection.End(xlDown).Select
    Range("A909").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A910").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A911").Select
    Selection.End(xlDown).Select
    Range("A927").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A928").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A929").Select
    Selection.End(xlDown).Select
    Range("A969").Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Range("A969").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A970").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A971").Select
    Selection.End(xlDown).Select
    Range("A1016").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1017").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1018").Select
    Selection.End(xlDown).Select
    Range("A1029").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1030").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1031").Select
    Selection.End(xlDown).Select
    Range("A1070").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1071").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1072").Select
    Selection.End(xlDown).Select
    Range("A1085").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1086").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1087").Select
    Selection.End(xlDown).Select
    Range("A1106").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1107").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1108").Select
    Selection.End(xlDown).Select
    Range("A1142").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1143").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1144").Select
    Selection.End(xlDown).Select
    Range("A1169").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1170").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1171").Select
    Selection.End(xlDown).Select
    Range("A1184").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1185").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1186").Select
    Selection.End(xlDown).Select
    Range("A1217").Select
    Selection.End(xlUp).Select
    Range("A1186").Select
    Selection.End(xlDown).Select
    Range("A1217").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1218").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1219").Select
    Selection.End(xlDown).Select
    Range("A1240").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1241").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1242").Select
    Selection.End(xlDown).Select
    Range("A1243").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1244").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1245").Select
    Selection.End(xlDown).Select
    Range("A1254").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1255").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1256").Select
    Selection.End(xlDown).Select
    Range("A1300").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1301").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1306").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1307").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1308").Select
    Selection.End(xlDown).Select
    Range("A1315").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1316").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1317").Select
    Selection.End(xlDown).Select
    Range("A1334").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1335").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1336").Select
    Selection.End(xlDown).Select
    Range("A1342").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1343").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1344").Select
    Selection.End(xlDown).Select
    Range("A1350").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1351").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1352").Select
    Selection.End(xlDown).Select
    Range("A1361").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1362").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1363").Select
    Selection.End(xlDown).Select
    Range("A1375").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1376").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Range("A1385").Select
    Selection.End(xlUp).Select
    Range("A1378").Select
    Selection.End(xlDown).Select
    Range("A1385").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1386").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1404").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1405").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1407").Select
    Selection.End(xlDown).Select
    Range("A1423").Select
    Selection.End(xlUp).Select
    Range("A1406").Select
    Selection.End(xlDown).Select
    Range("A1423").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1424").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1425").Select
    Selection.End(xlDown).Select
    Range("A1453").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1454").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1455").Select
    Selection.End(xlDown).Select
    Range("A1462").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1463").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1464").Select
    Selection.End(xlDown).Select
    Range("A1476").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1477").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1478").Select
    Selection.End(xlDown).Select
    Range("A1483").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1484").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1485").Select
    Selection.End(xlDown).Select
    Range("A1530").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1531").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1532").Select
    Selection.End(xlDown).Select
    Range("A1556").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1557").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1558").Select
    Selection.End(xlDown).Select
    Range("A1568").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1569").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1570").Select
    Selection.End(xlDown).Select
    Range("A1575").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1576").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1577").Select
    Selection.End(xlDown).Select
    Range("A1048575").Select
    Selection.End(xlUp).Select
    Range("A1589").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1588").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A11").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C11").Select
    Selection.End(xlDown).Select
    Range("A33").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C33").Select
    Selection.End(xlDown).Select
    Range("A76").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("A79").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D80").Select
    Selection.End(xlDown).Select
    Range("A99").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("E99").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A101").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D101").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("A119").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D119").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A121").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C123").Select
    Selection.End(xlDown).Select
    Range("A150").Select
    Selection.End(xlToRight).Select
    Range("A150").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D150").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A194").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D194").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A215").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D215").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A257").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C257").Select
    Selection.End(xlDown).Select
    Range("A288").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C289").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A311").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C311").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A335").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C335").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A352").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C352").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A398").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C398").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A407").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C409").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A417").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C418").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A424").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C424").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A437").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C437").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A439").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C438").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A446").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C446").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A452").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C452").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A481").Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Range("XFC481").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C481").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A502").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C502").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A537").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C537").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A539").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C539").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A577").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C577").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A608").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C609").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A634").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C634").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A637").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C637").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A681").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
  

   
End Sub


Sub CONSUMOINDIVIDUOMARCASPARTE2()

'
' CONSUMOINDIVIDUOMARCAS Macro
'

'

    Range("C681").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A693").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C694").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A751").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C751").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A815").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C815").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A873").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C873").Select
    Selection.End(xlDown).Select
    Range("A890").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C890").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A931").Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C931").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A977").Select
    Selection.End(xlToRight).Select
    Range("A977").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("B977").Select
    Selection.End(xlDown).Select
    Range("C1548").Select
    Selection.End(xlUp).Select
    Range("C1534").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A989").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C991").Select
    Selection.End(xlDown).Select
    Range("A1029").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1029").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1043").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1044").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1063").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1063").Select
    Selection.End(xlDown).Select
    Range("A1098").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1098").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
  
    Range("A1124").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1124").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1138").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1138").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1170").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1170").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1192").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1192").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1194").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1194").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1204").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1204").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("C1249").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Range("A1249").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1249").Select
    Selection.End(xlDown).Select
    Range("A1254").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1254").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Range("C1254").Select
    Selection.End(xlDown).Select
    Range("A1262").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1262").Select
    Selection.End(xlDown).Select
    Range("A1280:B1280").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1280").Select
    Selection.End(xlDown).Select
    Range("F1278").Select
    Selection.End(xlUp).Select
    Range("F1274").Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlToRight).Select
    Range("A1").Select
    Selection.End(xlDown).Select
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C1533").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("B1502").Select
    Selection.End(xlUp).Select
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("A1287").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1287").Select
    Selection.End(xlDown).Select
    Range("A1294").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1294").Select
    Selection.End(xlDown).Select
    Range("A1304").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1304").Select
    Selection.End(xlDown).Select
    Range("A1317").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1317").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1326").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1326").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1344").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1344").Select
    Selection.End(xlDown).Select
    Range("A1362").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1362").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Range("A1391").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1391").Select
    Selection.End(xlDown).Select
    Range("A1399").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1399").Select
    Selection.End(xlDown).Select
    Range("A1412").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1412").Select
    Selection.End(xlDown).Select
    Range("A1418").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1418").Select
    Selection.End(xlDown).Select
    Range("A1464").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1464").Select
    Selection.End(xlDown).Select
    Range("A1489").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1489").Select
    Selection.End(xlDown).Select
    Range("A1500").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1500").Select
    Selection.End(xlDown).Select
    Range("A1506").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1506").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("C1516").Select
    Selection.End(xlUp).Select
    ActiveWindow.SmallScroll Down:=-15
    Range("C4").Select
    Selection.EntireColumn.Insert
    Range("H4").Select
    Selection.EntireColumn.Insert
    Range("D2:G2").Select
    Selection.Cut
    Range("D3").Select
    ActiveSheet.Paste
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("M4").Select
    Selection.EntireColumn.Insert
    Range("R4").Select
    Selection.EntireColumn.Insert
    Range("W4").Select
    Selection.EntireColumn.Insert
    Range("AB3").Select
    Selection.EntireColumn.Insert
    Range("AG3").Select
    Selection.EntireColumn.Insert
    Range("AL3").Select
    Selection.EntireColumn.Insert
    Range("AQ3").Select
    Selection.EntireColumn.Insert
    Range("AV3").Select
    Selection.EntireColumn.Insert
    Range("BA3").Select
    Selection.EntireColumn.Insert
    Range("BF3").Select
    Selection.EntireColumn.Insert
    Range("BK3").Select
    Selection.EntireColumn.Insert
    Range("BP3").Select
    Selection.EntireColumn.Insert
    Range("BU3").Select
    Selection.EntireColumn.Insert
    Range("BZ3").Select
    Selection.EntireColumn.Insert
    Range("CE3").Select
    Selection.EntireColumn.Insert
    Range("CJ3").Select
    Selection.EntireColumn.Insert
    Range("CO3").Select
    Selection.EntireColumn.Insert
    Range("CT3").Select
    Selection.EntireColumn.Insert
    Range("CY3").Select
    Selection.EntireColumn.Insert
    Range("DD3").Select
    Selection.End(xlToLeft).Select
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Selection.End(xlToLeft).Select
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Range("AX2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Selection.End(xlToLeft).Select
    Range("AQ2").Select
    Selection.End(xlToRight).Select
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Selection.End(xlToLeft).Select
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Range("AH2").Select
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Range("X2").Select
    Selection.Cut
    Application.CutCopyMode = False
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Range("D3").Select
    Selection.End(xlDown).Select
    Range("C1518").Select
    Selection.End(xlUp).Select
    Selection.Copy
    Range("D4").Select
    Selection.End(xlDown).Select
    Range("C1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("H1518").Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("I4").Select
    Selection.End(xlDown).Select
    Range("H1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=-189
    Range("H1320").Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("N4").Select
    Selection.End(xlDown).Select
    Range("M1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("R1518").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("S4").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToRight).Select
    Range("R1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AC1517").Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AC4").Select
    Selection.End(xlDown).Select
    Range("AB1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AG1518").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AH4").Select
    Selection.End(xlDown).Select
    Range("AG1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AM1517").Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AM4").Select
    Selection.End(xlDown).Select
    Range("AL1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AQ1518").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AR4").Select
    Selection.End(xlDown).Select
    Range("AQ1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AV1518").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AW4").Select
    Selection.End(xlDown).Select
    Range("AV1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BB1518").Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BB4").Select
    Selection.End(xlDown).Select
    Range("BA1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BG1518").Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BG4").Select
    Selection.End(xlDown).Select
    Range("BF1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BK1518").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range("BJ1511").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BI1502").Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BL4").Select
    Selection.End(xlDown).Select
    Range("BK1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BP1518").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BQ4").Select
    Selection.End(xlDown).Select
    Range("BP1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BU1518").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BV4").Select
    Selection.End(xlDown).Select
    Range("BU1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BZ1518").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CA4").Select
    Selection.End(xlDown).Select
    Range("BZ1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CE1518").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CF4").Select
    Selection.End(xlDown).Select
    Range("CE1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CK1518").Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CK4").Select
    Selection.End(xlDown).Select
    Range("CJ1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CO1517").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CP5").Select
    Selection.End(xlDown).Select
    Range("CO1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CU1518").Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CU5").Select
    Selection.End(xlDown).Select
    Range("CT1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CZ1517").Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CZ5").Select
    Selection.End(xlDown).Select
    Range("CY1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("DF1518").Select
    Selection.End(xlUp).Select
    Range("DB4").Select
    Application.CutCopyMode = False
    Range("CY4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("CY4:DC4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT1519").Select
    ActiveSheet.Paste
    Range("CT1516").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("CT4:CX4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO1519").Select
    ActiveSheet.Paste
    Range("CO1517").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("CO4:CS4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ1519").Select
    ActiveSheet.Paste
    Range("CJ1517").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("CJ4:CN4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE1519").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=-3
    Range("CE1502").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("CE4:CI4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ1519").Select
    ActiveSheet.Paste
    Range("BZ1517").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("BZ4:CD4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU1519").Select
    ActiveSheet.Paste
    Range("BU1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("BU4:BY4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP1519").Select
    ActiveSheet.Paste
    Range("BP1517").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("BP4:BT4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK1519").Select
    ActiveSheet.Paste
    Range("BK1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("BK4:BO4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF1519").Select
    ActiveSheet.Paste
    Range("BF1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("BF4:BJ4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA1519").Select
    ActiveSheet.Paste
    Range("BA1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("BA4:BE4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV1519").Select
    ActiveSheet.Paste
    Range("AV1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("AV4:AZ4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ1519").Select
    ActiveSheet.Paste
    Range("AQ1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("AQ4:AU4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL1519").Select
    ActiveSheet.Paste
    Range("AL1518").Select
    Selection.End(xlUp).Select
    Range("AL4:AP4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG1519").Select
    ActiveSheet.Paste
    Range("AG1518").Select
    Selection.End(xlUp).Select
    Range("AG4:AK4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB1519").Select
    ActiveSheet.Paste
    Range("AB1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("AB4:AF4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("X4").Select
    Application.CutCopyMode = False
    Range("W4").Select
    Selection.Copy
    Range("X5").Select
    Selection.End(xlDown).Select
    Range("W1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AF1476").Select
    Selection.End(xlUp).Select
    Range("AE1475").Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("AB4:AF4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W1519").Select
    ActiveSheet.Paste
    Range("W1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("W4:AA4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R1519").Select
    ActiveSheet.Paste
    Range("R1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("R4:V4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M1519").Select
    ActiveSheet.Paste
    Range("M1518").Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("M4:Q4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H1519").Select
    ActiveSheet.Paste
    Range("H1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("H4:L4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C1519").Select
    ActiveSheet.Paste
    Range("C1528").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("B1").Select
    Selection.Cut
    Range("A3").Select
    ActiveSheet.Paste
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Pregunta"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("A4").Select
    Selection.EntireColumn.Insert
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "CONSUMO INDIVIDUO MARCAS"
    Range("A4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("A1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("B1518").Select
    Selection.End(xlUp).Select
    Range("B4:C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=21
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("B4:C31818"), Type:=xlFillDefault
    Range("B4:C31818").Select
    ActiveWindow.SmallScroll Down:=-27
    Range("A31804").Select
    Selection.End(xlUp).Select
    Selection.Copy
    Range("B1518").Select
    Selection.End(xlDown).Select
    Range("A31818").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("B31818").Select
    Selection.End(xlUp).Select
    Range("A1:I2").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("I1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("I1:XFD11").Select
    Selection.Delete Shift:=xlToLeft
    Range("H2").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("A2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Find(What:="", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    Selection.Replace What:="", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("D15").Select
    Selection.End(xlDown).Select
	
	
End Sub


Sub CONSUMOHOGAR()


'
' CONSUMOHOGAR Macro
'

    LIMPIAR 
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveWindow.SmallScroll Down:=-126
    Selection.SpecialCells(xlCellTypeBlanks).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A4").Select
    Selection.EntireColumn.Insert
    Range("B4").Select
    Selection.Copy
    Range("A4").Select
    ActiveSheet.Paste
    Range("A5:A24").Select
    ActiveSheet.Paste
    Range("A25").Select
    Application.CutCopyMode = False
    Range("B25").Select
    Selection.Copy
    Range("A25:A30").Select
    ActiveSheet.Paste
    Range("A29").Select
    Application.CutCopyMode = False
    Range("B31").Select
    Selection.Copy
    Range("A31:A39").Select
    ActiveSheet.Paste
    Range("B40").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A40:A45").Select
    ActiveSheet.Paste
    Range("A46").Select
    Application.CutCopyMode = False
    Range("B46").Select
    Selection.Copy
    Range("A46:A56").Select
    ActiveSheet.Paste
    Range("B57").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A57:A73").Select
    ActiveSheet.Paste
    Range("B58").Select
    Selection.End(xlDown).Select
    Range("B101").Select
    Selection.End(xlUp).Select
    Range("B5").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlDown).Select
    Range("B74").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A75").Select
    ActiveSheet.Paste
    Range("A74").Select
    ActiveSheet.Paste
    Range("B76").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A76:A80").Select
    ActiveSheet.Paste
    Range("B81").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B81").Select
    Selection.End(xlDown).Select
    Range("A81:A88").Select
    ActiveSheet.Paste
    Range("B89").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A89:A95").Select
    ActiveSheet.Paste
    Range("B96").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A96:A102").Select
    ActiveSheet.Paste
    Range("D102").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("D4:D102").Select
    Range("D102").Activate
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("D4").Select
    ActiveWindow.SmallScroll Down:=6
    Range("E16").Select
    ActiveWindow.SmallScroll Down:=-21
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A24").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("A29").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("A37").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("A42").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("A52").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Rows("52:52").Select
    Range("XFD52").Activate
    Range("XFC52").Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("A52").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("A68").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("A69").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("A73").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("A80").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("A86").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("A86").Select
    Selection.End(xlUp).Select
    Range("B3").Select
    Selection.End(xlToRight).Select
    Range("G3").Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Range("C3").Select
    Selection.EntireColumn.Insert
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("C4").Select
    Selection.Copy
    Range("D4").Select
    Selection.End(xlDown).Select
    Range("C91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("D89").Select
    Selection.End(xlUp).Select
    Range("D3").Select
    Application.CutCopyMode = False
    Range("D2:G2").Select
    Range("G2").Activate
    Selection.Cut
    Range("D3").Select
    ActiveSheet.Paste
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "target"
    Range("B1").Select
    Selection.Cut
    Range("A3").Select
    ActiveSheet.Paste
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("C3:G3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H3").Select
    Selection.EntireColumn.Insert
    Range("M3").Select
    Selection.EntireColumn.Insert
    Range("R3").Select
    Selection.EntireColumn.Insert
    Range("W3").Select
    Selection.EntireColumn.Insert
    Range("Y3").Select
    Selection.End(xlToRight).Select
    Range("CJ3").Select
    Selection.EntireColumn.Insert
    Range("CF3").Select
    Selection.EntireColumn.Insert
    Range("CB3").Select
    Selection.EntireColumn.Insert
    Range("BX3").Select
    Selection.EntireColumn.Insert
    Range("BT3").Select
    Selection.EntireColumn.Insert
    Range("BP3").Select
    Selection.EntireColumn.Insert
    Range("BL3").Select
    Selection.EntireColumn.Insert
    Range("BH3").Select
    Selection.EntireColumn.Insert
    Range("BD3").Select
    Selection.EntireColumn.Insert
    Range("AZ3").Select
    Selection.EntireColumn.Insert
    Range("AV3").Select
    Selection.EntireColumn.Insert
    Range("AR3").Select
    Selection.EntireColumn.Insert
    Range("AN3").Select
    Selection.EntireColumn.Insert
    Range("AJ3").Select
    Selection.EntireColumn.Insert
    Range("AF3").Select
    Selection.EntireColumn.Insert
    Range("AB3").Select
    Selection.EntireColumn.Insert
    Range("H3").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Range("H4").Select
    Selection.Copy
    Range("I4").Select
    Selection.End(xlDown).Select
    Range("H91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("N91").Select
    Selection.End(xlUp).Select
    Range("N2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Range("M4").Select
    Selection.Copy
    Range("N4").Select
    Selection.End(xlDown).Select
    Range("M91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("S91").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("S4").Select
    Selection.End(xlDown).Select
    Range("R91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("W91").Select
    Selection.End(xlUp).Select
    Range("X2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("X4").Select
    Selection.End(xlDown).Select
    Range("W89").Select
    Selection.End(xlUp).Select
    Range("W4").Select
    ActiveSheet.Paste
    Range("W4").Select
    Selection.Copy
    Range("X4").Select
    Selection.End(xlDown).Select
    Range("W91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AC91").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Range("AB4").Select
    Selection.Copy
    Range("AC4").Select
    Selection.End(xlDown).Select
    Range("AB91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AH91").Select
    Selection.End(xlUp).Select
    Range("AH2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AH4").Select
    Selection.End(xlDown).Select
    Range("AH90").Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Selection.Copy
    Range("AH4").Select
    Selection.End(xlDown).Select
    Range("AG91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AM90").Select
    Selection.End(xlUp).Select
    Range("AM2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Range("AM4").Select
    Selection.End(xlDown).Select
    Range("AL91").Select
    Selection.End(xlUp).Select
    Selection.Copy
    Range("AM4").Select
    Selection.End(xlDown).Select
    Range("AL91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AR91").Select
    Selection.End(xlUp).Select
    Range("AR2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Range("AR4").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Selection.Copy
    Range("AR4").Select
    Selection.End(xlDown).Select
    Range("AQ91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AW91").Select
    Selection.End(xlUp).Select
    Range("AW2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Range("AV4").Select
    Selection.Copy
    Range("AW4").Select
    Selection.End(xlDown).Select
    Range("AV91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BB91").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BB4").Select
    Selection.End(xlDown).Select
    Range("BA91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BG91").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BG4").Select
    Selection.End(xlDown).Select
    Range("BF91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BL91").Select
    Selection.End(xlUp).Select
    Range("BL2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Range("BL4").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Selection.Copy
    Range("BL4").Select
    Selection.End(xlDown).Select
    Range("BK91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BQ91").Select
    Selection.End(xlUp).Select
    Range("BQ2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Range("BP4").Select
    Selection.Copy
    Range("BQ4").Select
    Selection.End(xlDown).Select
    Range("BP91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BW91").Select
    Selection.End(xlUp).Select
    Range("BV2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BV4").Select
    Selection.End(xlDown).Select
    Range("BU91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CA91").Select
    Selection.End(xlUp).Select
    Range("CA2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("CA4").Select
    Selection.End(xlDown).Select
    Range("BZ91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CF91").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CE4").Select
    Selection.Copy
    Range("CF4").Select
    Selection.End(xlDown).Select
    Range("CE91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CJ91").Select
    Selection.End(xlUp).Select
    Range("CK2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("CK4").Select
    Selection.End(xlDown).Select
    Range("CJ91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CP91").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Range("CP4").Select
    Selection.End(xlDown).Select
    Range("CP91").Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Selection.Copy
    Range("CP4").Select
    Selection.End(xlDown).Select
    Range("CO91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CU91").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("CU4").Select
    Selection.End(xlDown).Select
    Range("CT91").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT92").Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("CU4").Select
    Selection.End(xlDown).Select
    Range("CT91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CZ91").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("CZ4").Select
    Selection.End(xlDown).Select
    Range("CY91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("DB90").Select
    Selection.End(xlUp).Select
    Range("CY4:DC4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT92").Select
    ActiveSheet.Paste
    Range("CT91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO92").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("CO92").Select
    ActiveSheet.Paste
    Range("CO91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ92").Select
    ActiveSheet.Paste
    Range("CJ91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE92").Select
    ActiveSheet.Paste
    Range("CE91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ92").Select
    ActiveSheet.Paste
    Range("BZ91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU92").Select
    ActiveSheet.Paste
    Range("BU91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP92").Select
    ActiveSheet.Paste
    Range("BP91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK92").Select
    ActiveSheet.Paste
    Range("BK91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF92").Select
    ActiveSheet.Paste
    Range("BF91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA92").Select
    ActiveSheet.Paste
    Range("BA91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV92").Select
    ActiveSheet.Paste
    Range("AV91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ92").Select
    ActiveSheet.Paste
    Range("AQ91").Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL92").Select
    ActiveSheet.Paste
    Range("AL91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG92").Select
    ActiveSheet.Paste
    Range("AG91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB92").Select
    ActiveSheet.Paste
    Range("AB91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W92").Select
    ActiveSheet.Paste
    Range("W91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R92").Select
    ActiveSheet.Paste
    Range("R91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M92").Select
    ActiveSheet.Paste
    Range("M91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlToRight).Select
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H92").Select
    ActiveSheet.Paste
    Range("H90").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C92").Select
    ActiveSheet.Paste
    Range("C90").Select
    Selection.End(xlUp).Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    Range("H1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlToLeft
    Range("H1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("H1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A105").Select
    ActiveWindow.SmallScroll Down:=-123
    Range("A2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=18
    Selection.AutoFill Destination:=Range("A2:B1849"), Type:=xlFillDefault
    Range("A2:B1849").Select
    Range("B1843").Select
    Selection.End(xlUp).Select
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "194211"
    Range("A1").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "CONSUMO HOGAR"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A1849").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A1849").Select
    Range("A1849").Activate
    ActiveSheet.Paste
    Range("A1848").Select
    Selection.End(xlUp).Select
    Range("A1:H1").Select
    Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
	
    Range("C5").Select
	 Range("I1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("I1:XFD2").Select
    Selection.Delete Shift:=xlUp
    Range("J9").Select
	
	   Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=3
    Selection.Replace What:="", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    ActiveWindow.SmallScroll Down:=-99
    Range("C1737").Select
    Selection.End(xlUp).Select
    Range("D1").Select
    Selection.End(xlDown).Select
    ActiveWindow.SmallScroll Down:=3
	
End Sub



Sub CONSUMOHOGARMARCASPARTE1()

'
' CONSUMOHOGARMARCAS Macro
'

'

    LIMPIAR
    Range("A4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A4").Select
    Application.CutCopyMode = False
    Selection.EntireColumn.Insert
    Range("B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A4").Select
    ActiveSheet.Paste
    Range("A1").Select
    Application.CutCopyMode = False
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C2:F2").Select
    Selection.Cut
    Range("C3").Select
    ActiveSheet.Paste
    Range("C3").Select
    Selection.EntireColumn.Insert
    Range("D1").Select
    Selection.Cut
    Range("C5").Select
    ActiveSheet.Paste
    Range("C5").Select
    Selection.Copy
    Range("C6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("D5").Select
    Selection.End(xlDown).Select
    Range("C23").Select
    Application.CutCopyMode = False
    Range("A5:A21").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=3
    Range("D5").Select
    Selection.End(xlDown).Select
    Range("D23").Select
    Selection.End(xlDown).Select
    Range("A23:A32").Select
    Range("A32").Activate
    Selection.Cut
    Range("B32").Select
    Application.CutCopyMode = False
    Range("A29").Select
    Application.CutCopyMode = False
    Range("A23:A32").Select
    Selection.ClearContents
    Range("D22").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("D34").Select
    Selection.End(xlDown).Select
    Range("A34:A49").Select
    Range("A49").Activate
    Selection.ClearContents
    Range("D49").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A51:A61").Select
    Range("A61").Activate
    Selection.ClearContents
    Range("D63").Select
    Selection.End(xlDown).Select
    Range("A63:A77").Select
    Range("A77").Activate
    Selection.ClearContents
    Range("D79").Select
    Selection.End(xlDown).Select
    Range("A79:A93").Select
    Range("A93").Activate
    Selection.ClearContents
    Range("D95").Select
    Selection.End(xlDown).Select
    Range("A95:A111").Select
    Range("A111").Activate
    Selection.ClearContents
    Range("D111").Select
    Selection.End(xlDown).Select
    Range("D114").Select
    Selection.End(xlDown).Select
    Range("A113:A122").Select
    Range("A122").Activate
    Selection.ClearContents
    Range("D122").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A124:A135").Select
    Range("A135").Activate
    Selection.ClearContents
    Range("D135").Select
    Selection.End(xlDown).Select
    Range("D137").Select
    Selection.End(xlDown).Select
    Range("A137:A142").Select
    Range("A142").Activate
    Selection.ClearContents
    Range("C142").Select
    Selection.End(xlToRight).Select
    Range("D143").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A144:A154").Select
    Range("A154").Activate
    Selection.ClearContents
    Range("D156").Select
    Selection.End(xlDown).Select
    Range("A156:A167").Select
    Range("A167").Activate
    Selection.ClearContents
    Range("D168").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A169:A185").Select
    Range("A185").Activate
    Selection.ClearContents
    Range("D186").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A187:A193").Select
    Range("A193").Activate
    Selection.ClearContents
    Range("D193").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A195:A209").Select
    Range("A209").Activate
    Selection.ClearContents
    Range("D209").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A211:A223").Select
    Range("A223").Activate
    Selection.ClearContents
    Range("D223").Select
    Selection.End(xlDown).Select
    Range("D226").Select
    Selection.End(xlDown).Select
    Range("A225:A245").Select
    Range("A245").Activate
    Selection.ClearContents
    Range("D245").Select
    Selection.End(xlDown).Select
    Range("D247").Select
    Selection.End(xlDown).Select
    Range("A247:A264").Select
    Range("A264").Activate
    Selection.Cut
    Range("B264").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("A247:A264").Select
    Range("A264").Activate
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("D265").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A266:A285").Select
    Range("A285").Activate
    Selection.ClearContents
    Range("D285").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A287:A304").Select
    Range("A304").Activate
    Selection.ClearContents
    Range("D304").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A306:A321").Select
    Range("A321").Activate
    Selection.ClearContents
    Range("D321").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A323:A337").Select
    Range("A337").Activate
    Selection.ClearContents
    Range("D337").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A339:A355").Select
    Range("A355").Activate
    Selection.ClearContents
    Range("D355").Select
    Selection.End(xlDown).Select
    Range("D357").Select
    Selection.End(xlDown).Select
    Range("A357:A373").Select
    Range("A373").Activate
    Selection.ClearContents
    Range("D373").Select
    Selection.End(xlDown).Select
    Range("D375").Select
    Selection.End(xlDown).Select
    Range("A375:A386").Select
    Range("A386").Activate
    Selection.ClearContents
    Range("D386").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A388:A412").Select
    Range("A412").Activate
    Selection.ClearContents
    Range("D413").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A414:A424").Select
    Range("A424").Activate
    Selection.ClearContents
    Range("D424").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A426:A431").Select
    Range("A431").Activate
    Selection.ClearContents
    Range("D431").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A433:A434").Select
    Range("A434").Activate
    Selection.ClearContents
    Range("D436").Select
    Selection.End(xlDown).Select
    Range("A446:A447").Select
    Range("A447").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A436:A447").Select
    Range("A447").Activate
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=18
    Range("B449").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("A5").Select
    Selection.End(xlDown).Select
    Range("A23").Select
    Selection.End(xlDown).Select
    Range("B33").Select
    Selection.End(xlDown).Select
    Range("A1554").Select
    Selection.End(xlUp).Select
    Range("A449:A459").Select
    Selection.ClearContents
    Range("A461:A477").Select
    Selection.ClearContents
    Range("C478").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("B5").Select
    Selection.End(xlDown).Select
    Range("A1554").Select
    Selection.End(xlUp).Select
    Range("A479:A489").Select
    Selection.ClearContents
    Range("I481").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A208").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("A1552").Select
    Selection.End(xlUp).Select
    Range("A491:A503").Select
    Selection.ClearContents
    Range("A505:A525").Select
    Selection.ClearContents
    Range("B506").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("A5").Select
    Selection.End(xlDown).Select
    Range("A23").Select
    Selection.End(xlDown).Select
    Range("A35").Select
    Selection.End(xlDown).Select
    Range("A52").Select
    Selection.End(xlDown).Select
    Range("A105").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B224").Select
    Selection.End(xlDown).Select
    Range("A1554").Select
    Selection.End(xlUp).Select
    Range("A527").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A527:A543").Select
    Range("A527:A546").Select
    Selection.ClearContents
    Range("A528").Select
    Selection.End(xlDown).Select
    Range("A548:A576").Select
    Selection.ClearContents
    Range("A549").Select
    Selection.End(xlDown).Select
    Range("A578:A603").Select
    Selection.ClearContents
    Range("A579").Select
    Selection.End(xlDown).Select
    Range("A605:A624").Select
    Selection.ClearContents
    Range("A606").Select
    Selection.End(xlDown).Select
    Range("A626:A641").Select
    Selection.ClearContents
    Range("A627").Select
    Selection.End(xlDown).Select
    Range("A643:A659").Select
    Range("A643:A658").Select
    Selection.ClearContents
    Range("A644").Select
    Selection.End(xlDown).Select
    Range("A660:A662").Select
    Selection.ClearContents
    Range("A664:A676").Select
    Selection.ClearContents
    Range("A678:A689").Select
    Selection.ClearContents
    Range("B678").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("A4").Select
    Selection.End(xlDown).Select
    Range("B22").Select
    Selection.End(xlDown).Select
    Range("A1554").Select
    Selection.End(xlUp).Select
    Range("D690").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A691:A702").Select
    Range("A702").Activate
    Selection.ClearContents
    Range("D703").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A704:A719").Select
    Range("A719").Activate
    Selection.ClearContents
    Range("D719").Select
    Selection.End(xlDown).Select
    Range("D721").Select
    Selection.End(xlDown).Select
    Range("A721:A737").Select
    Range("A737").Activate
    Selection.ClearContents
    Range("D737").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A753:A754").Select
    Range("A754").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A739:A754").Select
    Range("A754").Activate
    Selection.ClearContents
    Range("D754").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A756:A766").Select
    Range("A766").Activate
    Selection.ClearContents
    Range("D766").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A784").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A768:A784").Select
    Range("A784").Activate
    Selection.ClearContents
    Range("D784").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("D822").Select
    Selection.End(xlUp).Select
    Range("D786").Select
    Selection.End(xlDown).Select
    Range("A822").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A786:A822").Select
    Range("A822").Activate
    Selection.ClearContents
    Range("A824:A830").Select
    Selection.ClearContents
    Range("A832:A846").Select
    Range("A846").Activate
    Selection.ClearContents
    Range("A848:A850").Select
    Selection.ClearContents
    Range("A852:A863").Select
    Selection.ClearContents
    Range("A865:A866").Select
    Range("A866").Activate
    Selection.ClearContents
    Range("A868:A874").Select
    Range("A874").Activate
    Selection.ClearContents
    Range("D875").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A876:A880").Select
    Range("A880").Activate
    Selection.ClearContents
    Range("D881").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A906").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A882:A906").Select
    Range("A906").Activate
    Selection.ClearContents
    Range("D908").Select
    Selection.End(xlDown).Select
    Range("A938").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A908:A938").Select
    Range("A938").Activate
    Selection.ClearContents
    Range("A941").Select
    Selection.ClearContents
    Range("A940").Select
    Selection.ClearContents
    Range("D943").Select
    Selection.End(xlDown).Select
    Range("A984").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A943:A984").Select
    Range("A984").Activate
    Selection.ClearContents
    Range("D985").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1010").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A986:A1010").Select
    Range("A1010").Activate
    Selection.ClearContents
    Range("D1012").Select
    Selection.End(xlDown).Select
    Range("A1012:A1014").Select
    Range("A1014").Activate
    Selection.ClearContents
    Range("D1014").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1016:A1024").Select
    Range("A1024").Activate
    Selection.ClearContents
    Range("D1024").Select
    Selection.End(xlDown).Select
    Range("D1026").Select
    Selection.End(xlDown).Select
    Range("A1049").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1026:A1049").Select
    Range("A1049").Activate
    Selection.ClearContents
    Range("D1049").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1068").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1051:A1068").Select
    Range("A1068").Activate
    Selection.ClearContents
    Range("D1068").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1070:A1075").Select
    Range("A1075").Activate
    Selection.ClearContents
    Range("D1075").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1101").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1077:A1101").Select
    Range("A1101").Activate
    Selection.ClearContents
    Range("D1102").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1127").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1103:A1127").Select
    Range("A1127").Activate
    Selection.ClearContents
    Range("D1128").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1130:A1131").Select
    Range("A1131").Activate
    Selection.ClearContents
    Range("A1129:A1131").Select
    Range("A1131").Activate
    Selection.ClearContents
    Range("D1132").Select
    Selection.End(xlDown).Select
    Range("D1133").Select
    Selection.End(xlDown).Select
    Range("A1154").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1133:A1154").Select
    Range("A1154").Activate
    Selection.ClearContents
    Range("D1155").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1186").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1156:A1186").Select
    Range("A1186").Activate
    Selection.ClearContents
    Range("D1187").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1204").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1188:A1204").Select
    Range("A1204").Activate
    Selection.ClearContents
    Range("D1205").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1213").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1206:A1213").Select
    Range("A1213").Activate
    Selection.ClearContents
    Range("D1213").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1232").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1215:A1232").Select
    Range("A1232").Activate
    Selection.ClearContents
    Range("D1232").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1247").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1234:A1247").Select
    Range("A1247").Activate
    Selection.ClearContents
    Selection.End(xlToRight).Select
    Range("D1247").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1254").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1249:A1254").Select
    Range("A1254").Activate
    Selection.ClearContents
    Range("D1254").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1280").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1256:A1280").Select
    Range("A1280").Activate
    Selection.ClearContents
    Range("D1281").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1310").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1282:A1310").Select
    Range("A1310").Activate
    Selection.ClearContents
    Range("D1311").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1325").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1312:A1325").Select
    Range("A1325").Activate
    Selection.ClearContents
    Range("D1325").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1351").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1327:A1351").Select
    Range("A1351").Activate
    Selection.ClearContents
    Range("D1353").Select
    Selection.End(xlDown).Select
    Range("A1365").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1353:A1365").Select
    Range("A1365").Activate
    Selection.ClearContents
    Range("D1365").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1367:A1375").Select
    Range("A1375").Activate
    Selection.ClearContents
    Range("D1375").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1388").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1377:A1388").Select
    Range("A1388").Activate
    Selection.ClearContents
    Range("D1388").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1422").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1390:A1422").Select
    Range("A1422").Activate
    Selection.ClearContents
    Range("D1424").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1433").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1424:A1433").Select
    Range("A1433").Activate
    Selection.ClearContents
    Range("D1434").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1451").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1435:A1451").Select
    Range("A1451").Activate
    Selection.ClearContents
    Range("D1451").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1477").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1453:A1477").Select
    Range("A1477").Activate
    Selection.ClearContents
    Range("D1478").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1486").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1479:A1486").Select
    Range("A1486").Activate
    Selection.ClearContents
    Range("D1486").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1505").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1488:A1505").Select
    Range("A1505").Activate
    Selection.ClearContents
    Range("D1505").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1512").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1507:A1512").Select
    Range("A1512").Activate
    Selection.ClearContents
    Range("D1512").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1531").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1514:A1531").Select
    Range("A1531").Activate
    Selection.ClearContents
    Range("D1531").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1540").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1533:A1540").Select
    Range("A1540").Activate
    Selection.ClearContents
    Range("D1540").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1554").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1542:A1554").Select
    Range("A1554").Activate
    Selection.ClearContents
    Range("D1554").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("C1048576").Select
    Selection.End(xlUp).Select
    Range("B4").Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A4").Select
    Selection.Copy
    Range("A5:A21").Select
    ActiveSheet.Paste
    Range("A6").Select
    Selection.End(xlDown).Select
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A23:A32").Select
    ActiveSheet.Paste
    Range("A24").Select
    Selection.End(xlDown).Select
    Range("A33").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A34").Select
    Selection.End(xlDown).Select
    Range("A33").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A33:A49").Select
    ActiveSheet.Paste
    Range("A37").Select
    Selection.End(xlDown).Select
    Range("A50").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A51:A61").Select
    ActiveSheet.Paste
    Range("A62").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A63:A77").Select
    ActiveSheet.Paste
    Range("A65").Select
    Selection.End(xlDown).Select
    Range("A78").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A79:A93").Select
    ActiveSheet.Paste
    Range("A80").Select
    Selection.End(xlDown).Select
    Range("A94").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A95").Select
    Selection.End(xlDown).Select
    Range("A94").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A94:A111").Select
    ActiveSheet.Paste
    Range("A95").Select
    Selection.End(xlDown).Select
    Range("A112").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A113:A122").Select
    ActiveSheet.Paste
    Range("A114").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A123:A135").Select
    ActiveSheet.Paste
    Range("A125").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A136:A142").Select
    ActiveSheet.Paste
    Range("A137").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A143").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A143:A154").Select
    ActiveSheet.Paste
    Range("A144").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Range("A155").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A155:A167").Select
    ActiveSheet.Paste
    Range("A156").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A168:A185").Select
    ActiveSheet.Paste
    Range("A169").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A186").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A186:A193").Select
    ActiveSheet.Paste
    Range("A187").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Range("A208").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A194:A209").Select
    ActiveSheet.Paste
    Range("A195").Select
    Selection.End(xlDown).Select
    Range("A210").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A211").Select
    Selection.End(xlDown).Select
    Range("A223").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A224").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A225").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A225:A245").Select
    ActiveSheet.Paste
    Range("A226").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A246").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A246:A264").Select
    ActiveSheet.Paste
    Range("A247").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A285").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A286").Select
    Selection.End(xlDown).Select
    Range("A304").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A287").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A287:A304").Select
    ActiveSheet.Paste
    Range("A288").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A305:A321").Select
    ActiveSheet.Paste
    Range("A306").Select
    Selection.End(xlDown).Select
    Range("A323").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A323").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A323:A337").Select
    ActiveSheet.Paste
    Range("A324").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A339").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A339:A355").Select
    ActiveSheet.Paste
    Range("A340").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A356:A373").Select
    ActiveSheet.Paste
    Range("A357").Select
    Selection.End(xlDown).Select
    Range("A374").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A375").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A375:A386").Select
    ActiveSheet.Paste
    Range("A376").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A387:A412").Select
    ActiveSheet.Paste
    Range("A388").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A413:A424").Select
    ActiveSheet.Paste
    Range("A414").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Flanes y postres envasados"
    Range("A425").Select
    Selection.Copy
    Range("A426").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A426:A431").Select
    ActiveSheet.Paste
    Range("A430").Select
    Selection.End(xlDown).Select
    Range("A432").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A433").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A433:A434").Select
    ActiveSheet.Paste
    Range("A435").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A435").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A435:A447").Select
    ActiveSheet.Paste
    Range("A436").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A449:A459").Select
    ActiveSheet.Paste
    Range("A460").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A460:A477").Select
    ActiveSheet.Paste
    Range("A461").Select
    Selection.End(xlDown).Select
    Range("A478").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A479:A489").Select
    ActiveSheet.Paste
    Range("A490").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A491").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A491:A503").Select
    ActiveSheet.Paste
    Range("A492").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A505").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A505:A525").Select
    ActiveSheet.Paste
    Range("A506").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A526:A546").Select
    ActiveSheet.Paste
    Range("A527").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A548").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A548:A576").Select
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A577:A603").Select
    ActiveSheet.Paste
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlDown).Select
    Range("A604").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A604:A624").Select
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A625:A641").Select
    ActiveSheet.Paste
    Range("A624").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A642:A658").Select
    ActiveSheet.Paste
    Range("A643").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A659:A662").Select
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A663:A676").Select
    ActiveSheet.Paste
    Range("A664").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A665").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A677:A689").Select
    ActiveSheet.Paste
    Range("A678").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A690").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A690:A702").Select
    ActiveSheet.Paste
    Range("A691").Select
    Selection.End(xlDown).Select
    Range("A703").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A704").Select
    Selection.End(xlDown).Select
    Range("A719").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A720").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A721").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A721:A737").Select
    ActiveSheet.Paste
    Range("A722").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A739").Select
    Selection.End(xlDown).Select
    Range("A753:A754").Select
    Range("A754").Activate
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A755").Select
    Selection.End(xlDown).Select
    Range("A766").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A756").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A756:A766").Select
    ActiveSheet.Paste
    Range("A757").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A767:A784").Select
    ActiveSheet.Paste
    Range("A768").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A786").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A786:A822").Select
    ActiveSheet.Paste
    Range("A787").Select
    Selection.End(xlDown).Select
    Range("A824").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A823").Select
    Selection.End(xlDown).Select
    Range("A830").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A827").Select
    Application.CutCopyMode = False
    Range("A823").Select
    Selection.Copy
    Range("A823").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A830").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A831").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Range("A846").Select
    Range(Selection, Selection.End(xlUp)).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("A845").Select
    Selection.End(xlUp).Select
    Selection.Copy
    Range("A832").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A832:A846").Select
    ActiveSheet.Paste
    Range("A833").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A848").Select
    Selection.End(xlDown).Select
    Range("A850").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A851").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A852").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A852:A863").Select
    ActiveSheet.Paste
    Range("A853").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A865").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A865").Select
    ActiveSheet.Paste
    Range("A866").Select
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A868").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A868:A874").Select
    ActiveSheet.Paste

       Range("A869").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A876").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A876:A880").Select
    ActiveSheet.Paste
    Range("A877").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A881").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A881:A906").Select
    ActiveSheet.Paste
    Range("A882").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A908:A938").Select
    ActiveSheet.Paste
    Range("A909").Select
    Selection.End(xlDown).Select
    Range("A939").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A940").Select
    ActiveSheet.Paste
    Range("A941").Select
    ActiveSheet.Paste
    Range("A942").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A942").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A942:A984").Select
    ActiveSheet.Paste
    Range("A943").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A985:A1010").Select
    ActiveSheet.Paste
    Range("A986").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1011:A1014").Select
    ActiveSheet.Paste
    Range("A1012").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1015").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Range("A1015:A1024").Select
    ActiveSheet.Paste
    Range("A1016").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1025:A1049").Select
    ActiveSheet.Paste
    Range("A1026").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1050:A1068").Select
    ActiveSheet.Paste
    Range("A1051").Select
    Selection.End(xlDown).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1076:A1101").Select
    Range("A1069").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1069:A1075").Select
    ActiveSheet.Paste
    Range("A1076").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1076:A1101").Select
    ActiveSheet.Paste
    Range("A1077").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1103").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1103:A1127").Select
    ActiveSheet.Paste
    Range("A1104").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1128:A1131").Select
    ActiveSheet.Paste
    Range("A1129").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1132:A1133").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1132:A1154").Select
    ActiveSheet.Paste
    Range("A1133").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1156").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1156:A1186").Select
    ActiveSheet.Paste
    Range("A1157").Select
    Selection.End(xlDown).Select
    Range("A1187").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1188").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1188:A1204").Select
    ActiveSheet.Paste
    Range("A1190").Select
    Selection.End(xlDown).Select
    Range("A1205").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1206").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1206:A1213").Select
    ActiveSheet.Paste
    Range("A1207").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1214:A1232").Select
    ActiveSheet.Paste
    Range("A1215").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1233:A1247").Select
    ActiveSheet.Paste
    Range("A1234").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1249:A1254").Select
    ActiveSheet.Paste
    Range("A1250").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1255:A1280").Select
    ActiveSheet.Paste
    Range("A1256").Select
    Selection.End(xlDown).Select
    Range("A1281").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1282").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1282:A1310").Select
    ActiveSheet.Paste
    Range("A1283").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1312:A1325").Select
    ActiveSheet.Paste
    Range("A1314").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1327").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1327:A1351").Select
    ActiveSheet.Paste
    Range("A1328").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1352").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1352:A1365").Select
    ActiveSheet.Paste
    Range("A1353").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1366:A1375").Select
    ActiveSheet.Paste
    Range("A1367").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Range("A1388").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1389").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1422").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1423").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1424").Select
    Selection.End(xlDown).Select
    Range("A1433").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1451").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1452").Select
    Selection.End(xlDown).Select
    Range("A1477").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1453").Select
    Selection.End(xlDown).Select
    Range("A1477").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1478").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1479").Select
    Selection.End(xlDown).Select
    Range("A1486").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1487").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Range("A1505").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1507").Select
    Selection.End(xlDown).Select
    Range("A1513").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1512").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1507:A1512").Select
    Range("A1512").Activate
    ActiveSheet.Paste
    Range("A1513").Select
    Selection.End(xlDown).Select
    Range("A1531").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1513").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1514:A1515").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A1514:A1531").Select
    ActiveSheet.Paste
    Range("A1515").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1533").Select
    Selection.End(xlDown).Select
    Range("A1540").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1541").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("A1554").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1553").Select
    Application.CutCopyMode = False
    Range("A1555").Select


End Sub



Sub CONSUMOHOGARMARCASPARTE2()

'
' CONSUMOHOGARMARCAS Macro
'

'

    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("A21").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D21").Select
    Selection.End(xlDown).Select
    Range("A31").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D31").Select
    Selection.End(xlDown).Select
    Range("A47").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D47").Select
    Selection.End(xlDown).Select
    Range("A58").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D58").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A74").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("A73").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D73").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A88").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D88").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A105").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D105").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A115").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D115").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A127").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D127").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A133").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D133").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A144").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D144").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A156").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D156").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A173").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D173").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A180").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D180").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A195").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D195").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A208").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D208").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A229").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("A229").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D229").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A247").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D247").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A267").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D267").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A285").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D285").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A301").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D302").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A316").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D316").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A333").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D334").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A350").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("A350").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("E350").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A362").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D362").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A387").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D387").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A398").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D398").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A404").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D404").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("A406").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Rows("406:406").Select
    Range("XFD406").Activate
    Selection.Delete Shift:=xlUp
    Range("XFB406").Select
    Selection.End(xlToLeft).Select
    Range("AC406").Select
    Selection.End(xlToLeft).Select
    Range("D406").Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Range("A418").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D418").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A429").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D429").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A446").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D446").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A457").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D457").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A470").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D470").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A491").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D491").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A511").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D511").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A540").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D540").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A566").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D566").Select
    Selection.End(xlDown).Select
    Range("A586").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D586").Select
    Selection.End(xlDown).Select
    Range("A602").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D602").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A618").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D618").Select
    Selection.End(xlDown).Select
    Range("A621").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D621").Select
    Selection.End(xlDown).Select
    Range("A634").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D634").Select
    Selection.End(xlDown).Select
    Range("A646").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D646").Select
    Selection.End(xlDown).Select
    Range("A658").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D658").Select
    Selection.End(xlDown).Select
    Range("A674").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D674").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A691").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D691").Select
    Selection.End(xlDown).Select
    Range("A707").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D707").Select
    Selection.End(xlDown).Select
    Range("A718").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D718").Select
    Selection.End(xlDown).Select
    Range("A735").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D735").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A772").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D772").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A779").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D779").Select
    Selection.End(xlDown).Select
    Range("A794").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D794").Select
    Selection.End(xlDown).Select
    Range("A797").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D797").Select
    Selection.End(xlDown).Select
    Range("A809").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D809").Select
    Selection.End(xlDown).Select
    Range("D811").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D811").Select
    Selection.End(xlDown).Select
    Range("A818").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D818").Select
    Selection.End(xlDown).Select
    Range("A823").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D823").Select
    Selection.End(xlDown).Select
    Range("A848").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D848").Select
    Selection.End(xlDown).Select
    Range("A879").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D879").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A881").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D881").Select
    Selection.End(xlDown).Select
    Range("A923").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D923").Select
    Selection.End(xlDown).Select
    Range("A948").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D948").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A951").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D951").Select
    Selection.End(xlDown).Select
    Range("A960").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D960").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A984").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D984").Select
    Selection.End(xlDown).Select
    Range("A1002").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1002").Select
    Selection.End(xlDown).Select
    Range("A1008").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1008").Select
    Selection.End(xlDown).Select
    Range("A1033").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1033").Select
    Selection.End(xlDown).Select
    Range("A1058").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1058").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1061").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1061").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1083").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1083").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1114").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1114").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1131").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1131").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1139").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1139").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1158").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1157").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1157").Select
    Selection.End(xlDown).Select
    Range("A1171").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1171").Select
    Selection.End(xlDown).Select
    Range("A1177").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1177").Select
    Selection.End(xlDown).Select
    Range("A1202").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1202").Select
    Selection.End(xlDown).Select
    Range("A1231").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1231").Select
    Selection.End(xlDown).Select
    Range("A1245").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1245").Select
    Selection.End(xlDown).Select
    Range("A1270").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1270").Select
    Selection.End(xlDown).Select
    Range("A1283").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1283").Select
    Selection.End(xlDown).Select
    Range("A1292").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1292").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1304").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1304").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1337").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C1337").Select
    Selection.End(xlToRight).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1347").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("E1347").Select
    Selection.End(xlDown).Select
    Range("D1363").Select
    Selection.End(xlDown).Select
    Range("A1364").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1365").Select
    Selection.End(xlDown).Select
    Range("A1389").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1389").Select
    Selection.End(xlDown).Select
    Range("A1397").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1404").Select
    Selection.End(xlDown).Select
    Range("A1415").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1415").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1421").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1421").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1439").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1439").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1447").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("D1447").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("D1458").Select
    Selection.End(xlUp).Select
    Range("D2").Select
    Selection.End(xlDown).Select
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("B1048575").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("E15").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range("C3").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("H4").Select
    Selection.EntireColumn.Insert
    Range("M3").Select
    Selection.EntireColumn.Insert
    Range("R3").Select
    Selection.EntireColumn.Insert
    Range("W3").Select
    Selection.EntireColumn.Insert
    Range("AB3").Select
    Selection.EntireColumn.Insert
    Range("AG3").Select
    Selection.EntireColumn.Insert
    Range("AL3").Select
    Selection.EntireColumn.Insert
    Range("AQ3").Select
    Selection.EntireColumn.Insert
    Range("AV3").Select
    Selection.EntireColumn.Insert
    Range("BA3").Select
    Selection.EntireColumn.Insert
    Range("BF3").Select
    Selection.EntireColumn.Insert
    Range("BK3").Select
    Selection.EntireColumn.Insert
    Range("BP3").Select
    Selection.EntireColumn.Insert
    Range("BU3").Select
    Selection.EntireColumn.Insert
    Range("BZ3").Select
    Selection.EntireColumn.Insert
    Range("CE3").Select
    Selection.EntireColumn.Insert
    Range("CJ3").Select
    Selection.EntireColumn.Insert
    Range("CO3").Select
    Selection.EntireColumn.Insert
    Range("CT3").Select
    Selection.EntireColumn.Insert
    Range("CY3").Select
    Selection.EntireColumn.Insert
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Range("CK2").Select
    Selection.Cut
    Selection.End(xlDown).Select
    Range("CK1460").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Range("BQ2").Select
    Selection.Cut
    Range("BO4").Select
    Range("BP4").Select
    ActiveSheet.Paste
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Range("BH2").Select
    Selection.Cut
    Range("BG2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Application.CutCopyMode = False
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Range("AV2").Select
    Selection.End(xlToRight).Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Selection.End(xlToLeft).Select
    Range("AM2").Select
    Selection.Cut
    Range("AK4").Select
    Range("AL4").Select
    ActiveSheet.Paste
    Range("AL2").Select
    Selection.End(xlToLeft).Select
    Range("AH2").Select
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Selection.End(xlToLeft).Select
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Selection.End(xlToLeft).Select
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Range("C4").Select
    Selection.Copy
    Range("D5").Select
    Selection.End(xlDown).Select
    Range("C1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("I1459").Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("I4").Select
    Selection.End(xlDown).Select
    Range("H1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("N1459").Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("N5").Select
    Selection.End(xlDown).Select
    Range("M1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("S1459").Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("S4").Select
    Selection.End(xlDown).Select
    Range("R1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("X1459").Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("X4").Select
    Selection.End(xlDown).Select
    Range("W1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AC1458").Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AC4").Select
    Selection.End(xlDown).Select
    Range("AB1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AH1459").Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AH4").Select
    Selection.End(xlDown).Select
    Range("AG1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AM1458").Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AM4").Select
    Selection.End(xlDown).Select
    Range("AL1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AR1458").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AR1456").Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AR4").Select
    Selection.End(xlDown).Select
    Range("AQ1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AW1459").Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AW4").Select
    Selection.End(xlDown).Select
    Range("AV1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BB1459").Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BB4").Select
    Selection.End(xlDown).Select
    Range("BA1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BF1458").Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BG5").Select
    Selection.End(xlDown).Select
    Range("BF1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BK1459").Select
    Selection.End(xlUp).Select
    Range("BK3").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BL4").Select
    Selection.End(xlDown).Select
    Range("BK1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BQ1459").Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BQ4").Select
    Selection.End(xlDown).Select
    Range("BP1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BV1458").Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BV5").Select
    Selection.End(xlDown).Select
    Range("BU1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CA1459").Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CA4").Select
    Selection.End(xlDown).Select
    Range("BZ1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CF1459").Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlToRight).Select
    Range("CY4").Select
    Selection.End(xlToLeft).Select
    Range("C4").Select
    Selection.End(xlToRight).Select
    Range("DB2").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToRight).Select
    Range("XFD3").Select
    Selection.End(xlToLeft).Select
    Range("DB4").Select
    Selection.End(xlToLeft).Select
    Range("BB6").Select
    Selection.End(xlToRight).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CF4").Select
    Selection.End(xlDown).Select
    Range("CE1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CJ1354").Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CK4").Select
    Selection.End(xlDown).Select
    Range("CJ1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CO1459").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CP4").Select
    Selection.End(xlDown).Select
    Range("CO1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CT1459").Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CU4").Select
    Selection.End(xlDown).Select
    Range("CT1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CY1459").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CZ4").Select
    Selection.End(xlDown).Select
    Range("CY1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("DA1459").Select
    Selection.End(xlToRight).Select
    Range("DD1459").Select
    Selection.End(xlUp).Select
    Range("DG1").Select
    Selection.End(xlToLeft).Select
    Range("DG1").Select
    Selection.End(xlToLeft).Select
    Range("H4").Select
    Selection.End(xlToRight).Select
    Range("CW3").Select
    Application.CutCopyMode = False
    Range("CY4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT1460").Select
    ActiveSheet.Paste
    Range("CT1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO1460").Select
    ActiveSheet.Paste
    Range("CO1460").Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("CO4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ1460").Select
    ActiveSheet.Paste
    Range("CJ1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE1460").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveSheet.Paste
    Range("CF1458").Select
    Application.CutCopyMode = False
    Range("CK1457").Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE1460").Select
    ActiveSheet.Paste
    Range("CE1458").Select
   Selection.End(xlUp).Select
    Range("CE4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ1460").Select
    ActiveSheet.Paste
    Range("BZ1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU1460").Select
    ActiveSheet.Paste
    Range("BU1458").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP1460").Select
    ActiveSheet.Paste
    Range("BP1459").Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK1460").Select
    ActiveSheet.Paste
    Range("BK1458").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF1460").Select
    ActiveSheet.Paste
    Range("BF1458").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA1460").Select
    ActiveSheet.Paste
    Range("BA1458").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV1460").Select
    ActiveSheet.Paste
    Range("AV1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ1460").Select
    ActiveSheet.Paste
    Range("AQ1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL1460").Select
    ActiveSheet.Paste
    Range("AL1458").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG1460").Select
    ActiveSheet.Paste
    Range("AG1458").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Selection.End(xlToLeft).Select
    Range("B4").Select
    Selection.End(xlToRight).Select
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB1460").Select
    ActiveSheet.Paste
    Range("AB1457").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W1460").Select
    ActiveSheet.Paste
    Range("W1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R1460").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("R4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M1460").Select
    ActiveSheet.Paste
    Range("M1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H1460").Select
    ActiveSheet.Paste
    Range("H1458").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C1460").Select
    ActiveSheet.Paste
    Range("C1457").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range("B1").Select
    Selection.Cut
    Range("A3").Select
    ActiveSheet.Paste
    Range("A3").Select
    Selection.EntireColumn.Insert
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "CONSUMO HOGAR MARCAS"
    Range("A4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("A1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A4:A1459").Select
    Range("A1459").Activate
    Application.CutCopyMode = False
    Selection.Copy
    Range("D1462").Select
    Application.CutCopyMode = False
    Range("A1459").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("C1459").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("A1458:C1459").Select
    Range("C1459").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A5:C1459").Select
    Range("C1459").Activate
    ActiveWindow.SmallScroll Down:=-18
    Range("A4:C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=15
    Selection.AutoFill Destination:=Range("A4:C30579"), Type:=xlFillDefault
    Range("A4:C30579").Select
    Range("C30579").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("C30579").Select
    Selection.End(xlUp).Select
    Range("C1").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "Pregunta"
    Range("I1:J1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("I1:XFD7").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1:H2").Select
    Range("H2").Activate
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("G30564").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select



 
End Sub


Sub SERVICIOSFINANCIEROS()

'
' SERVICIOSFINANCIEROS Macro
'

'
    LIMPIAR
    Selection.EntireColumn.Insert
    Range("B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A4").Select
    ActiveSheet.Paste
    Range("C1").Select
    Application.CutCopyMode = False
    Range("C2:F2").Select
    Selection.Cut
    Range("C3").Select
    ActiveSheet.Paste
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("A5:A12").Select
    Selection.ClearContents
    Range("B6").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("C4").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A14").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlUp).Select
    Range("C13").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A32").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A14:A32").Select
    Range("A32").Activate
    Selection.ClearContents
    Range("C33").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A51").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A34:A51").Select
    Range("A51").Activate
    Selection.ClearContents
    Range("C52").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A76").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A53:A76").Select
    Range("A76").Activate
    Selection.ClearContents
    Range("C76").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A78:A84").Select
    Range("A84").Activate
    Selection.ClearContents
    Range("C84").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A106").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A86:A106").Select
    Range("A106").Activate
    Selection.ClearContents
    Range("C106").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A134").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A108:A134").Select
    Range("A134").Activate
    Selection.ClearContents
    Range("C135").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A159").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A136:A159").Select
    Range("A159").Activate
    Selection.ClearContents
    Range("C159").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A194").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A161:A194").Select
    Range("A194").Activate
    Selection.ClearContents
    Range("B193").Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range("A4").Select
    Selection.Copy
    Range("A4:A12").Select
    ActiveSheet.Paste
    Range("A13").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A14").Select
    Selection.End(xlDown).Select
    Range("A32").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A33").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A34").Select
    Selection.End(xlDown).Select
    Range("A51").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A52").Select
    Selection.End(xlDown).Select
    Range("A76").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A53").Select
    Selection.End(xlDown).Select
    Range("A76").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A77").Select
    Selection.End(xlDown).Select
    Range("A77").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A84").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A85").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A86").Select
    Selection.End(xlDown).Select
    Range("A106").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A107").Select
    Selection.End(xlDown).Select
    Range("A133").Select
    Selection.End(xlUp).Select
    Range("A107").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A108").Select
    Selection.End(xlDown).Select
    Range("A134").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A135").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A136").Select
    Selection.End(xlDown).Select
    Range("A159").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A160").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A161").Select
    Selection.End(xlDown).Select
    Range("A1048574").Select
    Selection.End(xlUp).Select
    Range("B165").Select
    Selection.End(xlDown).Select
    Range("A194").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A192").Select
    Selection.End(xlUp).Select
    Selection.End(xlToRight).Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A12").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.End(xlToLeft).Select
    Range("A8").Select
    Selection.End(xlUp).Select
    Range("A12").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C16").Select
    Selection.End(xlDown).Select
    Range("A31").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C31").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A49").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C50").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A73").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C73").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A80").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C80").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A101").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C101").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("A128").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C128").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A152").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C152").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B1048576").Select
    Selection.End(xlUp).Select
    Range("B184").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range("C3").Select
    Selection.EntireColumn.Insert
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("B1").Select
    Selection.Cut
    Range("A3").Select
    ActiveSheet.Paste
    Range("H1").Select
    Selection.Cut
    Range("C3").Select
    ActiveSheet.Paste
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Pregunta"
    Range("C4").Select
    Selection.Copy
    Range("D4").Select
    Selection.End(xlDown).Select
    Range("C185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("D185").Select
    Selection.End(xlUp).Select
    Range("G3").Select
    Application.CutCopyMode = False
    Range("H3").Select
    Selection.EntireColumn.Insert
    Range("M3").Select
    Selection.EntireColumn.Insert
    Range("R3").Select
    Selection.EntireColumn.Insert
    Range("W3").Select
    Selection.EntireColumn.Insert
    Range("AB3").Select
    Selection.EntireColumn.Insert
    Range("AG3").Select
    Selection.EntireColumn.Insert
    Range("AL3").Select
    Selection.EntireColumn.Insert
    Range("AQ3").Select
    Selection.EntireColumn.Insert
    Range("AV3").Select
    Selection.EntireColumn.Insert
    Range("BA3").Select
    Selection.EntireColumn.Insert
    Range("BF3").Select
    Selection.EntireColumn.Insert
    Range("BK3").Select
    Selection.EntireColumn.Insert
    Range("BP3").Select
    Selection.EntireColumn.Insert
    Range("BU3").Select
    Selection.EntireColumn.Insert
    Range("BZ3").Select
    Selection.EntireColumn.Insert
    Range("CE3").Select
    Selection.EntireColumn.Insert
    Range("CJ3").Select
    Selection.EntireColumn.Insert
    Range("CO3").Select
    Selection.EntireColumn.Insert
    Range("CT3").Select
    Range("CT3").Select
    Selection.EntireColumn.Insert
    Range("CY3").Select
    Selection.EntireColumn.Insert
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CI2").Select
    Selection.End(xlToLeft).Select
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Range("BG2").Select
    Selection.Cut
    Selection.End(xlDown).Select
    Range("BG186").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    ActiveSheet.Paste
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Range("AT2").Select
    Selection.Cut
    Selection.End(xlToLeft).Select
    Range("AQ4").Select
    ActiveSheet.Paste
    Selection.End(xlToLeft).Select
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Range("AI2").Select
    Selection.Cut
    Range("AH2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Selection.End(xlToLeft).Select
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Range("Q4").Select
    Selection.End(xlToLeft).Select
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Range("L4").Select
    Selection.End(xlToLeft).Select
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Range("F4").Select
    Selection.End(xlToLeft).Select
    Range("H4").Select
    Selection.Copy
    Range("I4").Select
    Selection.End(xlDown).Select
    Range("H185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("H184").Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("N4").Select
    Selection.End(xlDown).Select
    Range("M185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("M184").Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("S4").Select
    Selection.End(xlDown).Select
    Range("R185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("X4").Select
    Selection.End(xlDown).Select
    Range("W185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("W181").Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AC4").Select
    Selection.End(xlDown).Select
    Range("AB185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AB184").Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AH5").Select
    Selection.End(xlDown).Select
    Range("AG185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AJ183").Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AM4").Select
    Selection.End(xlDown).Select
    Range("AL185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AM184").Select
    Selection.End(xlToRight).Select
    Range("AQ182").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Copy
    Range("AR4").Select
    Selection.End(xlDown).Select
    Range("AQ185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AV185").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AW4").Select
    Selection.End(xlDown).Select
    Range("AV185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AW185").Select
    Selection.End(xlToRight).Select
    Range("BA185").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BB4").Select
    Selection.End(xlDown).Select
    Range("BA185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BF185").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BG4").Select
    Selection.End(xlDown).Select
    Range("BF185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BL185").Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BL4").Select
    Selection.End(xlDown).Select
    Range("BK185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BL185").Select
    Selection.End(xlToRight).Select
    Range("BP185").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BQ4").Select
    Selection.End(xlDown).Select
    Range("BP185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BV185").Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BV4").Select
    Selection.End(xlDown).Select
    Range("BU185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BU184").Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CA4").Select
    Selection.End(xlDown).Select
    Range("BZ185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CA185").Select
    Selection.End(xlToRight).Select
    Range("CD184").Select
    Selection.End(xlToRight).Select
    Range("CF182").Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CF4").Select
    Selection.End(xlDown).Select
    Range("CE185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CF185").Select
    Selection.End(xlToRight).Select
    Range("CK185").Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CK4").Select
    Selection.End(xlDown).Select
    Range("CJ185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CJ184").Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CP4").Select
    Selection.End(xlDown).Select
    Range("CO185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CP185").Select
    Selection.End(xlToRight).Select
    Range("CT184").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CU4").Select
    Selection.End(xlDown).Select
    Range("CT185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CZ185").Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CZ4").Select
    Selection.End(xlDown).Select
    Range("CY185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CZ185").Select
    Selection.End(xlToRight).Select
    Range("DC184").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("DC175").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlToRight).Select
    Range("XFC157").Select
    Selection.End(xlToLeft).Select
    Range("CV157").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY1").Select
    Application.CutCopyMode = False
    Range("CY4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT186").Select
    ActiveSheet.Paste
    Range("CT185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO186").Select
    ActiveSheet.Paste
    Range("CO185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ186").Select
    ActiveSheet.Paste
    Range("CJ185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE186").Select
    ActiveSheet.Paste
    Range("CE185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ186").Select
    ActiveSheet.Paste
    Range("BZ185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU186").Select
    ActiveSheet.Paste
    Range("BU185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP186").Select
    ActiveSheet.Paste
    Range("BP185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK186").Select
    ActiveSheet.Paste
    Range("BK185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF186").Select
    ActiveSheet.Paste
    Range("BF185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA186").Select
    ActiveSheet.Paste
    Range("BA185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV186").Select
    ActiveSheet.Paste
    Range("AV185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ186").Select
    ActiveSheet.Paste
    Range("AQ185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL186").Select
    ActiveSheet.Paste
    Range("AL185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG186").Select
    ActiveSheet.Paste
    Range("AG185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB186").Select
    ActiveSheet.Paste
    Range("AB185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W186").Select
    ActiveSheet.Paste
    Range("W185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R186").Select
    ActiveSheet.Paste
    Range("R185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M186").Select
    ActiveSheet.Paste
    Range("M185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("G4").Select
    Selection.End(xlDown).Select
    Range("H11").Select
    Selection.End(xlDown).Select
    Range("H186").Select
    ActiveSheet.Paste
    Range("H185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C186").Select
    ActiveSheet.Paste
    Range("C186").Select
    Selection.End(xlUp).Select
    Range("C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlUp).Select
    Range("A4:B132").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=18
    Selection.AutoFill Destination:=Range("A4:B3825"), Type:=xlFillDefault
    Range("A4:B3825").Select
    ActiveWindow.SmallScroll Down:=-18
    Range("B3813").Select
    Selection.End(xlUp).Select
    Range("H1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("H1:XFD6").Select
    Selection.Delete Shift:=xlToLeft
    Range("A2").Select
    Selection.EntireColumn.Insert
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "SERVICIOS FIN ANCIEROS"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "SERVICIOS FINANCIEROS"
    Range("A4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("A3825").Select
    ActiveWindow.SmallScroll Down:=9
    Range(Selection, Selection.End(xlUp)).Select
    Range("A5:A3825").Select
    Range("A3825").Activate
    ActiveSheet.Paste
    Range("A3824").Select
    Selection.End(xlUp).Select
    Range("A2").Select
    Application.CutCopyMode = False
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select

    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("B3800").Select
End Sub


Sub TRANSPORTE()

'
' Macro3 Macro
'
' Acceso directo: CTRL+i
'
    LIMPIAR
    Range("H11").Select
    ActiveWindow.SmallScroll Down:=-24
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 37.25
    Range("B5").Select
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=12
    Range("B26").Select
    Selection.Cut
    Range("A27").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=12
    Range("B33").Select
    Selection.Cut
    Range("A34").Select
    ActiveSheet.Paste
    Range("B39").Select
    Selection.Cut
    Range("A40").Select
    ActiveSheet.Paste
    Range("B42").Select
    Selection.Cut
    Range("A43").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("B47").Select
    Selection.Cut
    Range("A48").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=-51
    Rows("4:5").Select
    Selection.Delete Shift:=xlUp
    Range("B6").Select
    ActiveWindow.SmallScroll Down:=18
    Rows("24:24").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=9
    Range("36:36,39:39,44:44").Select
    Range("A44").Activate
    Selection.Delete Shift:=xlUp
    Range("A33").Select
    ActiveWindow.SmallScroll Down:=-45
    Range("A4").Select
    Rows("4:4").RowHeight = 22.5
    Range("A4").Select
    Selection.Copy
    Range("A5").Select
    ActiveWindow.SmallScroll Down:=6
    Range("A5:A23").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=15
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Cut
    Application.CutCopyMode = False
    Selection.Copy
    Range("A25:A30").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=6
    Range("A31").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A32:A35").Select
    ActiveSheet.Paste
    Range("A36").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A37").Select
    ActiveSheet.Paste
    Range("A38").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A39:A41").Select
    ActiveSheet.Paste
    Range("A42").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A43").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlUp).Select
    ActiveWindow.SmallScroll Down:=-27
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 2
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 18.13
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AH2").Select
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    Columns("AQ:AQ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Columns("AV:AV").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    Columns("BA:BA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Columns("BF:BF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    Columns("BK:BK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    Columns("BP:BP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    Columns("BU:BU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Columns("BZ:BZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    Columns("CE:CE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Columns("CJ:CJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Columns("CO:CO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    Columns("CT:CT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Columns("CY:CY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C50").Select
    ActiveWindow.SmallScroll Down:=-18
    Application.CutCopyMode = False
    ActiveWindow.SmallScroll Down:=-15
    Rows("30:30").Select
    Selection.Delete Shift:=xlUp
    Range("A37").Select
    ActiveWindow.SmallScroll Down:=-6
    Rows("24:29").Select
    Selection.RowHeight = 15.75
    Range("A31").Select
    ActiveWindow.SmallScroll Down:=-36
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C5:C42").Select
    Range("C42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G4").Select
    Selection.End(xlDown).Select
    Range("H42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H42").Select
    Range("H42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L4").Select
    Selection.End(xlDown).Select
    Range("M42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M42").Select
    Range("M42").Activate
    ActiveSheet.Paste
    Range("C4").Select
    Application.CutCopyMode = False
    Range("R4").Select
    Selection.Copy
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("R42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R5:R42").Select
    Range("R42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("R4,R6").Select
    Range("R6").Activate
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("W42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W5:W42").Select
    Range("W42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("AB42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB5:AB42").Select
    Range("AB42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AG42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG5:AG42").Select
    Range("AG42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AL42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL5:AL42").Select
    Range("AL42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AQ42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ5:AQ42").Select
    Range("AQ42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AV42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AV5:AV42").Select
    Range("AV42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("BA42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA5:BA42").Select
    Range("BA42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BF42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF5:BF42").Select
    Range("BF42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BF5").Select
    Selection.End(xlDown).Select
    Range("BK42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK5:BK42").Select
    Range("BK42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BP42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP5:BP42").Select
    Range("BP42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BU42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU5:BU42").Select
    Range("BU42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BZ42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ5:BZ42").Select
    Range("BZ42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("CE42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE5:CE42").Select
    Range("CE42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CJ42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ5:CJ42").Select
    Range("CJ42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CO42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO5:CO42").Select
    Range("CO42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CT42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT5:CT42").Select
    Range("CT42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CY42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY5:CY42").Select
    Range("CY42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    ActiveWindow.ScrollColumn = 90
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("D1:G1").Select
    Selection.Cut
    Range("D2").Select
    ActiveSheet.Paste
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("I1").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    Range("CY2:DC2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT2").Select
    Selection.End(xlDown).Select
    Range("CT41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO2").Select
    Selection.End(xlDown).Select
    Range("CO41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ2").Select
    Selection.End(xlDown).Select
    Range("CJ41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE2").Select
    Selection.End(xlDown).Select
    Range("CE41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ2").Select
    Selection.End(xlDown).Select
    Range("BZ41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU2").Select
    Selection.End(xlDown).Select
    Range("BU41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP2").Select
    Selection.End(xlDown).Select
    Range("BP41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK2").Select
    Selection.End(xlDown).Select
    Range("BK41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF2").Select
    Selection.End(xlDown).Select
    Range("BF41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA2").Select
    Selection.End(xlDown).Select
    Range("BA41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV2").Select
    Selection.End(xlDown).Select
    Range("AV41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ2").Select
    Selection.End(xlDown).Select
    Range("AQ41").Select
    ActiveSheet.Paste
    Range("AQ40").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL2").Select
    Selection.End(xlDown).Select
    Range("AL41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG2").Select
    Selection.End(xlDown).Select
    Range("AG41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB2").Select
    Selection.End(xlDown).Select
    Range("AB41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W2").Select
    Selection.End(xlDown).Select
    Range("W41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R2").Select
    Selection.End(xlDown).Select
    Range("R41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M2").Select
    Selection.End(xlDown).Select
    Range("M41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H2").Select
    Selection.End(xlDown).Select
    Range("H41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C2").Select
    Selection.End(xlDown).Select
    Range("C41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("C820").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("C2").Select
    Selection.End(xlDown).Select
    Range("A820:B820").Select
    Range("B820").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A41:B820").Select
    Range("B820").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 90
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variables"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "EQUIPAMIENTO"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A820").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A820").Select
    Range("A820").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("A:A").ColumnWidth = 23.13
    Range("A2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    Range("C1").Select
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2").Select
    Rows("2:2").RowHeight = 17
    Columns("C:C").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Selection.End(xlDown).Select
    Range("A820").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
End Sub



Sub MEDIOSDIAAYER()
'
' MEDIOSDIAAYER Macro
'
' Acceso directo: CTRL+t
'
    LIMPIARv2
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("Q:Q").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    Columns("V:V").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AA:AA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    Columns("AF:AF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AK:AK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    Columns("AP:AP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AU:AU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    Columns("AZ:AZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BE:BE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    Columns("BJ:BJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BO:BO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    Columns("BT:BT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BY:BY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    Columns("CD:CD").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CI:CI").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    Columns("CN:CN").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CS:CS").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    Columns("CX:CX").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("DA23").Select
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("H1").Select
    Selection.Cut
    Range("G4").Select
    ActiveSheet.Paste
    Range("M2").Select
    Selection.Cut
    Range("L4").Select
    ActiveSheet.Paste
    Range("R2").Select
    Selection.Cut
    Range("Q4").Select
    ActiveSheet.Paste
    Range("W2").Select
    Selection.Cut
    Range("V4").Select
    ActiveSheet.Paste
    Range("AB2").Select
    Selection.Cut
    Range("AA4").Select
    ActiveSheet.Paste
    Range("AG2").Select
    Selection.Cut
    Range("AF4").Select
    ActiveSheet.Paste
    Range("AL2").Select
    Selection.Cut
    Range("AK4").Select
    ActiveSheet.Paste
    Range("AQ2").Select
    Selection.Cut
    Range("AP4").Select
    ActiveSheet.Paste
    Range("AV2").Select
    Selection.Cut
    Range("AU4").Select
    ActiveSheet.Paste
    Range("BA2").Select
    Selection.Cut
    Range("AZ4").Select
    ActiveSheet.Paste
    Range("BF2").Select
    Selection.Cut
    Range("BE4").Select
    ActiveSheet.Paste
    Range("BK2").Select
    Selection.Cut
    Range("BJ4").Select
    ActiveSheet.Paste
    Range("BP2").Select
    Selection.Cut
    Range("BO4").Select
    ActiveSheet.Paste
    Range("BU2").Select
    Selection.Cut
    Range("BT4").Select
    ActiveSheet.Paste
    Range("BZ2").Select
    Selection.Cut
    Range("BY4").Select
    ActiveSheet.Paste
    Range("CE2").Select
    Selection.Cut
    Range("CD4").Select
    ActiveSheet.Paste
    Range("CJ2").Select
    Selection.Cut
    Selection.End(xlDown).Select
    Range("CJ86").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CI4").Select
    ActiveSheet.Paste
    Range("CO2").Select
    Selection.Cut
    Range("CN4").Select
    ActiveSheet.Paste
    Range("CT2").Select
    Selection.Cut
    Range("CS4").Select
    ActiveSheet.Paste
    Range("CY2").Select
    Selection.Cut
    Range("CX4").Select
    ActiveSheet.Paste
    Selection.End(xlToLeft).Select
    Range("C1").Select
    Selection.Cut
    Range("B4").Select
    ActiveSheet.Paste
    Range("B4").Select
    Selection.Copy
    Range("A5").Select
    Selection.End(xlDown).Select
    Range("B120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("B5:B120").Select
    Range("B120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("G4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("G120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("G5:G120").Select
    Range("G120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("L4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G4").Select
    Selection.End(xlDown).Select
    Range("L120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("L5:L120").Select
    Range("L120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("Q4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L4").Select
    Selection.End(xlDown).Select
    Range("Q120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("Q5:Q120").Select
    Range("Q120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("V4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("Q4").Select
    Selection.End(xlDown).Select
    Range("V120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("V5:V120").Select
    Range("V120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("AA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("V4").Select
    Selection.End(xlDown).Select
    Range("AA120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AA5:AA120").Select
    Range("AA120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AA4").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("AF120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AF5:AF120").Select
    Range("AF120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("AK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AF4").Select
    Selection.End(xlDown).Select
    Range("AK120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AK5:AK120").Select
    Range("AK120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AK4").Select
    Selection.End(xlDown).Select
    Range("AP120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AP5:AP120").Select
    Range("AP120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AP4").Select
    Selection.End(xlDown).Select
    Range("AU120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AU5:AU120").Select
    Range("AU120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("AZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AU4").Select
    Selection.End(xlDown).Select
    Range("AZ120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AZ5:AZ120").Select
    Range("AZ120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AZ4").Select
    Selection.End(xlDown).Select
    Range("BE120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BE5:BE120").Select
    Range("BE120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BE4").Select
    Selection.End(xlDown).Select
    Range("BJ120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BJ5:BJ120").Select
    Range("BJ120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BJ4").Select
    Selection.End(xlDown).Select
    Range("BO120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BO5:BO120").Select
    Range("BO120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BO4").Select
    Selection.End(xlDown).Select
    Range("BT120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BT5:BT120").Select
    Range("BT120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BT4").Select
    Selection.End(xlDown).Select
    Range("BY120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BY5:BY120").Select
    Range("BY120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CD4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BY4").Select
    Selection.End(xlDown).Select
    Range("CD120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CD5:CD120").Select
    Range("CD120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CI4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CD4").Select
    Selection.End(xlDown).Select
    Range("CI120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CI5:CI120").Select
    Range("CI120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CN4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CI4").Select
    Selection.End(xlDown).Select
    Range("CN120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CN5:CN120").Select
    Range("CN120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CS4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CN4").Select
    Selection.End(xlDown).Select
    Range("CS120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CS5:CS120").Select
    Range("CS120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CX4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CS4").Select
    Selection.End(xlDown).Select
    Range("CX120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CX5:CX120").Select
    Range("CX120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CX4").Select
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("CX4").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Range("G4").Select
    Selection.End(xlToRight).Select
    Range("CX4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CS4").Select
    Selection.End(xlDown).Select
    Range("CS120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CS4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CN4").Select
    Selection.End(xlDown).Select
    Range("CN120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CN4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CI4").Select
    Selection.End(xlDown).Select
    Range("CI120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CI4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CD4").Select
    Selection.End(xlDown).Select
    Range("CD120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CD4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BY4").Select
    Selection.End(xlDown).Select
    Range("BY120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BY4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BT4").Select
    Selection.End(xlDown).Select
    Range("BT120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BT4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BO4").Select
    Selection.End(xlDown).Select
    Range("BO120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BO4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BJ4").Select
    Selection.End(xlDown).Select
    Range("BJ120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BJ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BE6").Select
    Selection.End(xlDown).Select
    Range("BE120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BE4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AZ4").Select
    Selection.End(xlDown).Select
    Range("AZ120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AZ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AU4").Select
    Selection.End(xlDown).Select
    Range("AU120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AU4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AP4").Select
    Selection.End(xlDown).Select
    Range("AP120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AP4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AK4").Select
    Selection.End(xlDown).Select
    Range("AK120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AK4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AF4").Select
    Selection.End(xlDown).Select
    Range("AF120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AF4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AA4").Select
    Selection.End(xlDown).Select
    Range("AA120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AA4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("V4").Select
    Selection.End(xlDown).Select
    Range("V120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("V4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("Q4").Select
    Selection.End(xlDown).Select
    Range("Q120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("Q4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("L4").Select
    Selection.End(xlDown).Select
    Range("L120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("L4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("G4").Select
    Selection.End(xlDown).Select
    Range("G120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("G4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("B120").Select
    ActiveSheet.Paste
    Range("A119").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A4:A119").Select
    Range("A119").Activate
    Selection.Copy
    Range("B120").Select
    Selection.End(xlDown).Select
    Range("A2439").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A120:A2439").Select
    Range("A2439").Activate
    ActiveSheet.Paste
    Range("A2439").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("G:G").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    Columns("G:DB").Select
    Selection.Delete Shift:=xlToLeft
    Rows("3:3").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Television el dia de ayer"
    Range("B2").Select
    Selection.Copy
    Range("C2").Select
    Selection.End(xlDown).Select
    Range("B2437").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("B3:B2437").Select
    Range("B2437").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "MEDIOS"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A2437").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A2437").Select
    Range("A2437").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("C1").Select
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Columns("C:C").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Range("A2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    Range("B9").Select
    ActiveWindow.SmallScroll Down:=-15
    ActiveWorkbook.Save
End Sub



Sub TELEFONIA()

'
' Macro6 Macro
'
' Acceso directo: CTRL+e
'
    LIMPIAR
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 40.75
    Range("B5").Select
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    Range("B8").Select
    Selection.Cut
    Range("A9").Select
    ActiveSheet.Paste
    Range("B14").Select
    Selection.Cut
    Range("A15").Select
    ActiveSheet.Paste
    Range("B20").Select
    Selection.Cut
    Range("A21").Select
    ActiveSheet.Paste
    Range("B27").Select
    Selection.Cut
    Range("A28").Select
    ActiveSheet.Paste
    Range("B30").Select
    Selection.Cut
    Range("A31").Select
    ActiveSheet.Paste
    Range("B34").Select
    Selection.Cut
    Range("A35").Select
    ActiveSheet.Paste
    Range("B58").Select
    Selection.Cut
    Range("A59").Select
    ActiveSheet.Paste
    Range("B62").Select
    ActiveWindow.SmallScroll Down:=-90
    Rows("4:5").Select
    Selection.Delete Shift:=xlUp
    Rows("6:6").Select
    Selection.Delete Shift:=xlUp
    Rows("11:11").Select
    Selection.Delete Shift:=xlUp
    Rows("16:16").Select
    Selection.Delete Shift:=xlUp
    Range("B17").Select
    ActiveWindow.SmallScroll Down:=12
    Rows("22:22").Select
    Selection.Delete Shift:=xlUp
    Rows("24:24").Select
    Selection.Delete Shift:=xlUp
    Range("B20").Select
    ActiveWindow.SmallScroll Down:=9
    Rows("27:27").Select
    Selection.Delete Shift:=xlUp
    Range("A36").Select
    ActiveWindow.SmallScroll Down:=9
    Rows("50:50").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=3
    Range("B49").Select
    ActiveWindow.SmallScroll Down:=-72
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AQ:AQ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    Columns("AV:AV").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BA:BA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BF:BF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    Columns("BK:BK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BP:BP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    Columns("BU:BU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BZ:BZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    Columns("CE:CE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CJ:CJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CO:CO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    Columns("CT:CT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CY:CY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("I1").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Range("AH2").Select
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C5:C63").Select
    Range("C63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("C4").Select
    Application.CutCopyMode = False
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("H63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H63").Select
    Range("H63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("M63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M63").Select
    Range("M63").Activate
    ActiveSheet.Paste
    Range("M62").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("R63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R5:R63").Select
    Range("R63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("W63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W5:W63").Select
    Range("W63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("AB4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("AB63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB5:AB63").Select
    Range("AB63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AG63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG5:AG63").Select
    Range("AG63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AL63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL5:AL63").Select
    Range("AL63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AQ63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ5:AQ63").Select
    Range("AQ63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AV63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AV5:AV63").Select
    Range("AV63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("BA63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA5:BA63").Select
    Range("BA63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BF63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF5:BF63").Select
    Range("BF63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BK63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK5:BK63").Select
    Range("BK63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BP63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP5:BP63").Select
    Range("BP63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BU63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU5:BU63").Select
    Range("BU63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BZ63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ5:BZ63").Select
    Range("BZ63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("CE63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE5:CE63").Select
    Range("CE63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CJ63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ5:CJ63").Select
    Range("CJ63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CO63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO5:CO63").Select
    Range("CO63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CT63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT5:CT63").Select
    Range("CT63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CY63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY5:CY63").Select
    Range("CY63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("A4").Select
    Selection.Copy
    Range("A5").Select
    ActiveSheet.Paste
    Range("A6").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A7:A10").Select
    ActiveSheet.Paste
    Range("A11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A12:A15").Select
    ActiveSheet.Paste
    Range("A16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A17:A21").Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A23").Select
    ActiveSheet.Paste
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A25:A26").Select
    ActiveSheet.Paste
    Range("A27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A28:A49").Select
    ActiveSheet.Paste
    Range("A30").Select
    Selection.End(xlDown).Select
    Range("A50").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A51:A63").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A4:B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlUp).Select
    Range("A4").Select
    Application.CutCopyMode = False
    Range("A4:B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("A1263:B1263").Select
    Range("B1263").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A64:B1263").Select
    Range("B1263").Activate
    ActiveSheet.Paste
    Range("B1263").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Rows("3:3").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Telefonia"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "TELEFONIA"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A1261").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A1261").Select
    Range("A1261").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("A2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    Range("C1").Select
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A9").Select
    ActiveWindow.SmallScroll Down:=-18
    Selection.End(xlDown).Select
    Range("A1261").Select
    Selection.End(xlUp).Select
    Range("H1").Select
    ActiveWorkbook.Save
End Sub





Sub INTERNET()

'
' Internet Macro
'
' Acceso directo: CTRL+r
'
    LIMPIAR
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B5").Select
    Columns("A:A").ColumnWidth = 88.13
    Columns("A:A").ColumnWidth = 66.13
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    Range("B8").Select
    Selection.Cut
    Range("B8").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("A9").Select
    ActiveSheet.Paste
    Range("B14").Select
    Selection.Cut
    Range("A15").Select
    ActiveSheet.Paste
    Range("B24").Select
    Selection.Cut
    Range("A25").Select
    ActiveSheet.Paste
    Range("C25").Select
    Selection.End(xlDown).Select
    Range("B33").Select
    Selection.Cut
    Range("A34").Select
    ActiveSheet.Paste
    Range("C34").Select
    Selection.End(xlDown).Select
    Range("B36").Select
    Selection.Cut
    Range("A37").Select
    ActiveSheet.Paste
    Range("C37").Select
    Selection.End(xlDown).Select
    Range("B52").Select
    Selection.Cut
    Range("A53").Select
    ActiveSheet.Paste
    Range("B55").Select
    Selection.Cut
    Range("A56").Select
    ActiveSheet.Paste
    Range("C56").Select
    Selection.End(xlDown).Select
    Range("B60").Select
    Selection.Cut
    Range("A61").Select
    ActiveSheet.Paste
    Range("C61").Select
    Selection.End(xlDown).Select
    Range("B68").Select
    Selection.Cut
    Range("A69").Select
    ActiveSheet.Paste
    Range("B76").Select
    Selection.Cut
    Range("A77").Select
    ActiveSheet.Paste
    Range("B81").Select
    Selection.Cut
    Range("A82").Select
    ActiveSheet.Paste
    Range("C82").Select
    Selection.End(xlDown).Select
    Range("B84").Select
    Selection.Cut
    Range("A85").Select
    ActiveSheet.Paste
    Range("B88").Select
    Selection.Cut
    Range("A89").Select
    ActiveSheet.Paste
    Range("B91").Select
    Selection.Cut
    Range("A92").Select
    ActiveSheet.Paste
    Range("B96").Select
    Selection.Cut
    Range("A97").Select
    ActiveSheet.Paste
    Range("B99").Select
    Selection.Cut
    Range("A100").Select
    ActiveSheet.Paste
    Range("B110").Select
    Selection.Cut
    Range("A111").Select
    ActiveSheet.Paste
    Range("B114").Select
    Selection.Cut
    Range("A115").Select
    ActiveSheet.Paste
    Range("C116").Select
    Selection.End(xlDown).Select
    Range("B119").Select
    Selection.Cut
    Range("A120").Select
    ActiveSheet.Paste
    Range("B122").Select
    Selection.End(xlDown).Select
    Range("A196").Select
    Selection.End(xlUp).Select
    Range("B134").Select
    Selection.Cut
    Range("A135").Select
    ActiveSheet.Paste
    Range("B143").Select
    Selection.Cut
    Range("A144").Select
    ActiveSheet.Paste
    Range("B149").Select
    Selection.Cut
    Range("A150").Select
    ActiveSheet.Paste
    Range("B157").Select
    Selection.Cut
    Range("A158").Select
    ActiveSheet.Paste
    Range("B165").Select
    Selection.Cut
    Range("A166").Select
    ActiveSheet.Paste
    Range("B173").Select
    Selection.Cut
    Range("A174").Select
    ActiveSheet.Paste
    Range("B181").Select
    Selection.Cut
    Range("A182").Select
    ActiveSheet.Paste
    Range("B189").Select
    Selection.Cut
    Range("A190").Select
    ActiveSheet.Paste
    Range("B190").Select
    ActiveWindow.ScrollRow = 177
    ActiveWindow.ScrollRow = 176
    ActiveWindow.ScrollRow = 174
    ActiveWindow.ScrollRow = 166
    ActiveWindow.ScrollRow = 160
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 1
    Range("A6").Select
    Selection.Copy
    Range("A7:A8").Select
    ActiveSheet.Paste
    Range("A9").Select
    Application.CutCopyMode = False
    Selection.Cut
    Application.CutCopyMode = False
    Selection.Copy
    Range("A10:A14").Select
    ActiveSheet.Paste
    Range("A15").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A16").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A16:A24").Select
    ActiveSheet.Paste
    Range("A25").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A26").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A26:A33").Select
    ActiveSheet.Paste
    Range("A34").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A35:A36").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A37").Select
    Selection.Copy
    Range("A38").Select
    Selection.End(xlDown).Select
    Range("A52").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A38:A52").Select
    Range("A52").Activate
    ActiveSheet.Paste
    Range("A53").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A54:A55").Select
    ActiveSheet.Paste
    Range("A56").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A57").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A57:A60").Select
    ActiveSheet.Paste
    Range("A61").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A62").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A62:A68").Select
    ActiveSheet.Paste
    Range("A69").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A70").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A70:A76").Select
    ActiveSheet.Paste
    Range("A77").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A78").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A78:A81").Select
    ActiveSheet.Paste
    Range("A82").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A83").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A83:A84").Select
    ActiveSheet.Paste
    Range("A85").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A86").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A86:A88").Select
    ActiveSheet.Paste
    Range("A89").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A90:A91").Select
    ActiveSheet.Paste
    Range("A92").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A93").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A93:A96").Select
    ActiveSheet.Paste
    Range("A97").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A98:A99").Select
    ActiveSheet.Paste
    Range("A100").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A101").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A101:A110").Select
    ActiveSheet.Paste
    Range("A111").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A112:A114").Select
    ActiveSheet.Paste
    Range("A115").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A116").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A116:A119").Select
    ActiveSheet.Paste
    Range("A120").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A121").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A121:A134").Select
    ActiveSheet.Paste
    Range("A135").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A136").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A136:A143").Select
    ActiveSheet.Paste
    Range("A144").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A145").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A145:A149").Select
    ActiveSheet.Paste
    Range("A150").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Range("A151").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A151:A157").Select
    ActiveSheet.Paste
    Range("A158").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A159").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A159:A165").Select
    ActiveSheet.Paste
    Range("A166").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A167").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A167:A173").Select
    ActiveSheet.Paste
    Range("A174").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A175").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A175:A181").Select
    ActiveSheet.Paste
    Range("A182").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A183").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A183:A189").Select
    ActiveSheet.Paste
    Range("A190").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A191").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A191:A196").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Rows("4:5").Select
    Selection.Delete Shift:=xlUp
    Rows("6:6").Select
    Selection.Delete Shift:=xlUp
    Rows("11:11").Select
    Selection.Delete Shift:=xlUp
    Rows("20:20").Select
    Selection.Delete Shift:=xlUp
    Rows("28:28").Select
    Selection.Delete Shift:=xlUp
    Rows("30:30").Select
    Selection.Delete Shift:=xlUp
    Rows("45:45").Select
    Selection.Delete Shift:=xlUp
    Rows("47:47").Select
    Selection.Delete Shift:=xlUp
    Rows("51:51").Select
    Selection.Delete Shift:=xlUp
    Rows("58:58").Select
    Selection.Delete Shift:=xlUp
    Rows("70:70").Select
    Selection.Delete Shift:=xlUp
    Rows("72:72").Select
    Selection.Delete Shift:=xlUp
    Rows("75:75").Select
    Selection.Delete Shift:=xlUp
    Rows("77:77").Select
    Selection.Delete Shift:=xlUp
    Rows("81:81").Select
    Selection.Delete Shift:=xlUp
    Rows("83:83").Select
    Selection.Delete Shift:=xlUp
    Rows("93:93").Select
    Selection.Delete Shift:=xlUp
    Rows("96:96").Select
    Selection.Delete Shift:=xlUp
    Rows("100:100").Select
    Selection.Delete Shift:=xlUp
    Rows("114:114").Select
    Selection.Delete Shift:=xlUp
    Rows("122:122").Select
    Selection.Delete Shift:=xlUp
    Rows("127:127").Select
    Selection.Delete Shift:=xlUp
    Rows("134:134").Select
    Selection.Delete Shift:=xlUp
    Rows("141:141").Select
    Selection.Delete Shift:=xlUp
    Rows("148:148").Select
    Selection.Delete Shift:=xlUp
    Rows("155:155").Select
    Selection.Delete Shift:=xlUp
    Rows("162:162").Select
    Selection.Delete Shift:=xlUp
    Range("C168").Select
    Selection.End(xlUp).Select
    Rows("65:65").Select
    Selection.Delete Shift:=xlUp
    Range("C65").Select
    Selection.End(xlUp).Select
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AQ:AQ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    Columns("AV:AV").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BA:BA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    Columns("BF:BF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BK:BK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    Columns("BP:BP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BU:BU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    Columns("BZ:BZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CE:CE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    Columns("CJ:CJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CO:CO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CT:CT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    Columns("CY:CY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 1
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Range("N2").Select
    Selection.Cut
    Selection.End(xlDown).Select
    Range("N168").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    ActiveSheet.Paste
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Range("AH2").Select
    Selection.Cut
    Selection.End(xlDown).Select
    Range("AH168").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Range("H4").Select
    Selection.End(xlToLeft).Select
    Range("C4").Select
    Selection.Copy
    Range("B5").Select
    Selection.End(xlDown).Select
    Range("C167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C5:C167").Select
    Range("C167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G5").Select
    Selection.End(xlDown).Select
    Range("C9").Select
    Selection.End(xlDown).Select
    Range("H167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H167").Select
    Range("H167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("M167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M167").Select
    Range("M167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("R167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R5:R167").Select
    Range("R167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("W167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W5:W167").Select
    Range("W167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("AB167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB5:AB167").Select
    Range("AB167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AG167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG5:AG167").Select
    Range("AG167").Activate
    ActiveSheet.Paste
    Range("AG166").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AL167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL5:AL167").Select
    Range("AL167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AQ167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ5:AQ167").Select
    Range("AQ167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AV167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AV5:AV167").Select
    Range("AV167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("BA167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA5:BA167").Select
    Range("BA167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BF167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF5:BF167").Select
    Range("BF167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BK167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK5:BK167").Select
    Range("BK167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BP167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP5:BP167").Select
    Range("BP167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BU167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU5:BU167").Select
    Range("BU167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BZ167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ5:BZ167").Select
    Range("BZ167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("CE167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE5:CE167").Select
    Range("CE167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CJ167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ5:CJ167").Select
    Range("CJ167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CO167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO5:CO167").Select
    Range("CO167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CT167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT5:CT167").Select
    Range("CT167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CY167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY5:CY167").Select
    Range("CY167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("M4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C168").Select
    ActiveSheet.Paste
    Range("B167").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A4:B167").Select
    Range("B167").Activate
    Selection.Copy
    Range("C169").Select
    Selection.End(xlDown).Select
    Range("A3447:B3447").Select
    Range("B3447").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A168:B3447").Select
    Range("B3447").Activate
    ActiveSheet.Paste
    Range("C3447").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Rows("3:3").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "VARIABLE"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "INTERNET"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A3445").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A3445").Select
    Range("A3445").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Range("C1").Select
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    Sheets("INTERNET").Select
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A1").Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlDown).Select
    Range("A3445").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    Range("B7").Select
    ActiveWorkbook.Save
End Sub




Sub MEDIOSU30()

'
' Macro3 Macro
'
' Acceso directo: CTRL+w
'
    LIMPIAR
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 44.13
    Range("B5").Select
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    Range("B13").Select
    Selection.Cut
    Range("A14").Select
    ActiveSheet.Paste
    Range("B25").Select
    Selection.Cut
    Range("A26").Select
    ActiveSheet.Paste
    Range("B28").Select
    Selection.Cut
    Range("A29").Select
    ActiveSheet.Paste
    Range("B30").Select
    Selection.Cut
    Range("A31").Select
    ActiveSheet.Paste
    Range("C31").Select
    Selection.End(xlDown).Select
    Range("B79").Select
    Selection.Cut
    Range("A80").Select
    ActiveSheet.Paste
    Range("C132").Select
    Selection.End(xlDown).Select
    Range("B191").Select
    Selection.Cut
    Range("A192").Select
    ActiveSheet.Paste
    Range("C193").Select
    Selection.End(xlDown).Select
    Range("B217").Select
    Selection.Cut
    Range("A218").Select
    ActiveSheet.Paste
    Range("C218").Select
    Selection.End(xlDown).Select
    Range("B343").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Rows("11:11").Select
    Selection.Delete Shift:=xlUp
    Rows("22:22").Select
    Selection.Delete Shift:=xlUp
    Rows("24:24").Select
    Selection.Delete Shift:=xlUp
    Rows("25:25").Select
    Selection.Delete Shift:=xlUp
    Range("A57").Select
    Selection.End(xlDown).Select
    Rows("73:73").Select
    Selection.Delete Shift:=xlUp
    Range("A88").Select
    Selection.End(xlDown).Select
    Rows("184:184").Select
    Selection.Delete Shift:=xlUp
    Range("A184").Select
    Selection.End(xlDown).Select
    Rows("209:209").Select
    Selection.Delete Shift:=xlUp
    Range("A209").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("B209").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A4").Select
    Selection.Copy
    Range("A5:A10").Select
    ActiveSheet.Paste
    Range("A11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A12:A21").Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A23").Select
    ActiveSheet.Paste
    Range("A25").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A26").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A26:A72").Select
    ActiveSheet.Paste
    Range("A27").Select
    Selection.End(xlDown).Select
    Range("A73").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A74").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A74:A183").Select
    ActiveSheet.Paste
    Range("A77").Select
    Selection.End(xlDown).Select
    Range("A184").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A185").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A185:A208").Select
    ActiveSheet.Paste
    Range("A187").Select
    Selection.End(xlDown).Select
    Range("A209").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A210").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("B210").Select
    Selection.End(xlDown).Select
    Range("A334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A210:A334").Select
    Range("A334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AQ:AQ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AV:AV").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BA:BA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BF:BF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BK:BK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BP:BP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BU:BU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BZ:BZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CE:CE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CJ:CJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CO:CO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CT:CT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CY:CY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Range("AH2").Select
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Range("BL2").Select
    Selection.Cut
    Selection.End(xlDown).Select
    Range("BK4").Select
    ActiveSheet.Paste
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Range("CP1").Select
    Selection.End(xlToLeft).Select
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C5:C334").Select
    Range("C334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("H334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H334").Select
    Range("H334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("M334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M334").Select
    Range("M334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("R334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R5:R334").Select
    Range("R334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("W334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W5:W334").Select
    Range("W334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("AB334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB5:AB334").Select
    Range("AB334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AG334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG5:AG334").Select
    Range("AG334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AL334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL5:AL334").Select
    Range("AL334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AQ334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ5:AQ334").Select
    Range("AQ334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AV334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AV5:AV334").Select
    Range("AV334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("BA334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA5:BA334").Select
    Range("BA334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BF334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF5:BF334").Select
    Range("BF334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BK334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK5:BK334").Select
    Range("BK334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BP334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP5:BP334").Select
    Range("BP334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BU334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU5:BU334").Select
    Range("BU334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BZ334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ5:BZ334").Select
    Range("BZ334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("CE334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE5:CE334").Select
    Range("CE334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CJ334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ5:CJ334").Select
    Range("CJ334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CO334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO5:CO334").Select
    Range("CO334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CT334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT5:CT334").Select
    Range("CT334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CY334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY5:CY334").Select
    Range("CY334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("DF4").Select
    Application.CutCopyMode = False
    Range("CY4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("S4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("R4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C335").Select
    ActiveSheet.Paste
    Range("A334:B334").Select
    Range("B334").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A4:B334").Select
    Range("B334").Activate
    Selection.Copy
    Range("C335").Select
    Selection.End(xlDown).Select
    Range("A6954:B6954").Select
    Range("B6954").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A335:B6954").Select
    Range("B6954").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 90
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Rows("3:3").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "MEDIOS"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A6952").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A6952").Select
    Range("A6952").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("C1").Select
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    Range("B10").Select
    Columns("C:C").EntireColumn.AutoFit
    Columns("B:B").ColumnWidth = 39.5
    Range("A18").Select
    Selection.End(xlDown).Select
    Range("A6952").Select
    Selection.End(xlUp).Select
    ActiveWorkbook.Save
End Sub





Sub MEDIOSUPTOTAL()

'
' Macro2 Macro
'
' Acceso directo: CTRL+q
'
    LIMPIAR
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B5").Select
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    Range("B13").Select
    Selection.Cut
    Range("A14").Select
    ActiveSheet.Paste
    Range("C20").Select
    Selection.End(xlDown).Select
    Range("B25").Select
    Selection.Cut
    Range("A26").Select
    ActiveSheet.Paste
    Range("B28").Select
    Selection.Cut
    Range("A29").Select
    ActiveSheet.Paste
    Range("C48").Select
    Selection.End(xlDown).Select
    Range("B69").Select
    Selection.Cut
    Range("A70").Select
    ActiveSheet.Paste
    Range("A60").Select
    Selection.End(xlUp).Select
    Range("B30").Select
    Selection.Cut
    Range("A31").Select
    ActiveSheet.Paste
    Range("C32").Select
    Selection.End(xlDown).Select
    Range("C76").Select
    Selection.End(xlDown).Select
    Range("B160").Select
    Selection.Cut
    Range("A161").Select
    ActiveSheet.Paste
    Range("C161").Select
    Selection.End(xlDown).Select
    Range("B287").Select
    Selection.Cut
    Range("A288").Select
    ActiveSheet.Paste
    Range("C288").Select
    Selection.End(xlDown).Select
    Range("B308").Select
    Selection.Cut
    Range("A309").Select
    ActiveSheet.Paste
    Range("B315").Select
    Selection.Cut
    Range("A316").Select
    ActiveSheet.Paste
    Range("B325").Select
    Selection.Cut
    Range("A326").Select
    ActiveSheet.Paste
    Range("B332").Select
    Selection.Cut
    Range("A333").Select
    ActiveSheet.Paste
    Range("B340").Select
    Selection.Cut
    Range("A341").Select
    ActiveSheet.Paste
    Range("C342").Select
    Selection.End(xlDown).Select
    Range("B353").Select
    Selection.Cut
    Range("A354").Select
    ActiveSheet.Paste
    Range("B365").Select
    Selection.Cut
    Range("A366").Select
    ActiveSheet.Paste
    Range("B378").Select
    Selection.Cut
    Range("A379").Select
    ActiveSheet.Paste
    Range("B388").Select
    Selection.Cut
    Range("A389").Select
    ActiveSheet.Paste
    Range("B396").Select
    Selection.Cut
    Range("A397").Select
    ActiveSheet.Paste
    Range("B404").Select
    Selection.Cut
    Range("A405").Select
    ActiveSheet.Paste
    Range("B412").Select
    Selection.Cut
    Range("A413").Select
    ActiveSheet.Paste
    Range("B423").Select
    Selection.Cut
    Range("A424").Select
    ActiveSheet.Paste
    Range("B434").Select
    ActiveWindow.ScrollRow = 410
    ActiveWindow.ScrollRow = 408
    ActiveWindow.ScrollRow = 394
    ActiveWindow.ScrollRow = 336
    ActiveWindow.ScrollRow = 225
    ActiveWindow.ScrollRow = 224
    ActiveWindow.ScrollRow = 210
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 85
    ActiveWindow.ScrollRow = 83
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 76
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 1
    Columns("A:A").ColumnWidth = 21.5
    Rows("4:5").Select
    Range("A5").Activate
    Selection.Delete Shift:=xlUp
    Range("11:11,23:23").Select
    Range("A23").Activate
    Selection.Delete Shift:=xlUp
    Range("C16").Select
    ActiveWindow.SmallScroll Down:=18
    Range("24:24,26:26").Select
    Range("A26").Activate
    Selection.Delete Shift:=xlUp
    Rows("63:63").Select
    Selection.Delete Shift:=xlUp
    Range("A63").Select
    Selection.End(xlDown).Select
    Rows("153:153").Select
    Selection.Delete Shift:=xlUp
    Range("A153").Select
    Selection.End(xlDown).Select
    Rows("279:279").Select
    Selection.Delete Shift:=xlUp
    Range("A279").Select
    Selection.End(xlDown).Select
    Rows("299:299").Select
    Selection.Delete Shift:=xlUp
    Rows("305:305").Select
    Selection.Delete Shift:=xlUp
    Range("A305").Select
    Selection.End(xlDown).Select
    Rows("314:314").Select
    Selection.Delete Shift:=xlUp
    Range("B315").Select
    Selection.End(xlDown).Select
    Range("320:320,328:328").Select
    Range("A328").Activate
    Selection.Delete Shift:=xlUp
    Rows("339:339").Select
    Selection.Delete Shift:=xlUp
    Range("A339").Select
    Selection.End(xlDown).Select
    Rows("350:350").Select
    Selection.Delete Shift:=xlUp
    Rows("362:362").Select
    Selection.Delete Shift:=xlUp
    Rows("371:371").Select
    Selection.Delete Shift:=xlUp
    Range("378:378,386:386").Select
    Range("A386").Activate
    Selection.Delete Shift:=xlUp
    Rows("392:392").Select
    Selection.Delete Shift:=xlUp
    Rows("402:402").Select
    Selection.Delete Shift:=xlUp
    Range("B405").Select
    ActiveWindow.ScrollRow = 392
    ActiveWindow.ScrollRow = 391
    ActiveWindow.ScrollRow = 370
    ActiveWindow.ScrollRow = 261
    ActiveWindow.ScrollRow = 232
    ActiveWindow.ScrollRow = 195
    ActiveWindow.ScrollRow = 114
    ActiveWindow.ScrollRow = 83
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Range("A4").Select
    Columns("A:A").ColumnWidth = 66.25
    Selection.Copy
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A5:A10").Select
    ActiveSheet.Paste
    Range("A11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A12").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A12:A21").Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A23").Select
    ActiveSheet.Paste
    Range("A25").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A26").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A26:A62").Select
    ActiveSheet.Paste
    Range("A28").Select
    Selection.End(xlDown).Select
    Range("A63").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A64").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A64:A152").Select
    ActiveSheet.Paste
    Range("A66").Select
    Selection.End(xlDown).Select
    Range("A153").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A154").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A154:A278").Select
    ActiveSheet.Paste
    Range("A156").Select
    Selection.End(xlDown).Select
    Range("A279").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A280").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A280:A298").Select
    ActiveSheet.Paste
    Range("A281").Select
    Selection.End(xlDown).Select
    Range("A299").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A300").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A300:A304").Select
    ActiveSheet.Paste
    Range("A305").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A306").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A306:A313").Select
    ActiveSheet.Paste
    Range("A314").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A315").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A315:A319").Select
    ActiveSheet.Paste
    Range("A320").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A321").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A321:A326").Select
    ActiveSheet.Paste
    Range("A327").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A328").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A328:A338").Select
    ActiveSheet.Paste
    Range("A339").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A340").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A340:A349").Select
    ActiveSheet.Paste
    Range("A350").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A351").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A351:A361").Select
    ActiveSheet.Paste
    Range("A362").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A363").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A363:A370").Select
    ActiveSheet.Paste
    Range("A371").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A372").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A372:A377").Select
    ActiveSheet.Paste
    Range("A378").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A379").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A379:A384").Select
    ActiveSheet.Paste
    Range("A385").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A386").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A386:A391").Select
    ActiveSheet.Paste
    Range("A392").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A393").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A393:A401").Select
    ActiveSheet.Paste
    Range("A402").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A403:A407").Select
    ActiveSheet.Paste
    Range("B403").Select
    Application.CutCopyMode = False
    Rows("402:402").RowHeight = 22.5
    Range("A401").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A2").Select
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 1
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AG:AG").Select
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AQ:AQ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AV:AV").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    Columns("BA:BA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BF:BF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    Columns("BK:BK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BP:BP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BU:BU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    Columns("BZ:BZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CE:CE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    Columns("CJ:CJ").Select
    Columns("CJ:CJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CO:CO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    Columns("CT:CT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CY:CY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 1
    Range("D1").Select
    Selection.Cut
    Range("D1").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Range("AC2").Select
    Selection.Cut
    Range("AC2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Range("AH2").Select
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C5:C407").Select
    Range("C407").Activate
    ActiveSheet.Paste
    Range("D407").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("C18").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("C4").Select
    Application.CutCopyMode = False
    Range("H4").Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("H407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H407").Select
    Range("H407").Activate
    Selection.End(xlUp).Select
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("H407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H407").Select
    Range("H407").Activate
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("H407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H407").Select
    Range("H407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("M407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M407").Select
    Range("M407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("R407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R5:R407").Select
    Range("R407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("W407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W5:W407").Select
    Range("W407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W3").Select
    Application.CutCopyMode = False
    Range("AB4").Select
    Selection.Copy
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("AB407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB5:AB407").Select
    Range("AB407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("AG4").Select
    Selection.Copy
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AG407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG5:AG407").Select
    Range("AG407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("AL4").Select
    Selection.Copy
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AL407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL5:AL407").Select
    Range("AL407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL3").Select
    Application.CutCopyMode = False
    Range("AQ4").Select
    Selection.Copy
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AQ407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ5:AQ407").Select
    Range("AQ407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AV407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AV5:AV407").Select
    Range("AV407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("BA407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA5:BA407").Select
    Range("BA407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BF407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF5:BF407").Select
    Range("BF407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BK407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK5:BK407").Select
    Range("BK407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BP407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP5:BP407").Select
    Range("BP407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BU407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU5:BU407").Select
    Range("BU407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BZ407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ5:BZ407").Select
    Range("BZ407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("CE407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE5:CE407").Select
    Range("CE407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CJ407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ5:CJ407").Select
    Range("CJ407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CO407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO5:CO407").Select
    Range("CO407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CT407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT5:CT407").Select
    Range("CT407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CY407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY5:CY407").Select
    Range("CY407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("D4").Select
    Selection.End(xlToRight).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF5").Select
    Selection.End(xlDown).Select
    Range("BF408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C408").Select
    ActiveSheet.Paste
    Range("B407").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A4:B407").Select
    Range("B407").Activate
    Selection.Copy
    Range("C407").Select
    Selection.End(xlDown).Select
    Range("A8487:B8487").Select
    Range("B8487").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A408:B8487").Select
    Range("B8487").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Range("G5").Select
    ActiveWindow.ScrollColumn = 1
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "MEDIOS"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A8485").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A8485").Select
    Range("A8485").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("C1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.349986266670736
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.349986266670736
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.349986266670736
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.349986266670736
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.349986266670736
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.349986266670736
        .Weight = xlThin
    End With
    Range("B14").Select
    Columns("B:B").ColumnWidth = 54.38
    Columns("C:C").ColumnWidth = 23.25
    Range("A7").Select
    Selection.End(xlDown).Select
    Range("A8485").Select
    Selection.End(xlUp).Select
    ActiveWorkbook.Save
End Sub



Sub ESTILOSDEVIDA()

'
' Macro1 Macro
'
' Acceso directo: CTRL+y
'
    LIMPIAR
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 42.38
    Range("B5").Select
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    Range("B13").Select
    Selection.Cut
    Range("A14").Select
    ActiveSheet.Paste
    Range("B95").Select
    Selection.Cut
    Range("A96").Select
    ActiveSheet.Paste
    Range("B140").Select
    Selection.Cut
    Range("A141").Select
    ActiveSheet.Paste
    Range("B164").Select
    Selection.Cut
    Range("A165").Select
    ActiveSheet.Paste
    Range("B177").Select
    Selection.Cut
    Range("A178").Select
    ActiveSheet.Paste
    Range("B255").Select
    Selection.Cut
    Range("A256").Select
    ActiveSheet.Paste
    Range("B268").Select
    Selection.Cut
    Range("A269").Select
    ActiveSheet.Paste
    Range("C272").Select
    Selection.End(xlDown).Select
    Range("B290").Select
    Selection.Cut
    Range("A290").Select
    ActiveSheet.Paste
    Range("C290").Select
    Selection.End(xlDown).Select
    Range("C292").Select
    Selection.End(xlDown).Select
    Range("B324").Select
    Selection.Cut
    Range("A325").Select
    ActiveSheet.Paste
    Range("C325").Select
    Selection.End(xlDown).Select
    Range("B331").Select
    Selection.Cut
    Range("A332").Select
    ActiveSheet.Paste
    Range("C332").Select
    Selection.End(xlDown).Select
    ActiveWindow.SmallScroll Down:=-42
    ActiveWindow.ScrollRow = 299
    ActiveWindow.ScrollRow = 298
    ActiveWindow.ScrollRow = 297
    ActiveWindow.ScrollRow = 296
    ActiveWindow.ScrollRow = 294
    ActiveWindow.ScrollRow = 292
    ActiveWindow.ScrollRow = 289
    ActiveWindow.ScrollRow = 284
    ActiveWindow.ScrollRow = 278
    ActiveWindow.ScrollRow = 257
    ActiveWindow.ScrollRow = 225
    ActiveWindow.ScrollRow = 164
    ActiveWindow.ScrollRow = 142
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 97
    ActiveWindow.ScrollRow = 71
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Rows("4:5").Select
    Selection.Delete Shift:=xlUp
    Rows("11:11").Select
    ActiveWindow.SmallScroll Down:=72
    Range("11:11,93:93").Select
    Range("A93").Activate
    ActiveWindow.SmallScroll Down:=54
    Range("11:11,93:93,138:138").Select
    Range("A138").Activate
    ActiveWindow.SmallScroll Down:=21
    Range("11:11,93:93,138:138,162:162").Select
    Range("A162").Activate
    ActiveWindow.SmallScroll Down:=21
    Range("11:11,93:93,138:138,162:162,175:175").Select
    Range("A175").Activate
    ActiveWindow.SmallScroll Down:=75
    Range("11:11,93:93,138:138,162:162,175:175,253:253,266:266").Select
    Range("A266").Activate
    ActiveWindow.SmallScroll Down:=39
    Range("11:11,93:93,138:138,162:162,175:175,253:253,266:266,287:287").Select
    Range("A287").Activate
    ActiveWindow.SmallScroll Down:=-9
    Range("A288").Select
    Selection.Cut Destination:=Range("A289")
    Rows("288:288").Select
    Selection.Delete Shift:=xlUp
    Range("A283").Select
    ActiveWindow.ScrollRow = 273
    ActiveWindow.ScrollRow = 272
    ActiveWindow.ScrollRow = 271
    ActiveWindow.ScrollRow = 270
    ActiveWindow.ScrollRow = 268
    ActiveWindow.ScrollRow = 267
    ActiveWindow.ScrollRow = 265
    ActiveWindow.ScrollRow = 263
    ActiveWindow.ScrollRow = 261
    ActiveWindow.ScrollRow = 259
    ActiveWindow.ScrollRow = 256
    ActiveWindow.ScrollRow = 252
    ActiveWindow.ScrollRow = 246
    ActiveWindow.ScrollRow = 237
    ActiveWindow.ScrollRow = 201
    ActiveWindow.ScrollRow = 173
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 90
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 1
    Rows("11:11").Select
    Selection.Delete Shift:=xlUp
    Range("A17").Select
    ActiveWindow.SmallScroll Down:=69
    Rows("92:92").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=51
    Rows("136:136").Select
    Selection.Delete Shift:=xlUp
    Range("A141").Select
    ActiveWindow.SmallScroll Down:=30
    Rows("159:159").Select
    Selection.Delete Shift:=xlUp
    Rows("171:171").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=87
    Rows("248:248").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=18
    Rows("260:260").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=48
    Rows("314:314").Select
    Selection.Delete Shift:=xlUp
    Range("A311").Select
    ActiveWindow.SmallScroll Down:=3
    Rows("320:320").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=-6
    ActiveWindow.ScrollRow = 300
    ActiveWindow.ScrollRow = 299
    ActiveWindow.ScrollRow = 296
    ActiveWindow.ScrollRow = 294
    ActiveWindow.ScrollRow = 291
    ActiveWindow.ScrollRow = 287
    ActiveWindow.ScrollRow = 279
    ActiveWindow.ScrollRow = 274
    ActiveWindow.ScrollRow = 259
    ActiveWindow.ScrollRow = 233
    ActiveWindow.ScrollRow = 175
    ActiveWindow.ScrollRow = 148
    ActiveWindow.ScrollRow = 108
    ActiveWindow.ScrollRow = 96
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 81
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 55
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 1
    Range("A4").Select
    Selection.Copy
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A5:A10").Select
    ActiveSheet.Paste
    Range("A11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A12").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A12:A91").Select
    ActiveSheet.Paste
    Range("A13").Select
    Selection.End(xlDown).Select
    Range("A92").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A93").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A93:A135").Select
    ActiveSheet.Paste
    Range("A95").Select
    Selection.End(xlDown).Select
    Range("A136").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A137").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A137:A158").Select
    ActiveSheet.Paste
    Range("A140").Select
    Selection.End(xlDown).Select
    Range("A159").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A160").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A160:A170").Select
    ActiveSheet.Paste
    Range("A171").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A172").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A172:A247").Select
    ActiveSheet.Paste
    Range("A175").Select
    Selection.End(xlDown).Select
    Range("A248").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A249").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A249:A259").Select
    ActiveSheet.Paste
    Range("A260").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A261").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A261:A280").Select
    ActiveSheet.Paste
    Range("A265").Select
    Selection.End(xlDown).Select
    Range("A281").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Range("A313").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A282:A313").Select
    Range("A313").Activate
    ActiveSheet.Paste
    Range("A314").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A315:A319").Select
    ActiveSheet.Paste
    Range("A320").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A321").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A321:A1048574").Select
    Selection.End(xlUp).Select
    Range("B321").Select
    Selection.End(xlDown).Select
    Range("A347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A321:A347").Select
    Range("A347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AH2").Select
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Columns("AQ:AQ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Columns("AV:AV").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Columns("BA:BA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Columns("BF:BF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Columns("BK:BK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Columns("BP:BP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Columns("BU:BU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Columns("BZ:BZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CE:CE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Columns("CJ:CJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CO:CO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CT:CT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Columns("CY:CY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("C4").Select
    Selection.Copy
    Application.CutCopyMode = False
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C5:C347").Select
    Range("C347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("H347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H347").Select
    Range("H347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("M347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M347").Select
    Range("M347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("R347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R5:R347").Select
    Range("R347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("W347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W5:W347").Select
    Range("W347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("AB347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB5:AB347").Select
    Range("AB347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AG347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG5:AG347").Select
    Range("AG347").Activate
    ActiveSheet.Paste
    Range("AG346").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AL347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL5:AL347").Select
    Range("AL347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AQ347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ5:AQ347").Select
    Range("AQ347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AV347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AV5:AV347").Select
    Range("AV347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("BA347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA5:BA347").Select
    Range("BA347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BF347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF5:BF347").Select
    Range("BF347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BK347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK5:BK347").Select
    Range("BK347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BP347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP5:BP347").Select
    Range("BP347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BU347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU5:BU347").Select
    Range("BU347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BZ347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ5:BZ347").Select
    Range("BZ347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("CE347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE5:CE347").Select
    Range("CE347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CJ347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ5:CJ347").Select
    Range("CJ347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CO347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO5:CO347").Select
    Range("CO347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CT347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT5:CT347").Select
    Range("CT347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CY347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY5:CY347").Select
    Range("CY347").Activate
    ActiveSheet.Paste
    Range("CY346").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("DB4").Select
    Application.CutCopyMode = False
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("D2:G2").Select
    Selection.Cut Destination:=Range("D3:G3")
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    Range("E4").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    Range("CY1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=-150
    ActiveWindow.ScrollRow = 188
    ActiveWindow.ScrollRow = 186
    ActiveWindow.ScrollRow = 178
    ActiveWindow.ScrollRow = 173
    ActiveWindow.ScrollRow = 128
    ActiveWindow.ScrollRow = 91
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Range("CY2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT2").Select
    Selection.End(xlDown).Select
    Range("CT346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO2").Select
    Selection.End(xlDown).Select
    Range("CO346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ2").Select
    Selection.End(xlDown).Select
    Range("CJ346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE2").Select
    Selection.End(xlDown).Select
    Range("CE346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ2").Select
    Selection.End(xlDown).Select
    Range("BZ346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU2").Select
    Selection.End(xlDown).Select
    Range("BU346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP2").Select
    Selection.End(xlDown).Select
    Range("BP346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK2").Select
    Selection.End(xlDown).Select
    Range("BK346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF2").Select
    Selection.End(xlDown).Select
    Range("BF346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA2").Select
    Selection.End(xlDown).Select
    Range("BA346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV2").Select
    Selection.End(xlDown).Select
    Range("AV346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ2").Select
    Selection.End(xlDown).Select
    Range("AQ346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL2").Select
    Selection.End(xlDown).Select
    Range("AL346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG2").Select
    Selection.End(xlDown).Select
    Range("AG346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB2").Select
    Selection.End(xlDown).Select
    Range("AB346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W2").Select
    Selection.End(xlDown).Select
    Range("W346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R2").Select
    Selection.End(xlDown).Select
    Range("R346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M2").Select
    Selection.End(xlDown).Select
    Range("M346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H2").Select
    Selection.End(xlDown).Select
    Range("H346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C2").Select
    Selection.End(xlDown).Select
    Range("C346").Select
    ActiveSheet.Paste
    Range("C345").Select
    Selection.End(xlUp).Select
    Range("C1").Select
    Selection.End(xlUp).Select
    Range("A94").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("C2").Select
    Selection.End(xlDown).Select
    Range("A7225:B7225").Select
    Range("B7225").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A346:B7225").Select
    Range("B7225").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("C2").Select
    Application.CutCopyMode = False
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 90
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Respuestas"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Range("B1").Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2:G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    Range("A17").Select
    ActiveWindow.SmallScroll Down:=-21
    Range("A1").Select
    Selection.AutoFilter
    Range("A16").Select
    Selection.End(xlDown).Select
    Range("A7225").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
End Sub


Sub EQUIPAMIENTO()

'
' Macro3 Macro
'
' Acceso directo: CTRL+i
'
    LIMPIAR
    Range("H11").Select
    ActiveWindow.SmallScroll Down:=-24
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 37.25
    Range("B5").Select
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=12
    Range("B26").Select
    Selection.Cut
    Range("A27").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=12
    Range("B33").Select
    Selection.Cut
    Range("A34").Select
    ActiveSheet.Paste
    Range("B39").Select
    Selection.Cut
    Range("A40").Select
    ActiveSheet.Paste
    Range("B42").Select
    Selection.Cut
    Range("A43").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("B47").Select
    Selection.Cut
    Range("A48").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=-51
    Rows("4:5").Select
    Selection.Delete Shift:=xlUp
    Range("B6").Select
    ActiveWindow.SmallScroll Down:=18
    Rows("24:24").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=9
    Range("36:36,39:39,44:44").Select
    Range("A44").Activate
    Selection.Delete Shift:=xlUp
    Range("A33").Select
    ActiveWindow.SmallScroll Down:=-45
    Range("A4").Select
    Rows("4:4").RowHeight = 22.5
    Range("A4").Select
    Selection.Copy
    Range("A5").Select
    ActiveWindow.SmallScroll Down:=6
    Range("A5:A23").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=15
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Cut
    Application.CutCopyMode = False
    Selection.Copy
    Range("A25:A30").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=6
    Range("A31").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A32:A35").Select
    ActiveSheet.Paste
    Range("A36").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A37").Select
    ActiveSheet.Paste
    Range("A38").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A39:A41").Select
    ActiveSheet.Paste
    Range("A42").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A43").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlUp).Select
    ActiveWindow.SmallScroll Down:=-27
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 2
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 18.13
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AH2").Select
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    Columns("AQ:AQ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Columns("AV:AV").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    Columns("BA:BA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Columns("BF:BF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    Columns("BK:BK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    Columns("BP:BP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    Columns("BU:BU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Columns("BZ:BZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    Columns("CE:CE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Columns("CJ:CJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Columns("CO:CO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    Columns("CT:CT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Columns("CY:CY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C50").Select
    ActiveWindow.SmallScroll Down:=-18
    Application.CutCopyMode = False
    ActiveWindow.SmallScroll Down:=-15
    Rows("30:30").Select
    Selection.Delete Shift:=xlUp
    Range("A37").Select
    ActiveWindow.SmallScroll Down:=-6
    Rows("24:29").Select
    Selection.RowHeight = 15.75
    Range("A31").Select
    ActiveWindow.SmallScroll Down:=-36
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C5:C42").Select
    Range("C42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G4").Select
    Selection.End(xlDown).Select
    Range("H42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H42").Select
    Range("H42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L4").Select
    Selection.End(xlDown).Select
    Range("M42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M42").Select
    Range("M42").Activate
    ActiveSheet.Paste
    Range("C4").Select
    Application.CutCopyMode = False
    Range("R4").Select
    Selection.Copy
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("R42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R5:R42").Select
    Range("R42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("R4,R6").Select
    Range("R6").Activate
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("W42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W5:W42").Select
    Range("W42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("AB42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB5:AB42").Select
    Range("AB42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AG42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG5:AG42").Select
    Range("AG42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AL42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL5:AL42").Select
    Range("AL42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AQ42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ5:AQ42").Select
    Range("AQ42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AV42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AV5:AV42").Select
    Range("AV42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("BA42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA5:BA42").Select
    Range("BA42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BF42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF5:BF42").Select
    Range("BF42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BF5").Select
    Selection.End(xlDown).Select
    Range("BK42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK5:BK42").Select
    Range("BK42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BP42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP5:BP42").Select
    Range("BP42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BU42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU5:BU42").Select
    Range("BU42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BZ42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ5:BZ42").Select
    Range("BZ42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("CE42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE5:CE42").Select
    Range("CE42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CJ42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ5:CJ42").Select
    Range("CJ42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CO42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO5:CO42").Select
    Range("CO42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CT42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT5:CT42").Select
    Range("CT42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CY42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY5:CY42").Select
    Range("CY42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    ActiveWindow.ScrollColumn = 90
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("D1:G1").Select
    Selection.Cut
    Range("D2").Select
    ActiveSheet.Paste
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("I1").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    Range("CY2:DC2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT2").Select
    Selection.End(xlDown).Select
    Range("CT41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO2").Select
    Selection.End(xlDown).Select
    Range("CO41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ2").Select
    Selection.End(xlDown).Select
    Range("CJ41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE2").Select
    Selection.End(xlDown).Select
    Range("CE41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ2").Select
    Selection.End(xlDown).Select
    Range("BZ41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU2").Select
    Selection.End(xlDown).Select
    Range("BU41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP2").Select
    Selection.End(xlDown).Select
    Range("BP41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK2").Select
    Selection.End(xlDown).Select
    Range("BK41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF2").Select
    Selection.End(xlDown).Select
    Range("BF41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA2").Select
    Selection.End(xlDown).Select
    Range("BA41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV2").Select
    Selection.End(xlDown).Select
    Range("AV41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ2").Select
    Selection.End(xlDown).Select
    Range("AQ41").Select
    ActiveSheet.Paste
    Range("AQ40").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL2").Select
    Selection.End(xlDown).Select
    Range("AL41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG2").Select
    Selection.End(xlDown).Select
    Range("AG41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB2").Select
    Selection.End(xlDown).Select
    Range("AB41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W2").Select
    Selection.End(xlDown).Select
    Range("W41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R2").Select
    Selection.End(xlDown).Select
    Range("R41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M2").Select
    Selection.End(xlDown).Select
    Range("M41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H2").Select
    Selection.End(xlDown).Select
    Range("H41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C2").Select
    Selection.End(xlDown).Select
    Range("C41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("C820").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("C2").Select
    Selection.End(xlDown).Select
    Range("A820:B820").Select
    Range("B820").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A41:B820").Select
    Range("B820").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 90
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variables"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "EQUIPAMIENTO"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A820").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A820").Select
    Range("A820").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("A:A").ColumnWidth = 23.13
    Range("A2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    Range("C1").Select
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2").Select
    Rows("2:2").RowHeight = 17
    Columns("C:C").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Selection.End(xlDown).Select
    Range("A820").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
End Sub

Sub MEDIOSDEPTO()

'
' Macro1 Macro
'
' Acceso directo: CTRL+u
'
    LIMPIAR
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 53.63
    Range("B4").Select
    Selection.Cut
    Range("A5").Select
    ActiveSheet.Paste
    Range("B12").Select
    Selection.Cut
    Range("A13").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=15
    Range("B24").Select
    Selection.Cut
    Range("A25").Select
    ActiveSheet.Paste
    Range("B27").Select
    Selection.Cut
    Range("A28").Select
    ActiveSheet.Paste
    Range("B29").Select
    Selection.Cut
    Range("A30").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=45
    Range("B68").Select
    Selection.Cut
    Range("A69").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=90
    Range("B159").Select
    Selection.Cut
    Range("A160").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=129
    Range("B286").Select
    Selection.Cut
    Range("A287").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=27
    Range("B314").Select
    ActiveWindow.SmallScroll Down:=-6
    Selection.Cut
    Range("A308").Select
    ActiveSheet.Paste
    Selection.Cut Destination:=Range("A315")
    Range("B307").Select
    Selection.Cut
    Range("A308").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=15
    Range("B324").Select
    Selection.Cut
    Range("A325").Select
    ActiveSheet.Paste
    Range("B331").Select
    Selection.Cut
    Range("A332").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("B339").Select
    Selection.Cut
    Range("A340").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("B352").Select
    Selection.Cut
    Range("A353").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=18
    Range("B364").Select
    Selection.Cut
    Range("A365").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("B377").Select
    Selection.Cut
    Range("A378").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("B387").Select
    Selection.Cut
    Range("A388").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=12
    Range("B395").Select
    Selection.Cut
    Range("A396").Select
    ActiveSheet.Paste
    Range("B403").Select
    Selection.Cut
    Range("A404").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=18
    Range("B411").Select
    Selection.Cut
    Range("A412").Select
    ActiveSheet.Paste
    Range("B422").Select
    Selection.Cut
    Range("A423").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=18
    Range("B429").Select
    Selection.Cut
    Range("A430").Select
    ActiveSheet.Paste
    Range("B437").Select
    Selection.Cut
    Range("A438").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=21
    Range("B449").Select
    Selection.Cut
    Range("A450").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=186
    Range("B632").Select
    Selection.Cut
    Range("A633").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=30
    Range("B660").Select
    Selection.Cut
    Range("A661").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=123
    Range("B789").Select
    Selection.Cut
    Range("A790").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=93
    ActiveWindow.ScrollRow = 870
    ActiveWindow.ScrollRow = 868
    ActiveWindow.ScrollRow = 866
    ActiveWindow.ScrollRow = 854
    ActiveWindow.ScrollRow = 846
    ActiveWindow.ScrollRow = 798
    ActiveWindow.ScrollRow = 786
    ActiveWindow.ScrollRow = 719
    ActiveWindow.ScrollRow = 691
    ActiveWindow.ScrollRow = 590
    ActiveWindow.ScrollRow = 558
    ActiveWindow.ScrollRow = 471
    ActiveWindow.ScrollRow = 442
    ActiveWindow.ScrollRow = 380
    ActiveWindow.ScrollRow = 350
    ActiveWindow.ScrollRow = 271
    ActiveWindow.ScrollRow = 255
    ActiveWindow.ScrollRow = 222
    ActiveWindow.ScrollRow = 204
    ActiveWindow.ScrollRow = 181
    ActiveWindow.ScrollRow = 165
    ActiveWindow.ScrollRow = 118
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 79
    ActiveWindow.ScrollRow = 71
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Rows("11:11").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=9
    Rows("22:22").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=9
    Rows("24:24").Select
    Selection.Delete Shift:=xlUp
    Rows("25:25").Select
    Selection.Delete Shift:=xlUp
    Range("A29").Select
    ActiveWindow.SmallScroll Down:=36
    Rows("63:63").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=90
    Rows("153:153").Select
    Selection.Delete Shift:=xlUp
    Range("B154").Select
    ActiveWindow.SmallScroll Down:=126
    Rows("279:279").Select
    Selection.Delete Shift:=xlUp
    Range("A281").Select
    ActiveWindow.SmallScroll Down:=21
    Rows("299:299").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=9
    Rows("315:315").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=-6
    Rows("305:305").Select
    Selection.Delete Shift:=xlUp
    Range("A308").Select
    ActiveWindow.SmallScroll Down:=15
    Rows("320:320").Select
    Selection.Delete Shift:=xlUp
    Rows("327:327").Select
    Selection.Delete Shift:=xlUp
    Range("A324").Select
    ActiveWindow.SmallScroll Down:=21
    Rows("339:339").Select
    Selection.Delete Shift:=xlUp
    Rows("350:350").Select
    Selection.Delete Shift:=xlUp
    Range("A342").Select
    ActiveWindow.SmallScroll Down:=15
    Rows("362:362").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=15
    Rows("371:371").Select
    Selection.Delete Shift:=xlUp
    Range("A373").Select
    ActiveWindow.SmallScroll Down:=6
    Rows("378:378").Select
    Selection.Delete Shift:=xlUp
    Rows("385:385").Select
    Selection.Delete Shift:=xlUp
    Range("A378").Select
    ActiveWindow.SmallScroll Down:=6
    Rows("392:392").Select
    Selection.Delete Shift:=xlUp
    Range("A383").Select
    ActiveWindow.SmallScroll Down:=9
    Rows("402:402").Select
    Selection.Delete Shift:=xlUp
    Range("A396").Select
    ActiveWindow.SmallScroll Down:=15
    Rows("408:408").Select
    Selection.Delete Shift:=xlUp
    Rows("415:415").Select
    Selection.Delete Shift:=xlUp
    Range("A409").Select
    ActiveWindow.SmallScroll Down:=18
    Rows("426:426").Select
    Selection.Delete Shift:=xlUp
    Range("A430").Select
    ActiveWindow.SmallScroll Down:=189
    Rows("608:608").Select
    Selection.Delete Shift:=xlUp
    Range("A614").Select
    ActiveWindow.SmallScroll Down:=24
    Rows("635:635").Select
    Selection.Delete Shift:=xlUp
    Range("A641").Select
    ActiveWindow.SmallScroll Down:=126
    Rows("763:763").Select
    Selection.Delete Shift:=xlUp
    Range("A765").Select
    ActiveWindow.SmallScroll Down:=129
    ActiveWindow.ScrollRow = 884
    ActiveWindow.ScrollRow = 885
    ActiveWindow.ScrollRow = 886
    ActiveWindow.ScrollRow = 888
    ActiveWindow.ScrollRow = 885
    ActiveWindow.ScrollRow = 881
    ActiveWindow.ScrollRow = 876
    ActiveWindow.ScrollRow = 872
    ActiveWindow.ScrollRow = 868
    ActiveWindow.ScrollRow = 865
    ActiveWindow.ScrollRow = 860
    ActiveWindow.ScrollRow = 858
    ActiveWindow.ScrollRow = 856
    ActiveWindow.ScrollRow = 854
    ActiveWindow.ScrollRow = 852
    ActiveWindow.ScrollRow = 849
    ActiveWindow.ScrollRow = 848
    ActiveWindow.ScrollRow = 846
    ActiveWindow.ScrollRow = 845
    ActiveWindow.ScrollRow = 844
    ActiveWindow.ScrollRow = 842
    ActiveWindow.ScrollRow = 841
    ActiveWindow.ScrollRow = 840
    ActiveWindow.ScrollRow = 838
    ActiveWindow.ScrollRow = 837
    ActiveWindow.ScrollRow = 836
    ActiveWindow.ScrollRow = 834
    ActiveWindow.ScrollRow = 833
    ActiveWindow.ScrollRow = 832
    ActiveWindow.ScrollRow = 830
    ActiveWindow.ScrollRow = 829
    ActiveWindow.ScrollRow = 828
    ActiveWindow.ScrollRow = 826
    ActiveWindow.ScrollRow = 825
    ActiveWindow.ScrollRow = 824
    ActiveWindow.ScrollRow = 822
    ActiveWindow.ScrollRow = 821
    ActiveWindow.ScrollRow = 820
    ActiveWindow.ScrollRow = 818
    ActiveWindow.ScrollRow = 817
    ActiveWindow.ScrollRow = 816
    ActiveWindow.ScrollRow = 814
    ActiveWindow.ScrollRow = 813
    ActiveWindow.ScrollRow = 812
    ActiveWindow.ScrollRow = 810
    ActiveWindow.ScrollRow = 809
    ActiveWindow.ScrollRow = 808
    ActiveWindow.ScrollRow = 805
    ActiveWindow.ScrollRow = 804
    ActiveWindow.ScrollRow = 801
    ActiveWindow.ScrollRow = 800
    ActiveWindow.ScrollRow = 798
    ActiveWindow.ScrollRow = 797
    ActiveWindow.ScrollRow = 796
    ActiveWindow.ScrollRow = 794
    ActiveWindow.ScrollRow = 793
    ActiveWindow.ScrollRow = 790
    ActiveWindow.ScrollRow = 788
    ActiveWindow.ScrollRow = 784
    ActiveWindow.ScrollRow = 780
    ActiveWindow.ScrollRow = 773
    ActiveWindow.ScrollRow = 769
    ActiveWindow.ScrollRow = 761
    ActiveWindow.ScrollRow = 760
    ActiveWindow.ScrollRow = 758
    ActiveWindow.ScrollRow = 757
    ActiveWindow.ScrollRow = 756
    ActiveWindow.ScrollRow = 753
    ActiveWindow.ScrollRow = 750
    ActiveWindow.ScrollRow = 749
    ActiveWindow.ScrollRow = 746
    ActiveWindow.ScrollRow = 742
    ActiveWindow.ScrollRow = 738
    ActiveWindow.ScrollRow = 734
    ActiveWindow.ScrollRow = 733
    ActiveWindow.ScrollRow = 732
    ActiveWindow.ScrollRow = 730
    ActiveWindow.ScrollRow = 729
    ActiveWindow.ScrollRow = 728
    ActiveWindow.ScrollRow = 726
    ActiveWindow.ScrollRow = 729
    ActiveWindow.ScrollRow = 732
    ActiveWindow.ScrollRow = 734
    ActiveWindow.ScrollRow = 738
    ActiveWindow.ScrollRow = 741
    ActiveWindow.ScrollRow = 742
    ActiveWindow.ScrollRow = 744
    ActiveWindow.ScrollRow = 742
    ActiveWindow.ScrollRow = 741
    ActiveWindow.ScrollRow = 740
    ActiveWindow.ScrollRow = 738
    ActiveWindow.ScrollRow = 737
    ActiveWindow.ScrollRow = 734
    ActiveWindow.ScrollRow = 733
    ActiveWindow.ScrollRow = 729
    ActiveWindow.ScrollRow = 725
    ActiveWindow.ScrollRow = 722
    ActiveWindow.ScrollRow = 718
    ActiveWindow.ScrollRow = 714
    ActiveWindow.ScrollRow = 712
    ActiveWindow.ScrollRow = 709
    ActiveWindow.ScrollRow = 708
    ActiveWindow.ScrollRow = 704
    ActiveWindow.ScrollRow = 701
    ActiveWindow.ScrollRow = 697
    ActiveWindow.ScrollRow = 693
    ActiveWindow.ScrollRow = 690
    ActiveWindow.ScrollRow = 682
    ActiveWindow.ScrollRow = 677
    ActiveWindow.ScrollRow = 666
    ActiveWindow.ScrollRow = 662
    ActiveWindow.ScrollRow = 646
    ActiveWindow.ScrollRow = 633
    ActiveWindow.ScrollRow = 605
    ActiveWindow.ScrollRow = 592
    ActiveWindow.ScrollRow = 555
    ActiveWindow.ScrollRow = 548
    ActiveWindow.ScrollRow = 516
    ActiveWindow.ScrollRow = 499
    ActiveWindow.ScrollRow = 452
    ActiveWindow.ScrollRow = 444
    ActiveWindow.ScrollRow = 412
    ActiveWindow.ScrollRow = 397
    ActiveWindow.ScrollRow = 372
    ActiveWindow.ScrollRow = 363
    ActiveWindow.ScrollRow = 352
    ActiveWindow.ScrollRow = 348
    ActiveWindow.ScrollRow = 333
    ActiveWindow.ScrollRow = 324
    ActiveWindow.ScrollRow = 313
    ActiveWindow.ScrollRow = 309
    ActiveWindow.ScrollRow = 300
    ActiveWindow.ScrollRow = 298
    ActiveWindow.ScrollRow = 284
    ActiveWindow.ScrollRow = 280
    ActiveWindow.ScrollRow = 266
    ActiveWindow.ScrollRow = 256
    ActiveWindow.ScrollRow = 242
    ActiveWindow.ScrollRow = 238
    ActiveWindow.ScrollRow = 230
    ActiveWindow.ScrollRow = 227
    ActiveWindow.ScrollRow = 216
    ActiveWindow.ScrollRow = 212
    ActiveWindow.ScrollRow = 199
    ActiveWindow.ScrollRow = 195
    ActiveWindow.ScrollRow = 186
    ActiveWindow.ScrollRow = 184
    ActiveWindow.ScrollRow = 176
    ActiveWindow.ScrollRow = 174
    ActiveWindow.ScrollRow = 170
    ActiveWindow.ScrollRow = 163
    ActiveWindow.ScrollRow = 159
    ActiveWindow.ScrollRow = 156
    ActiveWindow.ScrollRow = 154
    ActiveWindow.ScrollRow = 148
    ActiveWindow.ScrollRow = 140
    ActiveWindow.ScrollRow = 131
    ActiveWindow.ScrollRow = 122
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 102
    ActiveWindow.ScrollRow = 91
    ActiveWindow.ScrollRow = 88
    ActiveWindow.ScrollRow = 80
    ActiveWindow.ScrollRow = 76
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 70
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 64
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 55
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Range("D3").Select
    Columns("B:B").ColumnWidth = 27.13
    Range("A4").Select
    Selection.Copy
    Range("A5:A10").Select
    ActiveSheet.Paste
    Range("A11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A12:A21").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A23").Select
    ActiveSheet.Paste
    Range("A25").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A26").Select
    ActiveWindow.SmallScroll Down:=42
    Range("A26:A62").Select
    ActiveSheet.Paste
    Range("A31").Select
    ActiveWindow.SmallScroll Down:=39
    Range("A63").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A64").Select
    ActiveWindow.SmallScroll Down:=78
    Range("A64:A152").Select
    ActiveSheet.Paste
    Range("A69").Select
    Application.CutCopyMode = False
    Rows("63:63").RowHeight = 17
    ActiveWindow.SmallScroll Down:=93
    Range("A153").Select
    Selection.Copy
    Range("A154").Select
    ActiveWindow.SmallScroll Down:=129
    Range("A154:A278").Select
    ActiveSheet.Paste
    Range("A158").Select
    ActiveWindow.SmallScroll Down:=135
    Range("A279").Select
    Application.CutCopyMode = False
    Selection.Copy
    Rows("279:279").RowHeight = 14.25
    Range("A280").Select
    ActiveWindow.SmallScroll Down:=9
    Range("A280:A298").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=27
    Range("A299").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A300:A304").Select
    ActiveSheet.Paste
    Range("A305").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A306:A313").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("A314").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A315:A319").Select
    ActiveSheet.Paste
    Range("A320").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A321:A326").Select
    ActiveSheet.Paste
    Range("A323").Select
    ActiveWindow.SmallScroll Down:=21
    Range("A327").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A328:A338").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("A339").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A340:A349").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("A350").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A351:A361").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=15
    Range("A362").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A363:A370").Select
    ActiveSheet.Paste
    Range("A371").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A372:A377").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("A378").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A379:A384").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("A385").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A386:A391").Select
    ActiveSheet.Paste
    Range("A392").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A393:A401").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=18
    Range("A402").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A403:A407").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("A408").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A409:A414").Select
    ActiveSheet.Paste
    Range("A415").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A416:A425").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=24
    Range("A426").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A427:A607").Select
    ActiveSheet.Paste
    Range("A429").Select
    Selection.End(xlDown).Select
    Range("A608").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.SmallScroll Down:=12
    Range("A609").Select
    ActiveWindow.SmallScroll Down:=33
    Range("A609:A634").Select
    ActiveSheet.Paste
    Range("A616").Select
    ActiveWindow.SmallScroll Down:=27
    Range("A635").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A636").Select
    ActiveWindow.SmallScroll Down:=129
    Range("A636:A762").Select
    ActiveSheet.Paste
    Range("A642").Select
    ActiveWindow.SmallScroll Down:=138
    Range("A763").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A764").Select
    ActiveWindow.SmallScroll Down:=117
    Range("A764:A890").Select
    ActiveSheet.Paste
    Range("A770").Select
    ActiveWindow.ScrollRow = 747
    ActiveWindow.ScrollRow = 748
    ActiveWindow.ScrollRow = 750
    ActiveWindow.ScrollRow = 752
    ActiveWindow.ScrollRow = 756
    ActiveWindow.ScrollRow = 760
    ActiveWindow.ScrollRow = 767
    ActiveWindow.ScrollRow = 774
    ActiveWindow.ScrollRow = 782
    ActiveWindow.ScrollRow = 788
    ActiveWindow.ScrollRow = 795
    ActiveWindow.ScrollRow = 798
    ActiveWindow.ScrollRow = 804
    ActiveWindow.ScrollRow = 808
    ActiveWindow.ScrollRow = 810
    ActiveWindow.ScrollRow = 811
    ActiveWindow.ScrollRow = 814
    ActiveWindow.ScrollRow = 815
    ActiveWindow.ScrollRow = 818
    ActiveWindow.ScrollRow = 820
    ActiveWindow.ScrollRow = 823
    ActiveWindow.ScrollRow = 824
    ActiveWindow.ScrollRow = 826
    ActiveWindow.ScrollRow = 828
    ActiveWindow.ScrollRow = 830
    ActiveWindow.ScrollRow = 832
    ActiveWindow.ScrollRow = 834
    ActiveWindow.ScrollRow = 836
    ActiveWindow.ScrollRow = 838
    ActiveWindow.ScrollRow = 839
    ActiveWindow.ScrollRow = 840
    ActiveWindow.ScrollRow = 842
    ActiveWindow.ScrollRow = 844
    ActiveWindow.ScrollRow = 847
    ActiveWindow.ScrollRow = 851
    ActiveWindow.ScrollRow = 852
    ActiveWindow.ScrollRow = 856
    ActiveWindow.ScrollRow = 858
    ActiveWindow.ScrollRow = 860
    ActiveWindow.ScrollRow = 863
    ActiveWindow.ScrollRow = 864
    ActiveWindow.ScrollRow = 867
    ActiveWindow.ScrollRow = 868
    ActiveWindow.ScrollRow = 871
    ActiveWindow.ScrollRow = 874
    ActiveWindow.ScrollRow = 875
    ActiveWindow.ScrollRow = 876
    ActiveWindow.ScrollRow = 878
    ActiveWindow.ScrollRow = 880
    ActiveWindow.ScrollRow = 882
    ActiveWindow.ScrollRow = 883
    ActiveWindow.ScrollRow = 886
    ActiveWindow.ScrollRow = 880
    ActiveWindow.ScrollRow = 876
    ActiveWindow.ScrollRow = 872
    ActiveWindow.ScrollRow = 866
    ActiveWindow.ScrollRow = 862
    ActiveWindow.ScrollRow = 852
    ActiveWindow.ScrollRow = 848
    ActiveWindow.ScrollRow = 843
    ActiveWindow.ScrollRow = 842
    ActiveWindow.ScrollRow = 840
    ActiveWindow.ScrollRow = 839
    ActiveWindow.ScrollRow = 838
    ActiveWindow.ScrollRow = 836
    ActiveWindow.ScrollRow = 835
    ActiveWindow.ScrollRow = 834
    ActiveWindow.ScrollRow = 832
    ActiveWindow.ScrollRow = 830
    ActiveWindow.ScrollRow = 826
    ActiveWindow.ScrollRow = 820
    ActiveWindow.ScrollRow = 815
    ActiveWindow.ScrollRow = 803
    ActiveWindow.ScrollRow = 796
    ActiveWindow.ScrollRow = 772
    ActiveWindow.ScrollRow = 752
    ActiveWindow.ScrollRow = 728
    ActiveWindow.ScrollRow = 719
    ActiveWindow.ScrollRow = 706
    ActiveWindow.ScrollRow = 702
    ActiveWindow.ScrollRow = 693
    ActiveWindow.ScrollRow = 683
    ActiveWindow.ScrollRow = 667
    ActiveWindow.ScrollRow = 663
    ActiveWindow.ScrollRow = 647
    ActiveWindow.ScrollRow = 641
    ActiveWindow.ScrollRow = 617
    ActiveWindow.ScrollRow = 602
    ActiveWindow.ScrollRow = 559
    ActiveWindow.ScrollRow = 545
    ActiveWindow.ScrollRow = 522
    ActiveWindow.ScrollRow = 513
    ActiveWindow.ScrollRow = 502
    ActiveWindow.ScrollRow = 496
    ActiveWindow.ScrollRow = 480
    ActiveWindow.ScrollRow = 474
    ActiveWindow.ScrollRow = 456
    ActiveWindow.ScrollRow = 444
    ActiveWindow.ScrollRow = 426
    ActiveWindow.ScrollRow = 416
    ActiveWindow.ScrollRow = 398
    ActiveWindow.ScrollRow = 394
    ActiveWindow.ScrollRow = 378
    ActiveWindow.ScrollRow = 373
    ActiveWindow.ScrollRow = 365
    ActiveWindow.ScrollRow = 362
    ActiveWindow.ScrollRow = 358
    ActiveWindow.ScrollRow = 356
    ActiveWindow.ScrollRow = 355
    ActiveWindow.ScrollRow = 352
    ActiveWindow.ScrollRow = 351
    ActiveWindow.ScrollRow = 349
    ActiveWindow.ScrollRow = 348
    ActiveWindow.ScrollRow = 345
    ActiveWindow.ScrollRow = 344
    ActiveWindow.ScrollRow = 343
    ActiveWindow.ScrollRow = 341
    ActiveWindow.ScrollRow = 340
    ActiveWindow.ScrollRow = 337
    ActiveWindow.ScrollRow = 336
    ActiveWindow.ScrollRow = 335
    ActiveWindow.ScrollRow = 333
    ActiveWindow.ScrollRow = 331
    ActiveWindow.ScrollRow = 328
    ActiveWindow.ScrollRow = 327
    ActiveWindow.ScrollRow = 323
    ActiveWindow.ScrollRow = 321
    ActiveWindow.ScrollRow = 319
    ActiveWindow.ScrollRow = 317
    ActiveWindow.ScrollRow = 315
    ActiveWindow.ScrollRow = 313
    ActiveWindow.ScrollRow = 309
    ActiveWindow.ScrollRow = 308
    ActiveWindow.ScrollRow = 307
    ActiveWindow.ScrollRow = 304
    ActiveWindow.ScrollRow = 301
    ActiveWindow.ScrollRow = 300
    ActiveWindow.ScrollRow = 299
    ActiveWindow.ScrollRow = 297
    ActiveWindow.ScrollRow = 296
    ActiveWindow.ScrollRow = 293
    ActiveWindow.ScrollRow = 292
    ActiveWindow.ScrollRow = 291
    ActiveWindow.ScrollRow = 289
    ActiveWindow.ScrollRow = 288
    ActiveWindow.ScrollRow = 287
    ActiveWindow.ScrollRow = 285
    ActiveWindow.ScrollRow = 284
    ActiveWindow.ScrollRow = 283
    ActiveWindow.ScrollRow = 280
    ActiveWindow.ScrollRow = 277
    ActiveWindow.ScrollRow = 273
    ActiveWindow.ScrollRow = 271
    ActiveWindow.ScrollRow = 268
    ActiveWindow.ScrollRow = 263
    ActiveWindow.ScrollRow = 260
    ActiveWindow.ScrollRow = 252
    ActiveWindow.ScrollRow = 245
    ActiveWindow.ScrollRow = 241
    ActiveWindow.ScrollRow = 233
    ActiveWindow.ScrollRow = 231
    ActiveWindow.ScrollRow = 223
    ActiveWindow.ScrollRow = 219
    ActiveWindow.ScrollRow = 211
    ActiveWindow.ScrollRow = 209
    ActiveWindow.ScrollRow = 205
    ActiveWindow.ScrollRow = 201
    ActiveWindow.ScrollRow = 197
    ActiveWindow.ScrollRow = 196
    ActiveWindow.ScrollRow = 193
    ActiveWindow.ScrollRow = 191
    ActiveWindow.ScrollRow = 189
    ActiveWindow.ScrollRow = 188
    ActiveWindow.ScrollRow = 185
    ActiveWindow.ScrollRow = 184
    ActiveWindow.ScrollRow = 181
    ActiveWindow.ScrollRow = 180
    ActiveWindow.ScrollRow = 178
    ActiveWindow.ScrollRow = 176
    ActiveWindow.ScrollRow = 175
    ActiveWindow.ScrollRow = 172
    ActiveWindow.ScrollRow = 171
    ActiveWindow.ScrollRow = 168
    ActiveWindow.ScrollRow = 164
    ActiveWindow.ScrollRow = 162
    ActiveWindow.ScrollRow = 159
    ActiveWindow.ScrollRow = 158
    ActiveWindow.ScrollRow = 155
    ActiveWindow.ScrollRow = 154
    ActiveWindow.ScrollRow = 152
    ActiveWindow.ScrollRow = 151
    ActiveWindow.ScrollRow = 147
    ActiveWindow.ScrollRow = 144
    ActiveWindow.ScrollRow = 142
    ActiveWindow.ScrollRow = 139
    ActiveWindow.ScrollRow = 136
    ActiveWindow.ScrollRow = 135
    ActiveWindow.ScrollRow = 134
    ActiveWindow.ScrollRow = 131
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 126
    ActiveWindow.ScrollRow = 124
    ActiveWindow.ScrollRow = 122
    ActiveWindow.ScrollRow = 119
    ActiveWindow.ScrollRow = 118
    ActiveWindow.ScrollRow = 116
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 111
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 103
    ActiveWindow.ScrollRow = 98
    ActiveWindow.ScrollRow = 95
    ActiveWindow.ScrollRow = 90
    ActiveWindow.ScrollRow = 88
    ActiveWindow.ScrollRow = 86
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 80
    ActiveWindow.ScrollRow = 79
    ActiveWindow.ScrollRow = 75
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 70
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 50
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Range("A9").Select
    Application.CutCopyMode = False
    ActiveWindow.SmallScroll Down:=885
    ActiveWindow.ScrollRow = 884
    ActiveWindow.ScrollRow = 883
    ActiveWindow.ScrollRow = 880
    ActiveWindow.ScrollRow = 876
    ActiveWindow.ScrollRow = 874
    ActiveWindow.ScrollRow = 868
    ActiveWindow.ScrollRow = 863
    ActiveWindow.ScrollRow = 856
    ActiveWindow.ScrollRow = 844
    ActiveWindow.ScrollRow = 835
    ActiveWindow.ScrollRow = 784
    ActiveWindow.ScrollRow = 752
    ActiveWindow.ScrollRow = 598
    ActiveWindow.ScrollRow = 563
    ActiveWindow.ScrollRow = 494
    ActiveWindow.ScrollRow = 465
    ActiveWindow.ScrollRow = 397
    ActiveWindow.ScrollRow = 369
    ActiveWindow.ScrollRow = 307
    ActiveWindow.ScrollRow = 275
    ActiveWindow.ScrollRow = 203
    ActiveWindow.ScrollRow = 179
    ActiveWindow.ScrollRow = 140
    ActiveWindow.ScrollRow = 120
    ActiveWindow.ScrollRow = 100
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 64
    ActiveWindow.ScrollRow = 55
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 1
    Columns("A:A").ColumnWidth = 41.5
    Range("B12").Select
    ActiveWindow.SmallScroll Down:=-27
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    Selection.Cut Destination:=Range("C4")
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("I2").Select
    Selection.Cut Destination:=Range("H4")
    Range("H4").Select
    Columns("H:H").ColumnWidth = 17.63
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Columns("M:M").ColumnWidth = 18.38
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("S:S").ColumnWidth = 25.38
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AH2").Select
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    Columns("AQ:AQ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Columns("AV:AV").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    Columns("BA:BA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Columns("BF:BF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    Columns("BK:BK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Columns("BP:BP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    Columns("BU:BU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 71
    Columns("BZ:BZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Columns("CE:CE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Columns("CJ:CJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    Columns("CO:CO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Columns("CT:CT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    Columns("CY:CY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 90
    ActiveWindow.ScrollColumn = 91
    ActiveWindow.ScrollColumn = 92
    ActiveWindow.ScrollColumn = 91
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 1
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("B6").Select
    Rows("3:3").RowHeight = 18.25
    Range("C3").Select
    Selection.Copy
    Range("A3").Select
    Selection.End(xlDown).Select
    Range("C889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C4:C889").Select
    Range("C889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C3").Select
    Selection.End(xlDown).Select
    Range("H889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H4:H889").Select
    Range("H889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H3").Select
    Selection.End(xlDown).Select
    Range("M889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M4:M889").Select
    Range("M889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M3").Select
    Selection.End(xlDown).Select
    Range("R889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R4:R889").Select
    Range("R889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R3").Select
    Selection.End(xlDown).Select
    Range("W889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W4:W889").Select
    Range("W889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W3").Select
    Selection.End(xlDown).Select
    Range("AB889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB4:AB889").Select
    Range("AB889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB3").Select
    Selection.End(xlDown).Select
    Range("AG889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG4:AG889").Select
    Range("AG889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AG3").Select
    Selection.End(xlDown).Select
    Range("AL889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL4:AL889").Select
    Range("AL889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AL3").Select
    Selection.End(xlDown).Select
    Range("AQ889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ4:AQ889").Select
    Range("AQ889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AQ3").Select
    Selection.End(xlDown).Select
    Range("AV889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AV4:AV889").Select
    Range("AV889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AV3").Select
    Selection.End(xlDown).Select
    Range("BA889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA4:BA889").Select
    Range("BA889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BA3").Select
    Selection.End(xlDown).Select
    Range("BF889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF4:BF889").Select
    Range("BF889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BF3").Select
    Selection.End(xlDown).Select
    Range("BK889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK4:BK889").Select
    Range("BK889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BK3").Select
    Selection.End(xlDown).Select
    Range("BP889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP4:BP889").Select
    Range("BP889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BP3").Select
    Selection.End(xlDown).Select
    Range("BU889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU4:BU889").Select
    Range("BU889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BU3").Select
    Selection.End(xlDown).Select
    Range("BZ889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ4:BZ889").Select
    Range("BZ889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BZ3").Select
    Selection.End(xlDown).Select
    Range("CE889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE4:CE889").Select
    Range("CE889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CE3").Select
    Selection.End(xlDown).Select
    Range("CJ889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ4:CJ889").Select
    Range("CJ889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CJ3").Select
    Selection.End(xlDown).Select
    Range("CO889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO4:CO889").Select
    Range("CO889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CO3").Select
    Selection.End(xlDown).Select
    Range("CT889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT4:CT889").Select
    Range("CT889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT3").Select
    Selection.End(xlDown).Select
    Range("CY889").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY4:CY889").Select
    Range("CY889").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    ActiveWindow.ScrollColumn = 92
    ActiveWindow.ScrollColumn = 90
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("H2").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    ActiveWindow.ScrollColumn = 91
    ActiveWindow.ScrollColumn = 92
    Range("CY3:DC3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT3").Select
    Selection.End(xlDown).Select
    Range("CT890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO3").Select
    Selection.End(xlDown).Select
    Range("CO890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ3").Select
    Selection.End(xlDown).Select
    Range("CJ890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE3").Select
    Selection.End(xlDown).Select
    Range("CE890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ3").Select
    Selection.End(xlDown).Select
    Range("BZ890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU3").Select
    Selection.End(xlDown).Select
    Range("BU890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP3").Select
    Selection.End(xlDown).Select
    Range("BP890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK3").Select
    Selection.End(xlDown).Select
    Range("BK890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF3").Select
    Selection.End(xlDown).Select
    Range("BF890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA3").Select
    Selection.End(xlDown).Select
    Range("BA890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV3").Select
    Selection.End(xlDown).Select
    Range("AV890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ3").Select
    Selection.End(xlDown).Select
    Range("AQ890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL3").Select
    Selection.End(xlDown).Select
    Range("AL890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG3").Select
    Selection.End(xlDown).Select
    Range("AG890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB3").Select
    Selection.End(xlDown).Select
    Range("AB890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W3").Select
    Selection.End(xlDown).Select
    Range("W890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R3").Select
    Selection.End(xlDown).Select
    Range("R890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M3").Select
    Selection.End(xlDown).Select
    Range("M890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("N3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("M3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H3").Select
    Selection.End(xlDown).Select
    Range("H890").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C3").Select
    Selection.End(xlDown).Select
    Range("C890").Select
    ActiveSheet.Paste
    Range("C889").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    ActiveWindow.ScrollColumn = 91
    ActiveWindow.ScrollColumn = 92
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Rows("2:2").Select
    Range("B2").Activate
    Selection.Delete Shift:=xlUp
    Range("A2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("C2").Select
    Selection.End(xlDown).Select
    Range("A18628:B18628").Select
    Range("B18628").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A889:B18628").Select
    Range("B18628").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("B1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$H$18628").AutoFilter Field:=2, Criteria1:= _
        "Audiencia general de medios alguna vez"
    ActiveSheet.Range("$A$1:$H$18628").AutoFilter Field:=2
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "MEDIOS"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A18628").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A18628").Select
    Range("A18628").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    Range("C1").Select
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A5").Select
    Selection.AutoFilter
    Range("B3").Select
End Sub



Sub PERFIL()

'
' Macro1 Macro
'
' Acceso directo: CTRL+�
'
    LIMPIAR
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B5").Select
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    Range("B8").Select
    Selection.Cut
    Range("A9").Select
    ActiveSheet.Paste
    Range("B11").Select
    Selection.Cut
    Range("A12").Select
    ActiveSheet.Paste
    Range("B17").Select
    Selection.Cut
    Range("A18").Select
    ActiveSheet.Paste
    Range("B29").Select
    Selection.Cut
    Range("A30").Select
    ActiveSheet.Paste
    Range("B37").Select
    Selection.Cut
    Range("A38").Select
    ActiveSheet.Paste
    Range("B43").Select
    Selection.Cut
    Range("A44").Select
    ActiveSheet.Paste
    Range("B47").Select
    Selection.Cut
    Range("A48").Select
    ActiveSheet.Paste
    Range("B50").Select
    Selection.Cut
    Range("A51").Select
    ActiveSheet.Paste
    Range("B56").Select
    Selection.Cut
    Range("A57").Select
    ActiveSheet.Paste
    Range("B59").Select
    Selection.Cut
    Range("A60").Select
    ActiveSheet.Paste
    Range("B68").Select
    Selection.Cut
    Range("A69").Select
    ActiveSheet.Paste
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Rows("6:6").Select
    Selection.Delete Shift:=xlUp
    Rows("8:8").Select
    Selection.Delete Shift:=xlUp
    Rows("13:13").Select
    Selection.Delete Shift:=xlUp
    Rows("24:24").Select
    Selection.Delete Shift:=xlUp
    Rows("31:31").Select
    Selection.Delete Shift:=xlUp
    Rows("36:36").Select
    Selection.Delete Shift:=xlUp
    Rows("39:39").Select
    Selection.Delete Shift:=xlUp
    Range("41:41,47:47,50:50").Select
    Range("A50").Activate
    Selection.Delete Shift:=xlUp
    Rows("56:56").Select
    Selection.Delete Shift:=xlUp
    Range("B68").Select
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 1
    Range("A4").Select
    Selection.Copy
    Range("A5").Select
    ActiveSheet.Paste
    Range("A6").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A7").Select
    ActiveSheet.Paste
    Range("A8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A9:A12").Select
    ActiveSheet.Paste
    Range("A13").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A14:A23").Select
    ActiveSheet.Paste
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A25:A30").Select
    ActiveSheet.Paste
    Range("A31").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A32:A35").Select
    ActiveSheet.Paste
    Range("A36").Select
    Columns("A:A").ColumnWidth = 35.63
    Application.CutCopyMode = False
    Selection.Copy
    Range("A37").Select
    ActiveSheet.Paste
    Range("A38").Select
    ActiveSheet.Paste
    Range("A39").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A40").Select
    ActiveSheet.Paste
    Range("A41").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A42:A45").Select
    ActiveSheet.Paste
    Range("A46").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A47").Select
    ActiveSheet.Paste
    Range("A48").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A49:A55,A41").Select
    Range("A41").Activate
    ActiveSheet.Paste
    Range("A56").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A57:A64").Select
    ActiveSheet.Paste
    Range("A56").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Rows("4:64").Select
    Application.CutCopyMode = False
    Rows("4:64").EntireRow.AutoFit
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    Columns("Q:Q").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("V:V").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AA:AA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    Columns("AF:AF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AK:AK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    Columns("AP:AP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AU:AU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    Columns("AZ:AZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BE:BE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    Columns("BJ:BJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BO:BO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    Columns("BT:BT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BY:BY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    Columns("CD:CD").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CI:CI").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    Columns("CN:CN").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CW:CW").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 1
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C5:C64").Select
    Range("C64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Columns("H:H").ColumnWidth = 14.13
    Selection.Copy
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("H64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H64").Select
    Range("H64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Columns("M:M").ColumnWidth = 13.63
    Selection.Copy
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("M64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M64").Select
    Range("M64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("S2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Columns("R:R").ColumnWidth = 23.88
    Columns("R:R").ColumnWidth = 37.5
    Columns("R:R").ColumnWidth = 39.88
    Selection.Copy
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("R64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R5:R64").Select
    Range("R64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("X2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    ActiveSheet.Paste
    Columns("W:W").ColumnWidth = 21.75
    Application.CutCopyMode = False
    Selection.Copy
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("W64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W5:W64").Select
    Range("W64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AJ2").Select
    Columns("AB:AB").ColumnWidth = 25
    Range("AC2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB4").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("AB64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB5:AB64").Select
    Range("AB64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ1").Select
    Columns("AG:AG").ColumnWidth = 26
    Range("AH2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AG64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG5:AG64").Select
    Range("AG64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AX2").Select
    Columns("AL:AL").ColumnWidth = 26.75
    Range("AM2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AL4").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AL64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL5:AL64").Select
    Range("AL64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BC1").Select
    Columns("AQ:AQ").ColumnWidth = 26
    Range("AR2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AQ64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ5:AQ64").Select
    Range("AQ64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BJ1").Select
    Columns("AV:AV").ColumnWidth = 20.5
    Range("AW2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("AV5:AV64").Select
    ActiveSheet.Paste
    Range("BB3").Select
    Columns("BA:BA").ColumnWidth = 23.75
    Range("BB2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BB2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("BA64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA5:BA64").Select
    Range("BA64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Columns("BF:BF").ColumnWidth = 21.75
    Range("BG2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BF64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF5:BF64").Select
    Range("BF64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BV4").Select
    Columns("BK:BK").ColumnWidth = 27.75
    Range("BL2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BK64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK5:BK64").Select
    Range("BK64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ3").Select
    Columns("BP:BP").ColumnWidth = 28.5
    Range("BQ2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BP64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP5:BP64").Select
    Range("BP64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CG1").Select
    Columns("BU:BU").ColumnWidth = 26.38
    Range("BV2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BU64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU5:BU64").Select
    Range("BU64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CL1").Select
    Columns("BZ:BZ").ColumnWidth = 29
    Range("CA2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BZ64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ5:BZ64").Select
    Range("BZ64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CR4").Select
    Columns("CE:CE").ColumnWidth = 27
    Range("CF2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BZ5").Select
    Selection.End(xlDown).Select
    Range("CE64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE5:CE64").Select
    Range("CE64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CN3").Select
    Columns("CJ:CJ").ColumnWidth = 17.38
    Range("CK2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CJ64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ5:CJ64").Select
    Range("CJ64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CW4").Select
    Columns("CO:CO").ColumnWidth = 23.25
    Range("CP2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CO64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO5:CO64").Select
    Range("CO64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Columns("CT:CT").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CU5").Select
    Columns("CT:CT").ColumnWidth = 16.75
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CT64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT5:CT64").Select
    Range("CT64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CZ2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CY4").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CY64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY5:CY64").Select
    Range("CY64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CZ5").Select
    Application.CutCopyMode = False
    Range("CM4").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("BZ4").Select
    Selection.End(xlToRight).Select
    Range("CY4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU65").Select
    ActiveSheet.Paste
    Range("BU64").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP65").Select
    ActiveSheet.Paste
    Range("BP64").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Range("H4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C65").Select
    ActiveSheet.Paste
    Range("C64").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("B64").Select
    Columns("B:B").ColumnWidth = 36.25
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A4:B64").Select
    Range("B64").Activate
    Selection.Copy
    Range("C64").Select
    Selection.End(xlDown).Select
    Range("A1284:B1284").Select
    Range("B1284").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A65:B1284").Select
    Range("B1284").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Rows("1:1").Select
    Range("B1").Activate
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Range("B2").Activate
    Selection.Delete Shift:=xlUp
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "PERFIL"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A1282").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A1282").Select
    Range("A1282").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("A2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    Range("C1").Select
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A5").Select
    Selection.End(xlDown).Select
    Range("A1282").Select
    Selection.End(xlUp).Select
End Sub


Sub MEDIOS()

'
' Macro1 Macro
'
' Acceso directo: CTRL+k

    LIMPIAR
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 26.5
    Range("B5").Select
    Selection.Cut Destination:=Range("A6")
    Range("B15").Select
    Selection.Cut Destination:=Range("A16")
    Range("A16").Select
    ActiveWindow.SmallScroll Down:=12
    Range("B25").Select
    Selection.Cut Destination:=Range("A26")
    Range("B28").Select
    Selection.Cut Destination:=Range("A29")
    Range("A29").Select
    ActiveWindow.SmallScroll Down:=12
    Range("B31").Select
    Selection.Cut Destination:=Range("A32")
    Range("A32").Select
    ActiveWindow.SmallScroll Down:=63
    Range("B89").Select
    Selection.Cut Destination:=Range("A90")
    Range("A90").Select
    ActiveWindow.SmallScroll Down:=174
    Range("B269").Select
    Selection.Cut Destination:=Range("A270")
    Range("A270").Select
    ActiveWindow.SmallScroll Down:=183
    Range("B450").Select
    Selection.Cut Destination:=Range("A451")
    Range("A451").Select
    ActiveWindow.SmallScroll Down:=27
    Range("B478").Select
    Selection.Cut Destination:=Range("A479")
    Range("A479").Select
    ActiveWindow.SmallScroll Down:=123
    Range("B605").Select
    Selection.Cut Destination:=Range("A607")
    Range("A607").Select
    Selection.Cut Destination:=Range("A606")
    Range("A606").Select
    ActiveWindow.SmallScroll Down:=126
    Range("B732").Select
    Selection.Cut Destination:=Range("A733")
    Range("A733").Select
    ActiveWindow.SmallScroll Down:=129
    Range("B859").Select
    Selection.Cut Destination:=Range("A860")
    Range("A860").Select
    ActiveWindow.SmallScroll Down:=24
    Range("B880").Select
    Selection.Cut Destination:=Range("A881")
    Range("A881").Select
    ActiveWindow.SmallScroll Down:=18
    Range("B910").Select
    Selection.Cut Destination:=Range("A911")
    Range("A911").Select
    ActiveWindow.SmallScroll Down:=39
    Range("B940").Select
    Selection.Cut Destination:=Range("A941")
    Range("A941").Select
    ActiveWindow.SmallScroll Down:=27
    Range("B970").Select
    Selection.Cut Destination:=Range("A971")
    Range("A971").Select
    ActiveWindow.SmallScroll Down:=24
    Range("B1000").Select
    ActiveWindow.SmallScroll Down:=15
    Selection.Cut Destination:=Range("A1001")
    Range("A1001").Select
    ActiveWindow.SmallScroll Down:=21
    Range("B1030").Select
    Selection.Cut Destination:=Range("A1031")
    Range("A1031").Select
    ActiveWindow.SmallScroll Down:=24
    Range("B1060").Select
    ActiveWindow.SmallScroll Down:=12
    Selection.Cut Destination:=Range("A1061")
    Range("A1061").Select
    ActiveWindow.SmallScroll Down:=33
    Range("B1090").Select
    Selection.Cut Destination:=Range("A1091")
    Range("A1091").Select
    ActiveWindow.SmallScroll Down:=27
    Range("B1120").Select
    Selection.Cut Destination:=Range("A1121")
    Range("A1121").Select
    ActiveWindow.SmallScroll Down:=30
    Range("B1150").Select
    Selection.Cut Destination:=Range("A1151")
    Range("A1151").Select
    ActiveWindow.SmallScroll Down:=30
    Range("B1180").Select
    Selection.Cut Destination:=Range("A1181")
    Range("A1181").Select
    ActiveWindow.SmallScroll Down:=33
    Range("B1210").Select
    Selection.Cut Destination:=Range("A1211")
    Range("B1211").Select
    ActiveWindow.ScrollRow = 1205
    ActiveWindow.ScrollRow = 1201
    ActiveWindow.ScrollRow = 1198
    ActiveWindow.ScrollRow = 1193
    ActiveWindow.ScrollRow = 1188
    ActiveWindow.ScrollRow = 1180
    ActiveWindow.ScrollRow = 1165
    ActiveWindow.ScrollRow = 1155
    ActiveWindow.ScrollRow = 1123
    ActiveWindow.ScrollRow = 1096
    ActiveWindow.ScrollRow = 1017
    ActiveWindow.ScrollRow = 955
    ActiveWindow.ScrollRow = 822
    ActiveWindow.ScrollRow = 780
    ActiveWindow.ScrollRow = 680
    ActiveWindow.ScrollRow = 606
    ActiveWindow.ScrollRow = 502
    ActiveWindow.ScrollRow = 468
    ActiveWindow.ScrollRow = 331
    ActiveWindow.ScrollRow = 314
    ActiveWindow.ScrollRow = 219
    ActiveWindow.ScrollRow = 205
    ActiveWindow.ScrollRow = 140
    ActiveWindow.ScrollRow = 131
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 108
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 79
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 1
    Rows("4:5").Select
    Selection.Delete Shift:=xlUp
    Rows("13:13").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=9
    Rows("22:22").Select
    Selection.Delete Shift:=xlUp
    Rows("24:24").Select
    Selection.Delete Shift:=xlUp
    Range("A27").Select
    ActiveWindow.SmallScroll Down:=6
    Rows("26:26").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=63
    Rows("83:83").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=174
    Rows("262:262").Select
    Selection.Delete Shift:=xlUp
    Range("A265").Select
    ActiveWindow.SmallScroll Down:=183
    Rows("442:442").Select
    Selection.Delete Shift:=xlUp
    Range("A445").Select
    ActiveWindow.SmallScroll Down:=33
    Rows("469:469").Select
    Selection.Delete Shift:=xlUp
    Range("A476").Select
    ActiveWindow.SmallScroll Down:=117
    Rows("595:595").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=129
    Rows("721:721").Select
    Selection.Delete Shift:=xlUp
    Range("A725").Select
    ActiveWindow.SmallScroll Down:=129
    Rows("847:847").Select
    Selection.Delete Shift:=xlUp
    Range("A850").Select
    ActiveWindow.SmallScroll Down:=21
    Rows("867:867").Select
    Selection.Delete Shift:=xlUp
    Range("A874").Select
    ActiveWindow.SmallScroll Down:=18
    Rows("896:896").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=27
    Rows("925:925").Select
    Selection.Delete Shift:=xlUp
    Range("A928").Select
    ActiveWindow.SmallScroll Down:=27
    Rows("954:954").Select
    Selection.Delete Shift:=xlUp
    Range("A958").Select
    ActiveWindow.SmallScroll Down:=36
    Rows("983:983").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=21
    Rows("1012:1012").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=33
    Rows("1041:1041").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=30
    Rows("1070:1070").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=33
    Rows("1099:1099").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=24
    Rows("1128:1128").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=27
    Rows("1157:1157").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=33
    Rows("1186:1186").Select
    Selection.Delete Shift:=xlUp
    Range("B1190").Select
    Selection.End(xlUp).Select
    Range("A4").Select
    Selection.Copy
    Range("A5:A12").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=12
    Range("A13").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A14:A21").Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A23").Select
    ActiveSheet.Paste
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A25").Select
    ActiveSheet.Paste
    Range("A26").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A27").Select
    ActiveWindow.SmallScroll Down:=57
    Range("A27:A82").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=60
    Range("A83").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A84").Select
    ActiveWindow.SmallScroll Down:=174
    Range("A84:A261").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=180
    Range("A262").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A263").Select
    ActiveWindow.SmallScroll Down:=180
    Range("A263:A441").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=183
    Range("A442").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A443").Select
    ActiveWindow.SmallScroll Down:=24
    Range("A443:A468").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=30
    Range("A469").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A470").Select
    ActiveWindow.SmallScroll Down:=123
    Range("A470:A594").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=129
    Range("A595").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A596").Select
    ActiveWindow.SmallScroll Down:=114
    Range("A596:A720").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=126
    Range("A721").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A722").Select
    ActiveWindow.SmallScroll Down:=123
    Range("A722:A846").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=123
    Range("A847").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A848").Select
    ActiveWindow.SmallScroll Down:=18
    Range("A848:A866").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=24
    Range("A867").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A868").Select
    ActiveWindow.SmallScroll Down:=15
    Range("A868:A895").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=27
    Range("A896").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A897").Select
    ActiveWindow.SmallScroll Down:=18
    Range("A897:A924").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=30
    Range("A925").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A926").Select
    ActiveWindow.SmallScroll Down:=18
    Range("A926:A953").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=27
    Range("A954").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A955").Select
    ActiveWindow.SmallScroll Down:=21
    Range("A955:A982").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=24
    Range("A983").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A984").Select
    ActiveWindow.SmallScroll Down:=21
    Range("A984:A1011").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=30
    Range("A1012").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1013").Select
    ActiveWindow.SmallScroll Down:=15
    Range("A1013:A1040").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=30
    Range("A1041").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1042").Select
    ActiveWindow.SmallScroll Down:=12
    Range("A1042:A1069").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=30
    Range("A1070").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1071").Select
    ActiveWindow.SmallScroll Down:=15
    Range("A1071:A1098").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=24
    Range("A1099").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1100").Select
    ActiveWindow.SmallScroll Down:=15
    Range("A1100:A1127").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=30
    Range("A1128").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1129").Select
    ActiveWindow.SmallScroll Down:=15
    Range("A1129:A1156").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=24
    Range("A1157").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1158").Select
    ActiveWindow.SmallScroll Down:=21
    Range("A1158:A1185").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=30
    Range("A1186").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1187:A1195").Select
    ActiveSheet.Paste
    Range("A1193").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    ActiveWindow.SmallScroll Down:=-24
    Columns("C:F").Select
    Range("F1").Activate
    Selection.Delete Shift:=xlToLeft
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("C:C,G:G,K:K,O1,O:O").Select
    Range("O1").Activate
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    Range("C:C,G:G,K:K,O1,O:O,S:S,W:W,AA:AA").Select
    Range("AA1").Activate
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    Range("C:C,G:G,K:K,O1,O:O,S:S,W:W,AA:AA,AE:AE,AI:AI,AM:AM").Select
    Range("AM1").Activate
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    Selection.Cut Destination:=Range("C3")
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("I1").Select
    Selection.Cut Destination:=Range("H3")
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("N1").Select
    Selection.Cut Destination:=Range("M3")
    Range("M3").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("S1").Select
    Selection.Cut Destination:=Range("R3")
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("X1").Select
    Selection.Cut Destination:=Range("W3")
    Range("W3").Select
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AC1").Select
    Selection.Cut Destination:=Range("AB3")
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AH1").Select
    Selection.Cut Destination:=Range("AG3")
    Range("AG3").Select
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AM1").Select
    Selection.Cut Destination:=Range("AL3")
    Columns("AQ:AQ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AR1").Select
    Selection.Cut Destination:=Range("AQ3")
    Range("AQ3").Select
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    Columns("AV:AV").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AW1").Select
    Selection.Cut Destination:=Range("AV3")
    Range("AV3").Select
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("C3").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C1194").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C4:C1194").Select
    Range("C1194").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C3").Select
    Selection.End(xlDown).Select
    Range("H1194").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H4:H1194").Select
    Range("H1194").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H3").Select
    Selection.End(xlDown).Select
    Range("M1194").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M4:M1194").Select
    Range("M1194").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M3").Select
    Selection.End(xlDown).Select
    Range("R1194").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R4:R1194").Select
    Range("R1194").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R3").Select
    Selection.End(xlDown).Select
    Range("W1194").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W4:W1194").Select
    Range("W1194").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W3").Select
    Selection.End(xlDown).Select
    Range("AB1194").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB4:AB1194").Select
    Range("AB1194").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB3").Select
    Selection.End(xlDown).Select
    Range("AG1194").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG4:AG1194").Select
    Range("AG1194").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AG3").Select
    Selection.End(xlDown).Select
    Range("AL1194").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL4:AL1194").Select
    Range("AL1194").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AL3").Select
    Selection.End(xlDown).Select
    Range("AQ1194").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ4:AQ1194").Select
    Range("AQ1194").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AQ3").Select
    Selection.End(xlDown).Select
    Range("AV1194").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AV4:AV1194").Select
    Range("AV1194").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A3").Select
    Application.CutCopyMode = False
    Rows("1:1").Select
    Range("AK1").Activate
    Selection.Delete Shift:=xlUp
    Range("AV2:AZ2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ2").Select
    Selection.End(xlDown).Select
    Range("AQ1194").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL2").Select
    Selection.End(xlDown).Select
    Range("AL1194").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG2").Select
    Selection.End(xlDown).Select
    Range("AG1194").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB2").Select
    Selection.End(xlDown).Select
    Range("AB1194").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W2").Select
    Selection.End(xlDown).Select
    Range("W1194").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R2").Select
    Selection.End(xlDown).Select
    Range("R1194").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M2").Select
    Selection.End(xlDown).Select
    Range("M1194").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H2").Select
    Selection.End(xlDown).Select
    Range("H1194").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C2").Select
    Selection.End(xlDown).Select
    Range("C1194").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("C2").Select
    Selection.End(xlDown).Select
    Range("A11921:B11921").Select
    Range("B11921").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A1194:B11921").Select
    Range("B11921").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("H:AZ").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "MODALIDADES"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "RESPUESTA"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "TARGET"
    Range("A1:G1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C1").Select
    Selection.Cut
    Range("B1").Select
    ActiveSheet.Paste
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=TRIM(RC[1])"
    Range("B2").Select
    Selection.Copy
    Range("C2").Select
    Selection.End(xlDown).Select
    Range("B11921").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("B3:B11921").Select
    Range("B11921").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("B1").Select
    Application.CutCopyMode = False
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B2").Select
    Application.CutCopyMode = False
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1:G1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("B9").Select
    ActiveWindow.SmallScroll Down:=-15
    ActiveWorkbook.Save
End Sub





Sub LIMPIAR()
'
' LIMPIAR Macro
'




'
' LIMPIAR CELDAS
'


    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Rows("1:9").Select
    Range("A9").Activate
    Selection.Delete Shift:=xlUp
    
'
' LIMPIAR CELDAS
'
    
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
    Selection.Delete
    
    
    
'
' SELECCIONAR TODO
'
    
'
Dim rngTemp As Range
Set rngTemp = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
If Not rngTemp Is Nothing Then
    Range(Cells(1, 1), rngTemp).Select
End If
  Selection.ClearFormats

End Sub


Sub LIMPIARv2()

'
' LIMPIAR CELDAS
'


    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Rows("1:10").Select
    Range("A10").Activate
    Selection.Delete Shift:=xlUp
    
'
' LIMPIAR CELDAS
'
    
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
    Selection.Delete
    
    
    
 '
 ' SELECCIONAR TODO
 '
    
Dim rngTemp As Range
Set rngTemp = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
If Not rngTemp Is Nothing Then
    Range(Cells(1, 1), rngTemp).Select
End If
  Selection.ClearFormats

End Sub
