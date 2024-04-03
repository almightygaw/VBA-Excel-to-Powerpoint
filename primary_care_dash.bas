Attribute VB_Name = "primary_care_dash"
' user instructions:
' 1. run get_provider_names, paste source file path into text box
' 2. list of provider last names (lower case) in Immediate Window
' 3. copy first name, run primary_care_dash, paste name in text box
' 4. repeat steps 1 thru 3 (run primary_care_dash for each provider separately because running it for everyone at once crashes Excel)

Sub get_provider_names()
  
  ' source file
  Dim sourceFilePath As String
  sourceFilePath = "Q:\FPO Business Development\Glenn White\Primary Care Dashboards\copy_data_fy2024_primary care provider dashboards.xlsx" ' InputBox("source file path: ")
  Dim sourceWb As Workbook: Set sourceWb = Workbooks.Open(sourceFilePath)
  
  ' source worksheets
  Dim pressGaney As Worksheet: Set pressGaney = sourceWb.Sheets("Press Ganey")

  ' loop errors and crashes Excel. run for each provider individually:
  pressGaney.Activate
  For Each i In pressGaney.Range(Range("A4"), Range("A4").End(xlDown))
    If Not InStr(LCase(i), "total") > 0 Then
      Debug.Print LCase(WorksheetFunction.Trim(Right(WorksheetFunction.Substitute(i, " ", WorksheetFunction.Rept(" ", 255)), 255)))
    End If
  Next i
  
  
End Sub


Sub primary_care_dash()

  Dim x As String
  x = InputBox("enter provider last name (text after last space in name, lower-case): ")  ' paste lower-case provider last name from get_provider_names

  ' source file
  Dim sourceFilePath As String
  sourceFilePath = "Q:\FPO Business Development\Glenn White\Primary Care Dashboards\copy_data_fy2024_primary care provider dashboards.xlsx"
  Dim sourceWb As Workbook: Set sourceWb = Workbooks.Open(sourceFilePath)

  ' source worksheets
  Dim pressGaney As Worksheet: Set pressGaney = sourceWb.Sheets("Press Ganey")
  Dim diabetes As Worksheet: Set diabetes = sourceWb.Sheets("Diabetes")
  Dim medicareAwv24 As Worksheet: Set medicareAwv24 = sourceWb.Sheets("Medicare AWV 2024 (2)")
  Dim combo3 As Worksheet: Set combo3 = sourceWb.Sheets("Combo 3")
  Dim wRVU As Worksheet: Set wRVU = sourceWb.Sheets("wRVUs")
  Dim encCloseTime As Worksheet: Set encCloseTime = sourceWb.Sheets("Encounter Closure Time")
  Dim totalVisits As Worksheet: Set totalVisits = sourceWb.Sheets("Total Visits")
  Dim averageLead As Worksheet: Set averageLead = sourceWb.Sheets("Average Lead Time")
    


  ' destination file: pptx
  Dim pptApp As PowerPoint.Application     ' pptx application object
  Dim pptPres As PowerPoint.Presentation   ' pptx presentation object
  Dim slide1 As PowerPoint.Slide           ' pptx slide objects
  Dim slide2 As PowerPoint.Slide
  Dim slide3 As PowerPoint.Slide

  ' new pptx file
  Set pptApp = New PowerPoint.Application  ' new instance of pptx
  pptApp.Visible = True

  ' new pptx presentation
  Set pptPres = pptApp.Presentations.Add
  With pptPres.PageSetup
    .SlideHeight = 8.5 * 72
    .SlideWidth = 11 * 72
  End With

  pptPres.PageSetup.SlideOrientation = msoOrientationVertical
  
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' SLIDE 1 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' new pptx slide
  Set slide1 = pptPres.Slides.Add(1, ppLayoutBlank)

  ' array of provider names from pressGaney
  Dim j As Long
  j = 1

  Dim providerName As String

  pressGaney.Activate
  For Each i In pressGaney.Range(Range("A4"), Range("A4").End(xlDown))
    If InStr(LCase(i.Value2), x) > 0 And _
      Not InStr(LCase(i.Value2), "total") > 0 Then
      providerName = i
      Exit For
    End If
  Next i


  ' header: left
  With slide1.Shapes.AddShape(msoShapeRectangle, _
    Left:=0, Top:=0, Width:=200, Height:=50)
    .Fill.Visible = msoFalse
    .Line.Visible = msoFalse
    .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
    .TextFrame.TextRange.Font.Name = "Calibri"
    .TextFrame.TextRange.Font.Size = 10
    .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    .TextFrame.TextRange.Text = "Columbia Primary Care Provider Dashboard" & _
                       vbCrLf & providerName & _
                       vbCrLf & VBA.MonthName(Month(Date - 28)) & " " & VBA.Year(Date - 28)
  End With

  ' header: right (Columbia logo)
  slide1.Shapes.AddPicture FileName:="Q:\FPO Business Development\Glenn White\Primary Care Dashboards\" & _
                                          "columbia header.jpg", _
    LinkToFile:=msoFalse, _
    SaveWithDocument:=msoTrue, _
    Left:=310, Top:=0, Width:=300, Height:=50

  ' "SERVICE" banner
  With slide1.Shapes.AddShape(msoShapeRectangle, _
    Left:=0, Top:=60, Width:=8.5 * 72, Height:=25)
    .Fill.Visible = msoTrue
    .Fill.ForeColor.RGB = RGB(50, 100, 160)
    .Line.Visible = msoFalse
    .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    .TextFrame.TextRange.Font.Name = "Calibri"
    .TextFrame.TextRange.Font.Size = 12
    .TextFrame2.TextRange.Font.Smallcaps = msoTrue
    .TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    .TextFrame.TextRange.Text = "Service"
  End With

  ' text box: "Press Ganey FY2023 (National Facilities Percentile Rank, Rate Provider 0-10)"
    With slide1.Shapes.AddShape(msoShapeRectangle, _
    Left:=100, Top:=70, Width:=400, Height:=60)
    .Fill.Visible = msoFalse
    .Line.Visible = msoFalse
    .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    .TextFrame.TextRange.Font.Name = "Calibri"
    .TextFrame.TextRange.Font.Size = 12
    .TextFrame.TextRange.Font.Bold = msoTrue
    .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    .TextFrame.TextRange.Text = "Press Ganey FY2023 (National Facilities Percentile Rank, Rate Provider 0-10)"
    .TextFrame.MarginBottom = 0
    .TextFrame.MarginLeft = 5
    .TextFrame.MarginRight = 0
    .TextFrame.MarginTop = 5
  End With


  ' Press Ganey table1
  j = 1

  ' get size of range to copy/paste
  For Each i In pressGaney.Range(Range("A4"), Range("A4").End(xlDown))
    If Not InStr(LCase(i.Value2), x) > 0 Then
      j = j + 1
    End If
  Next i

  ' filter pressGaney data
  For Each i In pressGaney.Range(Range("A3"), Range("A3").End(xlDown))
    If InStr(LCase(i), x) > 0 And _
      Not InStr(LCase(i), "total") > 0 Then
      pressGaney.Range(Range("A3"), Range("O3").End(xlDown)).AutoFilter Field:=1, Criteria1:=i
      Exit For
    End If
  Next i


  ' copy data to another location on same ws (get a version of filtered data without filter buttons)
  pressGaney.Range(Range("A3").Offset(0, 1), Range("A3").Offset(j, 14)).Copy _
    Destination:=pressGaney.Range("A3").Offset(50, 1)
  Application.CutCopyMode = False

  ' copy new filtered data without filter buttons and paste to pptx
  pressGaney.Cells.WrapText = False
  pressGaney.UsedRange.Columns.AutoFit
  pressGaney.Range(Range("A3").Offset(50, 1), Range("A3").Offset(51, 9)).Copy
  Set pgTable1 = slide1.Shapes.PasteSpecial(ppPasteEnhancedMetafile)

  ' scale/position pgTable1
  pressGaney.AutoFilter.ShowAllData
  With pgTable1(1)
    .Top = 120
    .Left = 20
    .Width = 572
    .Height = 210
'    .ScaleWidth 0.75, msoTrue
    .ScaleHeight 0.75, msoTrue
  End With


  ' hide unused rows
  For Each i In pressGaney.Range("B26:N26")
    If i.Value2 = "" Then
      i.EntireColumn.Hidden = True
    End If
  Next i


  ' Press Ganey key
  Dim pgTable2 As PowerPoint.ShapeRange

  j = 1

  For Each i In pressGaney.Range("P:P")
    If i.Value2 <> "Key" Then
      j = j + 1
    Else
      Exit For
    End If
  Next i

  pressGaney.Range(Range("P" & j), Range("Q" & j).Offset(5, 0)).Copy
  Set pgTable2 = slide1.Shapes.PasteSpecial(ppPasteEnhancedMetafile)
  With pgTable2(1)
    .Top = 155
    .Left = 525
    .ScaleWidth 0.85, msoTrue
    .ScaleHeight 0.85, msoTrue
  End With


  ' "QUALITY" banner
  With slide1.Shapes.AddShape(msoShapeRectangle, _
    Left:=0, Top:=240, Width:=8.5 * 72, Height:=25)
    .Fill.Visible = msoTrue
    .Fill.ForeColor.RGB = RGB(50, 100, 160)
    .Line.Visible = msoFalse
    .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    .TextFrame.TextRange.Font.Name = "Quality"
    .TextFrame.TextRange.Font.Size = 12
    .TextFrame2.TextRange.Font.Smallcaps = msoTrue
    .TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    .TextFrame.TextRange.Text = "Service"
  End With


  On Error Resume Next
  
  ' diabetesChart
  Dim diabetesChart As Object

  For Each i In diabetes.ChartObjects
    If InStr(LCase(diabetes.Range(i.TopLeftCell.Address).Offset(-1, 0).Value2), x) > 0 Then
      i.Copy
      Set diabetesChart = slide1.Shapes.Paste
      With diabetesChart(1)
        .Top = 280
        .Left = 20
        .Width = 572
        .Height = 210
      End With
      Exit For
    End If
  Next i


  ' medicareAwv24Chart
  Dim medicareAwv24Chart As Object
  
  For Each i In medicareAwv24.ChartObjects
    If InStr(LCase(medicareAwv24.Range(i.TopLeftCell.Address).Offset(-1, 0).Value2), x) > 0 Then
      i.Copy
      Set medicareAwv24Chart = slide1.Shapes.Paste
      With medicareAwv24Chart(1)
        .Top = 510
        .Left = 20
        .Width = 572
        .Height = 210
      End With
      Exit For
    End If
  Next i

  ' footer: left
  With slide1.Shapes.AddShape(msoShapeRectangle, _
    Left:=5, Top:=650, Width:=600, Height:=50)
    .Top = slide1.Application.Top + slide1.Application.Height + 50
    .Fill.Visible = msoFalse
    .Line.Visible = msoFalse
    .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
    .TextFrame.TextRange.Font.Name = "Calibri"
    .TextFrame.TextRange.Font.Size = 10
    .TextFrame.TextRange.Font.Italic = True
    .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    .TextFrame.TextRange.Text = "Data: Traditional Medicare ACO attributed patients, % completed represents the cumulative" & _
                       vbCrLf & "percentage of attributed patients who had an Annual Wellness Visit in Calendar Year 2022"
    .TextFrame.MarginBottom = 5
    .TextFrame.MarginLeft = 5
    .TextFrame.MarginRight = 0
    .TextFrame.MarginTop = 0
  End With

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' SLIDE 2 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' new pptx slide
  Set slide2 = pptPres.Slides.Add(2, ppLayoutBlank)

  ' header: left
  With slide2.Shapes.AddShape(msoShapeRectangle, _
    Left:=0, Top:=0, Width:=200, Height:=50)
    .Fill.Visible = msoFalse
    .Line.Visible = msoFalse
    .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
    .TextFrame.TextRange.Font.Name = "Calibri"
    .TextFrame.TextRange.Font.Size = 10
    .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    .TextFrame.TextRange.Text = "Columbia Primary Care Provider Dashboard" & _
                       vbCrLf & providerName & _
                       vbCrLf & VBA.MonthName(Month(Date - 28)) & " " & VBA.Year(Date - 28)
  End With

  ' header: right (Columbia logo)
  slide2.Shapes.AddPicture FileName:="Q:\FPO Business Development\Glenn White\Primary Care Dashboards\" & _
                                          "columbia header.jpg", _
    LinkToFile:=msoFalse, _
    SaveWithDocument:=msoTrue, _
    Left:=310, Top:=0, Width:=300, Height:=50

  ' "QUALITY (CONTINUED)" banner
  With slide2.Shapes.AddShape(msoShapeRectangle, _
    Left:=0, Top:=60, Width:=8.5 * 72, Height:=25)
    .Fill.Visible = msoTrue
    .Fill.ForeColor.RGB = RGB(50, 100, 160)
    .Line.Visible = msoFalse
    .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    .TextFrame.TextRange.Font.Name = "Calibri"
    .TextFrame.TextRange.Font.Size = 12
    .TextFrame2.TextRange.Font.Smallcaps = msoTrue
    .TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    .TextFrame.TextRange.Text = "Quality (continued)"
  End With


  ' combo3Chart
  Dim combo3Chart As Object
  
  For Each i In combo3.ChartObjects
    If InStr(LCase(combo3.Range(i.TopLeftCell.Address).Offset(-1, 0).Value2), x) > 0 Then
      i.Copy
      Set combo3Chart = slide2.Shapes.Paste
      With combo3Chart(1)
        .Top = 90
        .Left = 20
        .Width = 572
        .Height = 210
      End With
      Exit For
    End If
  Next i


  ' "FINANCE" banner
  With slide2.Shapes.AddShape(msoShapeRectangle, _
    Left:=0, Top:=310, Width:=8.5 * 72, Height:=25)
    .Fill.Visible = msoTrue
    .Fill.ForeColor.RGB = RGB(50, 100, 160)
    .Line.Visible = msoFalse
    .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    .TextFrame.TextRange.Font.Name = "Calibri"
    .TextFrame.TextRange.Font.Size = 12
    .TextFrame2.TextRange.Font.Smallcaps = msoTrue
    .TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    .TextFrame.TextRange.Text = "Finance"
  End With


  ' wRVUChart
  Dim wRVUChart As Object
  
  For Each i In wRVU.ChartObjects
    If InStr(LCase(wRVU.Range(i.TopLeftCell.Address).Offset(-1, 0).Value2), x) > 0 Then
      i.Copy
      Set wRVUChart = slide2.Shapes.Paste
      With wRVUChart(1)
        .Top = 340
        .Left = 20
        .Width = 572
        .Height = 210
      End With
      Exit For
    End If
  Next i


  ' encCloseTimeChart
  Dim encCloseTimeChart As Object
  
  For Each i In encCloseTime.ChartObjects
    If InStr(LCase(encCloseTime.Range(i.TopLeftCell.Address).Offset(-1, 0).Value2), x) > 0 Then
      i.Copy
      Set encCloseTimeChart = slide2.Shapes.Paste
      With encCloseTimeChart(1)
        .Top = 560
        .Left = 20
        .Width = 572
        .Height = 210
      End With
      Exit For
    End If
  Next i

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' SLIDE 3 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' new pptx slide
  Set slide3 = pptPres.Slides.Add(3, ppLayoutBlank)

  ' header: left
  With slide3.Shapes.AddShape(msoShapeRectangle, _
    Left:=0, Top:=0, Width:=200, Height:=50)
    .Fill.Visible = msoFalse
    .Line.Visible = msoFalse
    .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
    .TextFrame.TextRange.Font.Name = "Calibri"
    .TextFrame.TextRange.Font.Size = 10
    .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    .TextFrame.TextRange.Text = "Columbia Primary Care Provider Dashboard" & _
                       vbCrLf & providerName & _
                       vbCrLf & VBA.MonthName(Month(Date - 28)) & " " & VBA.Year(Date - 28)
  End With

  ' header: right (Columbia logo)
  slide3.Shapes.AddPicture FileName:="Q:\FPO Business Development\Glenn White\Primary Care Dashboards\" & _
                                          "columbia header.jpg", _
    LinkToFile:=msoFalse, _
    SaveWithDocument:=msoTrue, _
    Left:=310, Top:=0, Width:=300, Height:=50

  ' "GROWTH" banner
  With slide3.Shapes.AddShape(msoShapeRectangle, _
    Left:=0, Top:=60, Width:=8.5 * 72, Height:=25)
    .Fill.Visible = msoTrue
    .Fill.ForeColor.RGB = RGB(50, 100, 160)
    .Line.Visible = msoFalse
    .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    .TextFrame.TextRange.Font.Name = "Calibri"
    .TextFrame.TextRange.Font.Size = 12
    .TextFrame2.TextRange.Font.Smallcaps = msoTrue
    .TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    .TextFrame.TextRange.Text = "Growth"
  End With


  ' totalVisitsChart
  Dim totalVisitsChart As Object
  
  For Each i In totalVisits.ChartObjects
    If InStr(LCase(totalVisits.Range(i.TopLeftCell.Address).Offset(-1, 0).Value2), x) > 0 Then
      i.Copy
      Set totalVisitsChart = slide3.Shapes.Paste
      With totalVisitsChart(1)
        .Top = 90
        .Left = 20
        .Width = 572
        .Height = 210
      End With
      Exit For
    End If
  Next i


  ' "ACCESS" banner
  With slide3.Shapes.AddShape(msoShapeRectangle, _
    Left:=0, Top:=315, Width:=8.5 * 72, Height:=25)
    .Fill.Visible = msoTrue
    .Fill.ForeColor.RGB = RGB(50, 100, 160)
    .Line.Visible = msoFalse
    .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    .TextFrame.TextRange.Font.Name = "Calibri"
    .TextFrame.TextRange.Font.Size = 12
    .TextFrame2.TextRange.Font.Smallcaps = msoTrue
    .TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    .TextFrame.TextRange.Text = "Access"
  End With


  ' averageLeadChart
  Dim averageLeadChart As Object
  
  For Each i In averageLead.ChartObjects
    If InStr(LCase(averageLead.Range(i.TopLeftCell.Address).Offset(-1, 0).Value2), x) > 0 Then
      i.Copy
      Set averageLeadChart = slide3.Shapes.Paste
      With averageLeadChart(1)
        .Top = 350
        .Left = 20
        .Width = 572
        .Height = 210
      End With
      Exit For
    End If
  Next i


  ' averageLeadChart footer
    With slide3.Shapes.AddShape(msoShapeRectangle, _
    Left:=15, Top:=560, Width:=600, Height:=50)
    .Fill.Visible = msoFalse
    .Line.Visible = msoFalse
    .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
    .TextFrame.TextRange.Font.Name = "Calibri"
    .TextFrame.TextRange.Font.Size = 10
    .TextFrame.TextRange.Font.Italic = True
    .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    .TextFrame.TextRange.Text = "*Average Lead Time = time lapsed between date appointment was created " & _
                       vbCrLf & "and scheduled appointment date (average for all completed visits)"
    .TextFrame.MarginBottom = 5
    .TextFrame.MarginLeft = 5
    .TextFrame.MarginRight = 0
    .TextFrame.MarginTop = 0
  End With

  Dim fiscalYear As Long
  If Month(Date) <= 6 Then
    fiscalYear = Format$(Date, "yy")
  Else
    fiscalYear = Format$(Date + 365, "yy")
  End If

  ' save presentation
  pptPres.SaveAs ("Q:\FPO Business Development\Glenn White\Primary Care Dashboards\" & providerName & "_FY" & fiscalYear & ".pptx")
  pptPres.Close


  
 End Sub






