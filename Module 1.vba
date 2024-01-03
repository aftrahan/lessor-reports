Public Function InitialTime(hr As Double, min As Double)
'Allows entry of times above 10,000 hrs as calculable values

    Dim hr1 As Double
    Dim min1 As Double
    
    hr1 = hr / 24
    min1 = min / 24 / 60
    InitialTime = hr1 + min1
    
End Function

Public Sub CreateNewMonth()
    Dim prevMonth As Worksheet
    Dim newMonth As Worksheet
    Dim name As String
    
    Set prevMonth = ThisWorkbook.Worksheets(1)
    
    Worksheets("Blank").Copy Before:=prevMonth
    
    Set newMonth = ThisWorkbook.Worksheets(1)
    
    ' Set new report date & generate sheet name
    newMonth.Range("L8") = Application.WorksheetFunction _
            .EoMonth(prevMonth.Range("L8"), 1)

    newMonth.name = Format(newMonth.Range("L8").Value, "mmm yyyy")
            
    ' Insert date of last report
    newMonth.Range("E12, E23, E31, M13") = prevMonth.Range("L8")
    
    ' Insert values from last report as cell references
' Airframe
    newMonth.Range("F12").Formula = "='" + prevMonth.name + "'!F13"
    newMonth.Range("G12").Formula = "='" + prevMonth.name + "'!G13"
    
' Engine 1
    newMonth.Range("F23").Formula = "='" + prevMonth.name + "'!F24"
    newMonth.Range("G23").Formula = "='" + prevMonth.name + "'!G24"
    newMonth.Range("F26").Formula = "='" + prevMonth.name + "'!F26 +'" _
            + newMonth.name + "'!F25"
    newMonth.Range("G26").Formula = "='" + prevMonth.name + "'!G26 +'" _
            + newMonth.name + "'!G25"
' Engine 2
    newMonth.Range("F31").Formula = "='" + prevMonth.name + "'!F32"
    newMonth.Range("G31").Formula = "='" + prevMonth.name + "'!G32"
    newMonth.Range("F34").Formula = "='" + prevMonth.name + "'!F34 +'" _
            + newMonth.name + "'!F33"
    newMonth.Range("G34").Formula = "='" + prevMonth.name + "'!G34 +'" _
            + newMonth.name + "'!G33"
    
' APU times and meter readings
    newMonth.Range("N13").Formula = "='" + prevMonth.name + "'!N14"
    newMonth.Range("O13").Formula = "='" + prevMonth.name + "'!O14"
    
    newMonth.Range("M17").Formula = "='" + prevMonth.name + "'!M18"
    newMonth.Range("N17").Formula = "='" + prevMonth.name + "'!N18"
    newMonth.Range("O17").Formula = "='" + prevMonth.name + "'!O18"
    
' Landing gears
    newMonth.Range("N22").Formula = "='" + prevMonth.name + "'!N22 +'" _
            + newMonth.name + "'!F14"
    newMonth.Range("O22").Formula = "='" + prevMonth.name + "'!O22 +'" _
            + newMonth.name + "'!G14"
    newMonth.Range("N23").Formula = "='" + prevMonth.name + "'!N23 +'" _
            + newMonth.name + "'!F14"
    newMonth.Range("O23").Formula = "='" + prevMonth.name + "'!O23 +'" _
            + newMonth.name + "'!G14"
            
    newMonth.Range("N25").Formula = "='" + prevMonth.name + "'!N25 +'" _
            + newMonth.name + "'!F14"
    newMonth.Range("O25").Formula = "='" + prevMonth.name + "'!O25 +'" _
            + newMonth.name + "'!G14"
    newMonth.Range("N26").Formula = "='" + prevMonth.name + "'!N26 +'" _
            + newMonth.name + "'!F14"
    newMonth.Range("O26").Formula = "='" + prevMonth.name + "'!O26 +'" _
            + newMonth.name + "'!G14"
            
    newMonth.Range("N28").Formula = "='" + prevMonth.name + "'!N28 +'" _
            + newMonth.name + "'!F14"
    newMonth.Range("O28").Formula = "='" + prevMonth.name + "'!O28 +'" _
            + newMonth.name + "'!G14"
    newMonth.Range("N29").Formula = "='" + prevMonth.name + "'!N29 +'" _
            + newMonth.name + "'!F14"
    newMonth.Range("O29").Formula = "='" + prevMonth.name + "'!O29 +'" _
            + newMonth.name + "'!G14"
    
' Next scheduled inspection
    newMonth.Range("K33") = prevMonth.Range("K33")
    newMonth.Range("M33") = prevMonth.Range("M33")
    
End Sub




