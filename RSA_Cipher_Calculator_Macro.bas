Sub CreateRSACalculator()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Clear existing data and gridlines setting
    ws.Cells.Clear
    ActiveWindow.DisplayGridlines = True
    
    ' --- Title ---
    With ws.Range("A1:C1")
        .Merge
        .Value = "RSA Cipher: Key Selection Worksheet"
        .Font.Name = "Arial"
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' --- Prime Tester Box ---
    ws.Range("A3").Value = "PRIME TESTER"
    ws.Range("A3").Font.Bold = True
    
    ws.Range("A4").Value = "NUMBER"
    ws.Range("B4").Value = 17
    ws.Range("A5").Value = "Is it Prime Number?"
    
    ' Dynamic array formula for prime validation
    ws.Range("B5").Formula2 = "=IF(B4<2,""No"",IF(B4=2,""Yes"",IF(MOD(B4,2)=0,""No"",IF(SUMPRODUCT(--(MOD(B4,ROW(INDIRECT(""3:""&INT(SQRT(B4)))))=0))>0,""No"",""Yes""))))"
    
    ' Apply borders to the Prime Tester block to match the photo
    With ws.Range("A4:B5").Borders
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    ' --- RSA Parameters Section ---
    ws.Range("A7").Value = "RSA PARAMETERS"
    ws.Range("A7").Font.Bold = True
    
    ws.Range("A8").Value = "FIRST PRIME"
    ws.Range("B8").Value = 17
    
    ws.Range("A9").Value = "SECOND PRIME"
    ws.Range("B9").Value = 19
    
    ws.Range("A10").Value = "Modulus (n)"
    ws.Range("B10").Formula = "=B8*B9"
    
    ws.Range("A11").Value = "Euler Totient (φ)"
    ws.Range("B11").Formula = "=(B8-1)*(B9-1)"
    
    ws.Range("A12").Value = "Public Exponent (e)"
    ws.Range("B12").Value = 5
    ws.Range("C12").Formula = "=IF(GCD(B12,B11)=1,""OK"",""INVALID E"")"
    
    ws.Range("A13").Value = "PRIVATE KEY (d)"
    ' Generates the modular inverse array matching seamlessly
    ws.Range("B13").Formula2 = "=MATCH(1,INDEX(MOD(ROW(INDIRECT(""1:""&B11))*B12,B11),0),0)"
    
    ' --- Verification Section ---
    ws.Range("A15").Value = "VERIFICATION"
    ws.Range("A15").Font.Bold = True
    
    ws.Range("A16").Value = "(e * d) mod φ"
    ws.Range("B16").Formula = "=MOD(B12*B13,B11)"
    
    ws.Range("A17").Value = "Calculation Status"
    ws.Range("B17").Formula = "=IF(B16=1,""OK"",""ERROR"")"
    
    ' --- Formatting Polish ---
    ' Align columns beautifully
    ws.Range("A3:A17").HorizontalAlignment = xlLeft
    ws.Range("B4:C17").HorizontalAlignment = xlCenter
    ws.Columns("A:C").AutoFit
    
    ' Bold the important output targets
    ws.Range("B5,B10,B11,B13,B17").Font.Bold = True
    
    MsgBox "RSA Calculator Template Built Successfully!", vbInformation, "Done"
End Sub