Public Function Mod10CheckDigit(Barcode) As Integer
    Dim i As Integer
    Dim TotalOdd As Integer
    Dim TotalEven As Integer
    Dim Total As Integer
    Barcode = Trim(Barcode)
    'get odd numbers
    For i = 1 To Len(Barcode) Step 2
        TotalOdd = TotalOdd + CInt(Mid(Barcode, i, 1))
    Next i
    TotalOdd = TotalOdd * 3

    'get even numbers
    i = 0
    For i = 2 To Len(Barcode) Step 2
        TotalEven = TotalEven + CInt(Mid(Barcode, i, 1))
    Next i
    
    Total = TotalOdd + TotalEven
    Mod10CheckDigit = 10 - IIf(Right(Total, 1) = 0, 10, _
          Right(Total, 1))
End Function

Public Function CalculateGTIN(NDC, Optional Indicator = "0")
        Application.Volatile
        
        Barcode = Replace(CStr(NDC), "-", "")
        If Len(Barcode) = 10 Then
            Barcode = Indicator & "03" & Barcode
            CalculateGTIN = Barcode & Mod10CheckDigit(Barcode)
        End If
End Function
