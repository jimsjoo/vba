Attribute VB_Name = "Module7"
Option Explicit

'// How to get approximation value from array or list

Sub getApproximation()
    Dim arrSource
    Dim arrABS()
    Dim dblLookup   As Double
    Dim i           As Long
    
    arrSource = Array(2, 1.5, 1, 0.5, 0, -0.5, -1, -1.5, -2)
    dblLookup = 0.251
    
    ReDim arrABS(UBound(arrSource))
    
    For i = 0 To UBound(arrSource)
        arrABS(i) = Abs(arrSource(i) - dblLookup)
    Next

    With WorksheetFunction
        Dim approx As Double
        approx = .Index(arrSource, .Match(.Min(arrABS()), arrABS(), 0))
        Debug.Print approx
    End With
    
End Sub

