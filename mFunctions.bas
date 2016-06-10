Attribute VB_Name = "mFunctions"
Sub ShuffleArrayInPlace(InArray() As Variant)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShuffleArrayInPlace
' This shuffles InArray to random order, randomized in place.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim N As Long
    Dim Temp As Variant
    Dim J As Long

    Randomize
    For N = LBound(InArray) To UBound(InArray)
        J = CLng(((UBound(InArray) - N) * Rnd) + N)
        If N <> J Then
            Temp = InArray(N)
            InArray(N) = InArray(J)
            InArray(J) = Temp
        End If
    Next N
End Sub
Function CreateArrayWithConsIntegers(FirstNumber As Long, LastNumber As Long) As Variant
Dim arr() As Variant
Dim i As Long
ReDim arr(FirstNumber To LastNumber)

For i = FirstNumber To LastNumber
    arr(i) = i
Next
CreateArrayWithConsIntegers = arr
End Function
