Sub QuickSort(arr, Lo As Long, Hi As Long)
  Dim varPivot As Variant
  Dim varTmp As Variant
  Dim tmpLow As Long
  Dim tmpHi As Long
  tmpLow = Lo
  tmpHi = Hi
  varPivot = arr((Lo + Hi) \ 2)
  Do While tmpLow <= tmpHi
    Do While arr(tmpLow) < varPivot And tmpLow < Hi
      tmpLow = tmpLow + 1
    Loop
    Do While varPivot < arr(tmpHi) And tmpHi > Lo
      tmpHi = tmpHi - 1
    Loop
    If tmpLow <= tmpHi Then
      varTmp = arr(tmpLow)
      arr(tmpLow) = arr(tmpHi)
      arr(tmpHi) = varTmp
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
    End If
  Loop
  If Lo < tmpHi Then QuickSort arr, Lo, tmpHi
  If tmpLow < Hi Then QuickSort arr, tmpLow, Hi
End Sub
