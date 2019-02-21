Function ReturnBefore(haystack, needle)
    ReturnBefore = Left(haystack, InStr(haystack, needle) - 1)
End Function
