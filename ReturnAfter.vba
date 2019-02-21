Function ReturnAfter(haystack, needle)
    ReturnAfter = Right(haystack, Len(haystack) - InStr(haystack, needle) - Len(needle) + 1)
End Function
