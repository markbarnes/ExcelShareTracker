Function ReturnBetween(haystack, StartString, EndString)
    ReturnBetween = ReturnBefore(ReturnAfter(haystack, StartString), EndString)
End Function
