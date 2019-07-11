
'nLib Version = 1.0

Function dateFormat(format)

' format
' yyyymmdd
' yyyymmddhhmmss
' yyyymmddhhmm

format = LCase(format)

Select Case format

Case "yyyymmdd"
  dateFormat = Replace(Left(Now(),10), "/", "")
Case "yyyymmddhhmmss"
  dateFormat = Replace(Replace(Replace(Now(), "/", ""), ":", ""), " ", "")
Case "yyyymmddhhmm"
  dateFormat = Replace(Replace(Replace(Left(Now(),16), "/", ""), ":", ""), " ", "")
Case Else
  dateFormat = Replace(Left(Now(),10), "/", "")
End Select

End Function
