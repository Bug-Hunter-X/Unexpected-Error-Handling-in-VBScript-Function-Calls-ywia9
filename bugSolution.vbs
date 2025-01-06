Function MyFunc(param)
  On Error Resume Next
  If IsEmpty(param) Then
    Err.Raise 9999, , "Parameter cannot be empty"
  End If
  If Err.Number <> 0 Then
    'Handle the error appropriately, e.g., log it, display a user-friendly message
    MsgBox "Error: " & Err.Description
    Err.Clear
    MyFunc = Null 'or return an appropriate default value
    Exit Function 
  End If
  ' ... rest of the function
End Function