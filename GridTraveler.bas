Sub Test()
    Debug.Print gridTraveler(2, 3)
End Sub

Function gridTraveler(m, n, Optional memo As Scripting.Dictionary) As LongLong
    Dim key As String
    
        key = m & "," & n

        If memo Is Nothing Then Set memo = New Scripting.Dictionary
        
        If memo.Exists(key) Then
            gridTraveler = memo(key)
            Exit Function
        End If
        
        If m = 1 And n = 1 Then
            gridTraveler = 1
            Exit Function
        End If
        
        If m = 0 Or n = 0 Then
            gridTraveler = 0
            Exit Function
        End If
        
        memo.Add key, gridTraveler(m - 1, n, memo) + gridTraveler(m, n - 1, memo)
        
        gridTraveler = memo(key)
    
End Function
