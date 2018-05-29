Attribute VB_Name = "Helpers"
Public Function CollectUniques(rng As Range) As Collection
    
    Dim varArray As Variant, var As Variant
    Dim col As Collection
    
    'Guard clause - if Range is nothing, return a Nothing collection
    'Guard clause - if Range is empty, return a Nothing collection
    If rng Is Nothing Or WorksheetFunction.CountA(rng) = 0 Then
        Set CollectUniques = col
        Exit Function
    End If
        
    If rng.Count = 1 Then '<~ check for a single cell range
        Set col = New Collection
        col.Add Item:=CStr(rng.Value), Key:=CStr(rng.Value)
    Else '<~ otherwise the range contains multiple cells
        
        'Convert the passed-in range to a Variant array for SPEED and bind the Collection
        varArray = rng.Value
        Set col = New Collection
        
        'Ignore errors temporarily, as each attempt to add a repeat
        'entry to the collection will cause an error
        On Error Resume Next
        
            'Loop through everything in the variant array, adding
            'to the collection if it's not an empty string
            For Each var In varArray
                If CStr(var) <> vbNullString Then
                    col.Add Item:=CStr(var), Key:=CStr(var)
                End If
            Next var
    
        On Error GoTo 0
    End If
    
    'Return the contains-uniques-only collection
    Set CollectUniques = col
    
End Function
