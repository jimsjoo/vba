Attribute VB_Name = "Module7"
Option Explicit

Sub demoCollection()
    Dim myCol As Collection
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Create New Collection
    Set myCol = New Collection
        
    'Add items to Collection
    myCol.Add 10 'Items: 10
    myCol.Add 20 'Items: 10, 20
    myCol.Add 30 'Items: 10, 20, 30
    
    Debug.Print myCol.Count '3
    Debug.Print myCol(1)
    
    Dim it As Variant
    For Each it In myCol
      Debug.Print it '10, 20, 30
    Next it
    
    Dim i As Long
    For i = 1 To myCol.Count
      Debug.Print myCol(i) '10, 20, 30
    Next i
    Stop
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set myCol = Nothing
    Set myCol = New Collection
    
    myCol.Add 10 'Items: 10
    myCol.Add 20 'Items: 10, 20
    
    myCol.Add 30, Before:=1  'Items: 30, 10, 20
    Stop
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set myCol = Nothing
    Set myCol = New Collection
    
    myCol.Add 10 'Items: 10
    myCol.Add 20 'Items: 10, 20

    myCol.Add 30, After:=1  'Items: 10, 30, 20
    Stop
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Remove selected items from Collection
    'Before Items: 10, 20, 30
    myCol.Remove (2) 'Items: 10, 30
    myCol.Remove (2) 'Items: 10
    Stop
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set myCol = Nothing
    Set myCol = New Collection
    
    myCol.Add 10, "Key10"
    myCol.Add 20, "Key20" 'Items: 10, 20
    myCol.Add 30, "Key30" 'Items: 10, 20, 30
    Debug.Print myCol("Key10") 'Returns 10
    
    
End Sub

Sub Store_And_Print_Keys()
    Dim col1 As New Collection
    
    col1.Add Array("first key", "first string"), "first key"
    col1.Add Array("second key", "second string"), "second key"
    col1.Add Array("third key", "third string"), "third key"
    
    Dim item As Variant
        
    '// print items
    For Each item In col1
      Debug.Print item(1)
    Next
    
    '// print keys
    For Each item In col1
      Debug.Print item(0)
    Next
End Sub

Sub Make_Unique_List()
    Dim colFruits As New Collection
    Dim fruit
    
    On Error Resume Next
    For Each fruit In Array("Banana", "Apple", "Orange", "Tomato", "Apple", "Lemon", "Lime", "Lime", "Apple")
       colFruits.Add fruit, CStr(fruit)
    Next
    
    On Error GoTo 0
    For Each fruit In colFruits
        Debug.Print fruit
    Next

End Sub

