Attribute VB_Name = "Module1"
Option Explicit

Sub demo_CListClass()
    Dim lstDemo    As New CList
    
    '// enter samples
    lstDemo.Add "Tumin"
    lstDemo.Add "Lenski"
    lstDemo.Add "Davis"
    lstDemo.Add "Moore"
    lstDemo.Add "Ogburn"
    lstDemo.Add "Simmel"
    
    '// remove the 2nd item from the list
    lstDemo.Remove 2
    
    Dim i           As Long
    
    '// print all items
    For i = 1 To lstDemo.Count
        Debug.Print lstDemo.Item(i)
    Next
    
    Set lstDemo = Nothing
End Sub

