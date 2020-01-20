Attribute VB_Name = "StackExample"
Option Explicit

Sub TestStacks()
    Dim myStack As New Stack
            
    ' Push some items, and then pop them.
    myStack.Push "kept on top of each other."
    Debug.Print "Current Top Item :", myStack.GetTop
    
    myStack.Push "a pile of plates"
    Debug.Print "Current Top Item :", myStack.GetTop
    
    myStack.Push "It is just like"
    Debug.Print "Current Top Item :", myStack.GetTop
    
    myStack.Push "in programming."
    Debug.Print "Current Top Item :", myStack.GetTop
    
    myStack.Push "a useful data structure"
    Debug.Print "Current Top Item :", myStack.GetTop
    
    myStack.Push "is"
    Debug.Print "Current Top Item :", myStack.GetTop
    
    myStack.Push "A stack"
    Debug.Print "Current Top Item :", myStack.GetTop

    Debug.Print
    Debug.Print "--------------------------------------"
    Debug.Print "It's time to pop out"
    Do While Not myStack.IsEmpty
        Debug.Print ">> Item just poped  :", myStack.Pop()
        Debug.Print ">> Current Top Item :", myStack.GetTop
        Debug.Print
    Loop
End Sub

