Attribute VB_Name = "QueueExample"
Option Explicit

Sub demoQueue()
    Dim myQueue As New Queue
    
    With myQueue
        .Enqueue "A queue is "
        .Enqueue "a useful data structure "
        .Enqueue "in programming. "
        .Enqueue "It is similar to the ticket queue "
        .Enqueue "outside a cinema hall, "
        .Enqueue "where the first person "
        .Enqueue "entering the queue "
        .Enqueue "is the first person "
        .Enqueue "who gets the ticket."
        
        Debug.Print ">> Get the value of the front of queue :", .Peek
        
        Do While Not .IsEmpty
            Debug.Print .Dequeue()
        Loop
        
        Debug.Print ">> Is myQueue empty?", .IsEmpty
    End With
End Sub
