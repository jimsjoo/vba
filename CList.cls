VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'// User-defined Errors
'//
Const ERR_LIST_BEYOND_BOUNDS     As Long = 999
Const ERR_LIST_NOT_INITIALIZED   As Long = 998

Private Item_Data()     As Variant
Private Item_Pointer    As Long
Private Item_Init       As Boolean

Sub Add(vNewValue As Variant)
    Item_Pointer = Item_Pointer + 1
    If Item_Init Then
        ReDim Preserve Item_Data(Item_Pointer)
    Else
        ReDim Item_Data(Item_Pointer)
        Item_Init = True
    End If
    Item_Data(Item_Pointer) = vNewValue
End Sub

Sub Remove(Index As Long)
    Dim Imsi_Data()     As Variant
    Dim i               As Long
    Dim j               As Long
    
    If Item_Init Then
        ReDim Imsi_Data(Item_Pointer - 1)
        j = 1
        For i = 1 To UBound(Item_Data)
            If i <> Index Then
                Imsi_Data(j) = Item_Data(i)
                j = j + 1
            End If
        Next
        
        Item_Pointer = Item_Pointer - 1
        ReDim Item_Data(1 To Item_Pointer)
        
        For i = 1 To UBound(Imsi_Data)
            Item_Data(i) = Imsi_Data(i)
        Next
    End If
End Sub

Function Count() As Long
    Count = Item_Pointer
End Function

Function Item(Index As Long) As Variant
    If Item_Init Then
        If Index < 1 Or Index > Item_Pointer Then
            Err.Raise ERR_LIST_BEYOND_BOUNDS, "CList", "Out of Bound"
        Else
            Item = Item_Data(Index)
        End If
    Else
        Err.Raise ERR_LIST_NOT_INITIALIZED, "CList", "List() Not Initialized"
    End If
End Function

Private Sub Class_Initialize()
    Item_Pointer = 0
    Item_Init = False
    Erase Item_Data()
End Sub

Private Sub Class_Terminate()
    Erase Item_Data()
End Sub
