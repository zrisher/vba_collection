VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub Add(varNewItem As Variant)
    Dim qNew As New QueueItem
    qNew.Value = varNewItem
    ' What if the queue is empty? Better point
    ' both the front and rear pointers at the
    ' new item.
    If QueueEmpty Then
        Set qFront = qNew
        Set qRear = qNew
    Else
        Set qRear.NextItem = qNew
        Set qRear = qNew
    End If
End Sub

Public Function Remove() As Variant
    ' Remove an item from the head of the
    ' list, and return its value.
    If QueueEmpty Then
        Remove = Null
    Else
        Remove = qFront.Value
        ' If there's only one item
        ' in the queue, qFront and qRear
        ' will be pointing to the same node.
        ' Use the Is operator to test for that.
        If qFront Is qRear Then
            Set qFront = Nothing
            Set qRear = Nothing
        Else
            Set qFront = qFront.NextItem
        End If
    End If
End Function

Property Get QueueEmpty() As Boolean
    ' Return True if the queue contains
    ' no items.
    QueueEmpty = ((qFront Is Nothing) And (qRear Is Nothing))
End Property

