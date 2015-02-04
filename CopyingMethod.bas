Attribute VB_Name = "CopyingMethod"
'Type MyMemento
'    Value1 As Integer
'    Value2 As String
'End Type
'
'Private Memento As MyMemento

'Friend Sub SetMemento(NewMemento As MyMemento)
'    Memento = NewMemento
'End Sub

Public Function Copy() As Applicant
    Dim Result As Applicant
    Set Result = New Applicant
    'Call Result.SetMemento(Memento)
    Set Copy = Result
End Function
