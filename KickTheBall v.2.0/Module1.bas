Attribute VB_Name = "Module1"
Public Function add_record(r As Integer)
Rem adds score if > from previous record
Dim file As String
Dim rec As Integer
file = App.Path + "\records.rec"
Open file For Random As 1
Get #1, 1, rec
If r > rec Then
Put #1, 1, r
End If
Close #1
End Function
Public Function load_record()
Dim file As String
Dim rec As Integer
file = App.Path + "\records.rec"
Open file For Random As 1
Get #1, 1, rec
load_record = rec
Close #1
End Function

