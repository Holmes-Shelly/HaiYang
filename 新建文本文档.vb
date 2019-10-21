For Each fuel In Sheet3.Range("a1:a157")
    l = Application.WorksheetFunction.CountIf(Sheet3.Range("a1:a157"), fuel)
    MsgBox fuel
    If l > 1 Then
        MsgBox fuel
    End If
Next

Dim fuel(157) As String, assb(157) As String
Dim fuel_sort() As Integer
Dim dict As Object

Set dict = CreateObject("scripting.dictionary")
dict.Add "AA", 1
dict.Add "AB", 2
dict.Add "AC", 3
dict.Add "AD", 4
dict.Add "AE", 5


For Each fu In fuel
    if dict(right(fu,2))