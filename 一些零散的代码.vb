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

'把F组件的编码提出来
Sub find_f_fuel()
Dim arr1(64) As String, arr2(64) As String
Dim i As Integer, j As Integer, k As Integer
Dim pos As String

Sheets("cycle2_FFF").Select

k = 1
For i = 1 To 15
    For j = 1 To 15
        pos = Cells(i * 2, j + 1).Value
        If Len(pos) = 4 Then
            arr1(k) = Cells(i * 2, j + 1).Value
            arr2(k) = Cells(i * 2 + 1, j + 1).Value
            k = k + 1
        End If
    Next
Next

Range("a34:a97") = Application.Transpose(arr1)
Range("b34:b97") = Application.Transpose(arr2)

End Sub

'周末写的完整程序，判定是否重复完整
Option Base 1
Sub list()
Dim i As Integer, j As Integer, k As Integer, l As Integer, m As Integer
Dim fu As String
Dim stru, star
Dim fuel(157) As String, assb(157) As String
Dim fuel_aa(16) As Integer, fuel_bb(49) As Integer, fuel_cc(28) As Integer, fuel_dd(36) As Integer, fuel_ee(28) As Integer

stru = Array(3, 7, 9, 11, 13, 13, 15, 15, 15, 13, 13, 11, 9, 7, 3)
star = Array(8, 6, 5, 4, 3, 3, 2, 2, 2, 3, 3, 4, 5, 6, 8)

k = 1
For i = 1 To 15
    For j = 1 To stru(i)
        fuel(k) = Cells(i * 2, star(i) + j - 1)
        k = k + 1
    Next
Next

i = 1
j = 1
k = 1
l = 1
m = 1

For Each fu In fuel
    Select Case Left(fu, 2)
        Case "AA"
            fuel_aa(i) = CInt(Right(fu, 2))
            i = i + 1
        Case "AB"
            fuel_bb(j) = CInt(Right(fu, 2))
            j = j + 1
        Case "AC"
            fuel_cc(k) = CInt(Right(fu, 2))
            k = k + 1
        Case "AD"
            fuel_dd(l) = CInt(Right(fu, 2))
            l = l + 1
        Case "AE"
            fuel_ee(m) = CInt(Right(fu, 2))
            m = m + 1
        Case Else
            MsgBox ("attention")
    End Select
Next

Debug.Print fuel_aa(1)
Debug.Print fuel_dd(36)


'Sheet3.Range("a1:a157") = Application.Transpose(arr)

'Debug.Print arr(4)

End Sub