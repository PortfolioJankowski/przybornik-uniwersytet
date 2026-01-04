Attribute VB_Name = "Utilities"
Option Compare Database
Option Explicit

Public Function IsNullOrWhiteSpace(ByVal text As Variant) As Boolean
    If IsNull(text) Then
        IsNullOrWhiteSpace = True
    Else
        IsNullOrWhiteSpace = Trim(CStr(text)) = ""
    End If
End Function


'------------------------------------------------------------
' NzSafe
' Zamienia wartoœæ Null na wartoœæ domyœln¹.
'
' value        - dowolna wartoœæ (Variant), np. z formularza lub Recordsetu
' defaultValue - wartoœæ zwracana, gdy value = Null
'
' Przyk³ady:
'   NzSafe(Null, "")        -> ""
'   NzSafe(Null, 0)         -> 0
'   NzSafe("Jan", "")       -> "Jan"
'   NzSafe(Empty, "")       -> ""
'
' Dlaczego:
' - chroni przed "Invalid use of Null"
' - pozwala trzymaæ Null TYLKO na granicy UI/DB
' - encje domenowe nie powinny znaæ Null
'------------------------------------------------------------
Public Function NzSafe(ByVal value As Variant, _
                       Optional ByVal defaultValue As Variant) As Variant

    If IsNull(value) Then
        NzSafe = defaultValue
    Else
        NzSafe = value
    End If

End Function

