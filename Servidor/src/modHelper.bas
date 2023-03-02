Attribute VB_Name = "modHelper"
Public Function IsArrayInitialized(ByRef arr As Variant) As Boolean

  Dim rv As Long

  On Error Resume Next

  rv = UBound(arr)
  IsArrayInitialized = (Err.Number = 0)

End Function
