Attribute VB_Name = "Module1"
Type sideinfo
  edge As Integer
  empty As Integer
  pipe As Integer
End Type
Type brdinfo
  pieceno As Integer
  available As Integer
  side(4) As sideinfo
End Type
Type pieceinfo
top As Integer
bottom As Integer
left As Integer
right As Integer
End Type


