Attribute VB_Name = "Module1"
Public n As Integer
Public Function GetRandom(under, over As Integer) As Integer
     Randomize
     GetRandom = Int((under - over + 1) * Rnd + over)
End Function
