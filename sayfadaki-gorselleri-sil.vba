Sub gorselleriSil()
  ActiveSheet.Pictures.Delete
End Sub

Sub gorselleriSil()
    Dim aralik As Range
    Dim sekil As Shape
    
    Set aralik = Application.InputBox(Prompt:="Lütfen görsellerin silineceği aralığı seçin", Type:=8)
    
    If Not aralik Is Nothing Then
        For Each sekil In ActiveSheet.Shapes
            If sekil.Type = msoPicture Then
                If Not Intersect(sekil.TopLeftCell, aralik) Is Nothing Then
                    sekil.Delete
                End If
            End If
        Next sekil
    End If
End Sub
