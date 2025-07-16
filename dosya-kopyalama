'Kodun çalışma mantığını ve nerelerde kullanabileceğinize dair örnekleri https://www.ilerlemedefteri.net/2025/06/excel-vba-ile-dosya-kopyalama-nasil-yapilir.html adresinde bulabilirsiniz.
Sub Kopyala()
    Dim FSO As New FileSystemObject
    Dim KaynakKlasor As String
    Dim KopyalanacakKlasor As String
    Dim KaynakKlasordekiDosya As Object

    KaynakKlasor = "C:\Users\user1\Desktop\hisseseneditarama\tablolar\"
    KopyalanacakKlasor = "C:\Users\user1\Desktop\hisseseneditarama\table\"
   
    Set FSO = CreateObject("Scripting.FileSystemObject")

    For Each KaynakKlasordekiDosya In FSO.GetFolder(KaynakKlasor).Files
         KaynakKlasordekiDosya.Copy KopyalanacakKlasor
    Next KaynakKlasordekiDosya

End Sub
