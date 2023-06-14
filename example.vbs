Sub Macro4 ()
Dim ThisFile As String
ThisFile = ThisWorkbook.Sheets("EMAIL").Range("D5").Value
With ActiveWorkbook
  .Save As Filename:=ThisFile

Dim OutApp As Object
Dim OutMails As Object
    
Set OutApp = CreateObject("Outlook. Application")
  OutApp.Session.logon
Set OutMail OutApp.CreateItem(0)
    
 ActiveWorkbook.Save

    With OutMail
'Referencias al contenido del correo
.To = ThisWorkbook.Sheets("EMAIL").Range("D2").Value
.cc = ThisWorkbook.Sheets("EMAIL").Range("D3").Value.
.Subject = ThisWorkbook.Sheets("EMAIL").Range("D4").Value
.Body = "Este es el mensaje que quieres poner en el cuerpo del correo"
.Attachments. Add ThisWorkbook.Sheets("EMAIL").Range ("D5").Value
.Send

End With
    
Set OutMail = Nothing
Set OutApp
    
Nothing
    
End With
End Sub
