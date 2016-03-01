Attribute VB_Name = "Module1"
Public cn As New ADODB.Connection

Public Function conecxao() As Boolean

On Error GoTo Trata_Erro:
 cn.Open "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA;PWD=masterkey;DBNAME=DADOS.FDB"
 conecxao = True
 Exit Function
Trata_Erro:
   conecxao = False
   MsgBox Err.Description
   
End Function


