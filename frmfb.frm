VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Acessando o FireBird"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exibir Listas dos Clientes"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sql As String

cn.ConnectionString = "Provider=ZStyle IBOLE Provider;Data Source=c:\teste\Employee.gdb;UID=sysdba;password=masterkey"

sql = "Select * From Employees"
       
cn.Open

Set rs = cn.Execute(sql)

Do While Not rs.EOF
     
     List1.AddItem rs(0) & vbTab & rs(1) & vbTab & rs(2)
     rs.MoveNext
Loop

rs.Close
cn.Close

End Sub

Private Sub Command2_Click()


Dim sql As String
Dim i, rowCount As Integer
Dim lineItem As String
Dim itemArray


cn.Open "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA;PWD=masterkey;DBNAME=C:\Users\pc1\Desktop\Nova pasta (2)\s\DADOS.FDB"


'Inclui itens no banco de dados
rowCount = List1.ListCount 'obtem o numero de linhas no listbox

Do Until i = rowCount

   lineItem = List1.List(i)
   itemArray = Split(lineItem, vbTab)
            
   sql = "Insert Into States (State_Code, State_Name) Values" _
         & "('" & UCase(itemArray(0)) & "','" & itemArray(1) & "')"
   i = i + 1
   cn.Execute sql
Loop

cn.Close
List1.Clear

End Sub

Private Sub Command3_Click()
cn.Open "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA;PWD=masterkey;DBNAME=DADOS.FDB"



sql = "Select * From clientes"
       

Set rs = cn.Execute(sql)

Do While Not rs.EOF
     
     List2.AddItem rs(0) & vbTab & rs(1)
     rs.MoveNext
Loop

rs.Close
cn.Close
End Sub

