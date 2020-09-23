VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "MsFlexCool"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox menos 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   1
      Left            =   6600
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox mais 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   0
      Left            =   6360
      Picture         =   "Form1.frx":0381
      ScaleHeight     =   195
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox arrow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   0
      Left            =   1920
      Picture         =   "Form1.frx":0709
      ScaleHeight     =   180
      ScaleWidth      =   945
      TabIndex        =   5
      Top             =   240
      Width           =   945
   End
   Begin VB.PictureBox arrow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   1920
      Picture         =   "Form1.frx":0E4B
      ScaleHeight     =   180
      ScaleWidth      =   945
      TabIndex        =   4
      Top             =   0
      Width           =   945
   End
   Begin VB.PictureBox cx 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   0
      Left            =   5760
      Picture         =   "Form1.frx":158D
      ScaleHeight     =   195
      ScaleWidth      =   210
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox cx 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   1
      Left            =   6000
      Picture         =   "Form1.frx":1A9F
      ScaleHeight     =   195
      ScaleWidth      =   210
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOAD FIRST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   4440
      Width           =   3015
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   21
      FixedCols       =   0
      BackColor       =   16777215
      ForeColorFixed  =   0
      BackColorSel    =   16777215
      ForeColorSel    =   4194304
      WordWrap        =   -1  'True
      AllowUserResizing=   1
      FormatString    =   $"Form1.frx":1FB1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   670
      X2              =   670
      Y1              =   240
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   360
      X2              =   360
      Y1              =   240
      Y2              =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Click here"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "transferred"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   9
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "used"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   8
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DB As DAO.Database
Public RS As DAO.Recordset
Public DBC As DAO.Database
Public RSC As DAO.Recordset
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim valor4 As Double
Dim valor5 As Double
Dim add As Long


Private Sub Command1_Click()
valor1 = 0
valor2 = 0
valor3 = 0
valor4 = 0
valor5 = 0

MSFlexGrid1.Visible = False
MSFlexGrid1.Redraw = False
setarbd1 "orfix"
setarbd "orfix"
MSFlexGrid1.Rows = 1
MSFlexGrid1.ColWidth(0) = mais(0).Width + 40
MSFlexGrid1.ColWidth(2) = 0
MSFlexGrid1.ColWidth(3) = 0
For t = 1 To RS.RecordCount
MSFlexGrid1.Col = 0
If Len(RS("verba")) > 0 Then
If Len(RS("valorinicial")) > 0 Then
valor1 = valor1 + Format(RS("valorinicial"), "###,###,##0.00")
End If
If Len(RS("total")) > 0 Then
valor3 = valor3 + Format(RS("total"), "###,###,##0.00")
End If
If Len(RS("saldoautorizado")) > 0 Then
valor5 = valor5 + Format(RS("saldoautorizado"), "###,###,##0.00")
End If
MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
Set MSFlexGrid1.CellPicture = mais(0)
MSFlexGrid1.CellForeColor = &H40&
MSFlexGrid1.CellPictureAlignment = 1
MSFlexGrid1.RowHeight(MSFlexGrid1.Row) = 250
MSFlexGrid1.CellFontSize = 6
MSFlexGrid1.CellBackColor = &HFFFFFF  'color White
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = "+"
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = RS("numeracao")
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = RS("verba")
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = RS("item") & ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = RS("evento") & ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7) = RS("datap") & ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8) = RS("ncontrato") & ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9) = RS("prefixo") & ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10) = RS("nomedependencia") & ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11) = RS("uf") & ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12) = RS("prefdestino") & ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 13) = RS("depdestino") & ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 14) = RS("descricao") & ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 15) = RS("valorinicial") & ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 16) = RS("valorreman") & ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 17) = RS("total") & ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 18) = RS("valorcontratacao") & ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 19) = RS("saldoautorizado") & ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 20) = RS("autorizadopor") & ""
MSFlexGrid1.Col = 1
Set MSFlexGrid1.CellPicture = cx(1)
MSFlexGrid1.CellPictureAlignment = 1
For t0 = 4 To MSFlexGrid1.Cols - 1
MSFlexGrid1.Col = t0
MSFlexGrid1.CellBackColor = &HFDE3EA 'color Blue
Next t0
End If
RS.MoveNext
Next t

For t1 = MSFlexGrid1.Rows To 1 Step -1
If MSFlexGrid1.TextMatrix(t1 - 1, 4) <> "" Then
add = MSFlexGrid1.RowSel
RSC.FindFirst "vinculo like '" & MSFlexGrid1.TextMatrix(t1 - 1, 4) & "'"
volte:
If RSC.NoMatch = False Then
MSFlexGrid1.AddItem field1 & vbTab, t1
MSFlexGrid1.Row = t1
MSFlexGrid1.RowHeight(t1) = 0
MSFlexGrid1.CellPictureAlignment = 1
MSFlexGrid1.CellFontSize = 6
MSFlexGrid1.CellBackColor = &HFFFFFF  'color white
MSFlexGrid1.TextMatrix(t1, 3) = RSC("numeracao")
MSFlexGrid1.TextMatrix(t1, 4) = RSC("verba") & ""
MSFlexGrid1.TextMatrix(t1, 5) = RSC("item") & ""
MSFlexGrid1.TextMatrix(t1, 6) = RSC("evento") & ""
MSFlexGrid1.TextMatrix(t1, 7) = RSC("datap") & ""
MSFlexGrid1.TextMatrix(t1, 8) = RSC("ncontrato") & ""
MSFlexGrid1.TextMatrix(t1, 9) = RSC("preforigem") & ""
MSFlexGrid1.TextMatrix(t1, 10) = RSC("nomedependencia") & ""
MSFlexGrid1.TextMatrix(t1, 11) = RSC("uf") & ""
MSFlexGrid1.TextMatrix(t1, 12) = RSC("prefdestino") & ""
MSFlexGrid1.TextMatrix(t1, 13) = RSC("depdestino") & ""
MSFlexGrid1.TextMatrix(t1, 14) = RSC("descricao") & ""
MSFlexGrid1.TextMatrix(t1, 15) = RSC("valorinicial") & ""
If Len(RSC("valorreman")) > 0 Then
MSFlexGrid1.TextMatrix(t1, 16) = RSC("valorreman")
'------------------------------------ Start
'Creating for random info UTILIZADO/REMANEJADO
Dim reandomizar As Byte
reandomizar = CInt(Int((2 * Rnd()) + 1))
If reandomizar = 1 Then
MSFlexGrid1.TextMatrix(t1, 2) = "r" '< - Picture Remanejado
Else
MSFlexGrid1.TextMatrix(t1, 2) = ""  '< - Picture Utilizado
End If
End If
'------------------------------------ End
MSFlexGrid1.TextMatrix(t1, 17) = RSC("total") & ""
MSFlexGrid1.TextMatrix(t1, 18) = RSC("valorcontratacao") & ""
MSFlexGrid1.TextMatrix(t1, 19) = RSC("saldoautorizado") & ""
MSFlexGrid1.TextMatrix(t1, 20) = RSC("autorizadopor") & ""
'-----
If Len(RSC("valorreman")) > 0 Then valor2 = valor2 + Format(RSC("valorreman"), "###,###,##0.00")
If Len(RSC("valorcontratacao")) > 0 Then valor4 = valor4 + Format(RSC("valorcontratacao"), "###,###,##0.00")
If Len(RSC("saldoautorizado")) > 0 Then valor5 = valor5 + Format(RSC("saldoautorizado"), "###,###,##0.00")
'-----
RSC.FindNext "vinculo like '" & MSFlexGrid1.TextMatrix(t1 - 1, 4) & "'"
If RSC.NoMatch = False Then GoTo volte
End If
End If
Next t1
fechar
'-----
'-----
MSFlexGrid1.Redraw = True
MSFlexGrid1.Visible = True
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 27 And KeyAscii <> 13 Then
    If KeyAscii = 8 Then
      If Len(MSFlexGrid1.Text) <> 0 Then
          MSFlexGrid1.Text = Mid(MSFlexGrid1.Text, 1, Len(MSFlexGrid1.Text) - 1)
      End If
    Else
       MSFlexGrid1.Text = MSFlexGrid1.Text + Chr(KeyAscii)
    End If
Else
MSFlexGrid1.AddItem "" & vbTab, MSFlexGrid1.RowSel
End If
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
linha = MSFlexGrid1.RowSel

If MSFlexGrid1.ColSel = 1 Then
MSFlexGrid1.Col = 0

If MSFlexGrid1.CellPicture = menos(1) Or MSFlexGrid1.CellPicture = mais(0) Then
MSFlexGrid1.Col = 1

If MSFlexGrid1.CellPicture = cx(0) Then
Set MSFlexGrid1.CellPicture = cx(1)
For t1 = MSFlexGrid1.RowSel + 1 To MSFlexGrid1.Rows - 1
MSFlexGrid1.CellBackColor = &HFFFFFF
If MSFlexGrid1.TextMatrix(t1, 2) = "+" Then GoTo pule
MSFlexGrid1.Row = t1
Set MSFlexGrid1.CellPicture = cx(1)
Next t1
Exit Sub
End If

If MSFlexGrid1.CellPicture = cx(1) Then
Set MSFlexGrid1.CellPicture = cx(0)
For t1 = MSFlexGrid1.RowSel + 1 To MSFlexGrid1.Rows - 1
MSFlexGrid1.CellBackColor = &HFFFFFF
If MSFlexGrid1.TextMatrix(t1, 2) = "+" Then GoTo pule
MSFlexGrid1.Row = t1
Set MSFlexGrid1.CellPicture = cx(0)
Next t1
Exit Sub
End If
Else
MSFlexGrid1.Col = 1
If MSFlexGrid1.CellPicture = cx(0) Then
Set MSFlexGrid1.CellPicture = cx(1)
For t1 = MSFlexGrid1.RowSel To 1 Step -1
If MSFlexGrid1.TextMatrix(t1, 2) = "+" Then
MSFlexGrid1.Col = 1
MSFlexGrid1.Row = t1
Set MSFlexGrid1.CellPicture = cx(1)
GoTo pule
End If
Next t1
Else
Set MSFlexGrid1.CellPicture = cx(0)
End If
End If
End If

'------------------------------
If MSFlexGrid1.ColSel = 0 Then
MSFlexGrid1.Visible = False
MSFlexGrid1.Redraw = False
MSFlexGrid1.BackColorSel = &HFFFFFF
If MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 2) = "+" Then
If MSFlexGrid1.CellPicture = mais(0) Then
Set MSFlexGrid1.CellPicture = menos(1)
For t = MSFlexGrid1.RowSel + 1 To MSFlexGrid1.Rows - 1
If MSFlexGrid1.TextMatrix(t, 2) = "+" Then GoTo pule
MSFlexGrid1.RowHeight(t) = 250
MSFlexGrid1.Row = t
MSFlexGrid1.Col = 4
If MSFlexGrid1.TextMatrix(t, 2) = "" Then
Set MSFlexGrid1.CellPicture = arrow(0)
Else
MSFlexGrid1.Row = t
MSFlexGrid1.Col = 4
Set MSFlexGrid1.CellPicture = arrow(1)
End If
MSFlexGrid1.Row = linha
MSFlexGrid1.Col = 1
If MSFlexGrid1.CellPicture = cx(1) Then
MSFlexGrid1.Row = t
MSFlexGrid1.Col = 1
Set MSFlexGrid1.CellPicture = cx(1)
End If
If MSFlexGrid1.CellPicture = cx(0) Then
MSFlexGrid1.Row = t
MSFlexGrid1.Col = 1
Set MSFlexGrid1.CellPicture = cx(0)
End If
MSFlexGrid1.Col = 2
MSFlexGrid1.BackColorSel = &HFFFFFF
MSFlexGrid1.CellPictureAlignment = 1
Next t
Else
Set MSFlexGrid1.CellPicture = mais(0)
For t1 = MSFlexGrid1.RowSel + 1 To MSFlexGrid1.Rows - 1
MSFlexGrid1.CellBackColor = &HFFFFFF
If MSFlexGrid1.TextMatrix(t1, 2) = "+" Then GoTo pule
MSFlexGrid1.Row = t1
MSFlexGrid1.ColSel = 4
Set MSFlexGrid1.CellPicture = Nothing
MSFlexGrid1.RowHeight(t1) = 0
Next t1
End If
End If
End If

pule:
linha = 0
MSFlexGrid1.Redraw = True
MSFlexGrid1.Visible = True
End Sub

Public Sub setarbd(TBL As String)
Set DBC = OpenDatabase(App.Path + "\helio.mdb")
Set RSC = DBC.OpenRecordset(TBL, dbOpenDynaset)
Exit Sub
vaix:
MsgBox ("O Banco de Dados está indisponível. Verifique sua REDE.")
End Sub
Public Sub setarbd1(TB As String)
Set DB = OpenDatabase(App.Path + "\helio.mdb")
Set RS = DB.OpenRecordset(TB)
Exit Sub
vai:
MsgBox ("O Banco de Dados está indisponível. Verifique sua REDE.")
End Sub

Public Sub fechar()
RS.Close
DB.Close
RSC.Close
DBC.Close
End Sub
