VERSION 5.00
Begin VB.Form frmeficiencia 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Caprind ( Gerprod v2.0 - Eficiencia operacional )"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6570
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txteficiencia 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox txtutilizado 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox txtprevisto 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox txtoperador 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   495
      Left            =   240
      Top             =   2400
      Width           =   6135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Esc - Voltar a tela anterior"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   240
      Left            =   2340
      TabIndex        =   8
      Top             =   2520
      Width           =   2040
   End
   Begin VB.Label lblprevisto 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo previsto : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   825
      TabIndex        =   3
      Top             =   855
      Width           =   1695
   End
   Begin VB.Label lblutilizado 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo utilizado : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   825
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Eficiência : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Label lbloperador 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operador : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1410
      TabIndex        =   0
      Top             =   405
      Width           =   1110
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   240
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "frmeficiencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

If KeyCode = vbKeyEscape Then
    Unload Me
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
Dim OF As String, ordem As String, fase As String, CONDICAO2 As String, criterio As String
Dim dias As Date

dias = "23:59:59"
dias = Format(dias, "hh:mm:ss")
CONDICAO2 = "SIM"
    
'**********************************************************
' DECLARA AS VARIAVEIS DE TABELA
'**********************************************************

Dim prodoperador As Recordset, totalprodoper As Recordset
Set TBProducao = BD.OpenRecordset("Select * from producaofases where of ='" & frmProducao.txtof & "' and fase = " & frmProducao.txtfase.Text & ";")
If TBProducao.EOF = False Then
    txtoperador.Text = TBProducao("usuario")
    txtprevisto.Text = frmProducao.txttotal.Text

'-----------------------------------------------
'***************************************************
' FILTRA POR MÁQUINA, EVENTO, ITEM PRONTO, OF A TABELA DE PRODUÇÃO
' PARA PEGAR A SOMATÓRIA DE TOTALEXECUCAO
'***************************************************
          
    Set tbprodutividade = BD.OpenRecordset("select * from producaofases where codigodesc =  2 AND of = '" & frmProducao.txtof.Text & "' and fase = " & frmProducao.txtfase & ";")
    If tbprodutividade.EOF = True Then
        Exit Sub
    End If
    tbprodutividade.MoveFirst
    Do While tbprodutividade.EOF = False
        If IsNull(tbprodutividade("quantidade")) = False Then
           QUANTIDADE = QUANTIDADE + tbprodutividade("quantidade")
        End If
        tbprodutividade.MoveNext
    Loop
    quant = QUANTIDADE
    tbprodutividade.MoveFirst
        
'***************************************************
' SOMA O TEMPO DE EXECUCAO VEZES A QUANTIDADE DE PECAS
'***************************************************
    hsprevista = 0
    Do While QUANTIDADE <> 0
        hsprevista = hsprevista + tbprodutividade("Execucao")
        QUANTIDADE = QUANTIDADE - 1
    Loop
    execucao = tbprodutividade("execucao")
    hstotalprevista = hstotalprevista + hsprevista
    somaprevisto = 0
    somaprevisto = somaprevisto + hsprevista
    
'***************************************************
' ENQUANTO HOUVER EVENTO COM ESTA OF
'***************************************************
    quantreal = 0
    Do While tbprodutividade.EOF = False
        hsutilizada = hsutilizada + tbprodutividade("tempototal")
        quantreal = quantreal + tbprodutividade("quantidade")
        tbprodutividade.MoveNext
    Loop
    hstotalutilizada = hstotalutilizada + hsutilizada
    somautilizado = 0
    somautilizado = somautilizado + hsutilizada
                        
'***************************************************
' TRANSFORMA TEMPO PREVISTO EM SEGUNDOS
'***************************************************
    
    Dia = Day(hsprevista)
    If Dia = 30 Then Dia = 0
        Hora = Hour(hsprevista)
        Minuto = Minute(hsprevista)
        Segundo = Second(hsprevista)
        totalsegprev = (Dia * 24 * 60 * 60 + (Hora * 60 * 60 + (Minuto * 60 + Segundo)))

'***************************************************
' TRANSFORMA TEMPO UTILIZADO EM SEGUNDOS
'***************************************************
    
        Dia = Day(hsutilizada)
        If Dia = 30 Then Dia = 0
            Hora = Hour(hsutilizada)
            Minuto = Minute(hsutilizada)
            Segundo = Second(hsutilizada)
            totalsegutil = (Dia * 24 * 60 * 60 + (Hora * 60 * 60 + (Minuto * 60 + Segundo)))
                        
'***************************************************
' CALCULA A EFICIENCIA DO OPERADOR
'***************************************************
                        
            eficiencia = totalsegprev / totalsegutil * 100
    
'***************************************************
' SE A HORA PREVISTA FOR MAIOR QUE UM DIA
'***************************************************
            diaprevisto = 0
            Do While hsprevista > dias
                hsprevista = hsprevista - 1
                diaprevisto = diaprevisto + 1
            Loop
    
'***************************************************
' SOMA O TOTAL DE DIAS * 24HS + AS HORAS PREVISTAS
'***************************************************
            If diaprevisto <> 0 Then
                Hora = Left(hsprevista, Len(hsprevista) - 6) + (24 * diaprevisto)
                Totaldia = Hora & Right(hsprevista, Len(hsprevista) - 2)
            Else
                Totaldia = hsprevista
            End If
               
'***************************************************
' SE A HORA UTILIZADA FOR MAIOR QUE UM DIA
'***************************************************
            diautilizado = 0
            Do While hsutilizada > dias
                hsutilizada = hsutilizada - 1
                diautilizado = diautilizado + 1
            Loop
    
'***************************************************
' SOMA O TOTAL DE DIAS * 24HS + AS HORAS UTILIZADAS
'***************************************************
            If diautilizado <> 0 Then
                Hora = Left(hsutilizada, Len(hsutilizada) - 6) + (24 * diautilizado)
                Totalhora = Hora & Right(hsutilizada, Len(hsutilizada) - 2)
            Else
                Totalhora = hsutilizada
            End If
    
'***************************************************
' ADICIONA VALORES ENCONTRADOS A TABELA
'***************************************************
    
            If Len(eficiencia) > 5 Then
                Contador = Len(eficiencia) - 5
                eficiencia = Left(eficiencia, Len(eficiencia) - Contador)
                Contador = 0
            End If
'-----------------------------------------------

            tbprodutividade.Close
            txtprevisto.Text = Totaldia
            txtutilizado.Text = Totalhora
            txteficiencia.Text = eficiencia & " %"
End If
TBProducao.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub
