VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmNumeroSerieUtilizados 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Numero de série Utilizados"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNumeroSerieUtilizados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   705
      Left            =   180
      TabIndex        =   2
      Top             =   5790
      Width           =   5865
      Begin VB.Label TTOK 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Aprovado : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   330
         TabIndex        =   5
         Top             =   330
         Width           =   1260
      End
      Begin VB.Label TTNC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total NC : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2220
         TabIndex        =   4
         Top             =   330
         Width           =   765
      End
      Begin VB.Label TTProd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Produzido :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3810
         TabIndex        =   3
         Top             =   330
         Width           =   1215
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   714
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmNumeroSerieUtilizados.frx":000C
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   6690
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USFlexGrid USFlexGrid1 
      Height          =   5115
      Left            =   180
      TabIndex        =   6
      Top             =   660
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9022
      BackColorSelected1=   10114859
      BackColorSelected2=   10114859
      BorderColor     =   10114859
      ForeColorHeader =   16777215
      ForeColorHeaderOver=   16777215
      HeaderGradientColor1=   10114859
      HeaderGradientColor2=   10114859
      HeaderGradientColorOver1=   10114859
      HeaderGradientColorOver2=   10114859
      HeaderGradientColorDown1=   10114859
      HeaderGradientColorDown2=   10114859
      ArrowColor      =   0
      CaptionHeight   =   28
      ColumnHeaderSmall=   -1  'True
      FocusRowHighlightKeepTextForeColor=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFormatString=   "Numero série;1800;5;0;;;4;0;0;-1;0|Status;4100;5;0;;;4;0;0;-1;0"
      MinRowHeight    =   20
      PreserveSortedColumns=   0   'False
      ScrollBars      =   1
      ShowRowNumbers  =   -1  'True
      TotalLineShow   =   -1  'True
   End
End
Attribute VB_Name = "frmNumeroSerieUtilizados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcCarregaNS()
On Error GoTo tratar_erro
Dim rCons As Object
        
        USFlexGrid1.Redraw = False
        USFlexGrid1.ScrollBars = Scroll_Vertical
        StrSQL = "Select Codigo,Status from ProducaoFases_Codigos where OS = '" & OS & "' order by CODIGO"
        Set rCons = Conexao.Execute(StrSQL)
         
        USFlexGrid1.FillGridFromQuery rCons
        Set rCons = Nothing
        USFlexGrid1.Row = 0
        USFlexGrid1.StretchCol 1
        USFlexGrid1.Redraw = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaNS

TOTALOK = 0
TOTALNC = 0

Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select Count(Status) AS TTOK from ProducaoFases_Codigos where os = '" & OS & "' and Status = 'APROVADO'", Conexao, adOpenKeyset, adLockOptimistic
  If TBAcessos.EOF = False Then
        TTOK.Caption = "Total aprovado: " & TBAcessos!TTOK
        TOTALOK = TBAcessos!TTOK
        Else
         TTOK.Caption = "Total aprovado: 0"
    End If
TBAcessos.Close

Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select Count(Status) AS TTNC from ProducaoFases_Codigos where os = '" & OS & "' and Status = 'NÃO CONFORME'", Conexao, adOpenKeyset, adLockOptimistic
  If TBAcessos.EOF = False Then
        TTNC.Caption = "Total rejeitado: " & TBAcessos!TTNC
        TOTALNC = TBAcessos!TTNC
        Else
        TTNC.Caption = "Total rejeitado: 0"
    End If
TBAcessos.Close

Totalprod = TOTALOK + TOTALNC
TTProd.Caption = "Total prod: " & Totalprod

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub



