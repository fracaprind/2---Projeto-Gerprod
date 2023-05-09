VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNumeroSerieDisponiveis 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'None
   Caption         =   "GERPROD |  Numero de série disponíveis"
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNumeroSerieDisponiveis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1590
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   25
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNumeroSerieDisponiveis.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmNumeroSerieDisponiveis.frx":06A0
      ShowMaximizeButton=   0   'False
      ShowMinimizeButton=   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   5205
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USFlexGrid USFlexGrid1 
      Height          =   4365
      Left            =   150
      TabIndex        =   0
      Top             =   630
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7699
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
      HeaderFormatString=   "Numero série;1800;5;0;;;4;0;0;-1;0|Status;3100;5;0;;;4;0;0;-1;0|;1500;5;0;;;4;0;0;-1;0"
      MinRowHeight    =   20
      PreserveSortedColumns=   0   'False
      ScrollBars      =   1
      ShowRowNumbers  =   -1  'True
      TotalLineShow   =   -1  'True
   End
End
Attribute VB_Name = "frmNumeroSerieDisponiveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo tratar_erro

    ProcCarregaNS

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    'Case vbKeyF3: Cmd_F3_Click
    Case vbKeyEscape: Unload Me
    Case vbKeyReturn:
    
       If USMsgBox("Deseja realmente utilizar o numero de série " & USFlexGrid1.CellText(Row, 0) & " no apontamento?", vbYesNo, "GERPROD") = vbYes Then
            frmNumeroSerieOK.txtNSerie = USFlexGrid1.CellText(Row, 0)
            ProcCarregaNS
        End If
    
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub USFlexGrid1_CellClick(ByVal Row As Long, ByVal Col As Long, Shift As Integer)
On Error GoTo tratar_erro

If Col = 2 Then
       If USMsgBox("Deseja realmente utilizar o numero de série " & USFlexGrid1.CellText(Row, 0) & " no apontamento?", vbYesNo, "GERPROD") = vbYes Then
            frmNumeroSerieOK.txtNSerie = USFlexGrid1.CellText(Row, 0)
            Unload Me
        End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaNS()
On Error GoTo tratar_erro
Dim rCons As Object
Dim L As Long
        USFlexGrid1.ImageList = ImageList1 'Seta a ImageList
        
        USFlexGrid1.Redraw = False
        USFlexGrid1.ScrollBars = Scroll_Vertical
        Set rCons = Conexao.Execute("Select N_Serie, Status from Producao_rastreavel where OS = '" & OS & "' and data is null ORDER BY OS, N_serie")
         
        USFlexGrid1.FillGridFromQuery rCons
        Set rCons = Nothing
        
        For L = 0 To USFlexGrid1.ListCount - 1 'Carrega a ImageList
            USFlexGrid1.CellImage(L, 2) = 1 'Carrega a imagem 1
            USFlexGrid1.CellHandPointer(L, 2) = True
        Next
        
        USFlexGrid1.Row = 0
        USFlexGrid1.StretchCol 1
        USFlexGrid1.Redraw = True
        
        

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
