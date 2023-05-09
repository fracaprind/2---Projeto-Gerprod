VERSION 5.00
Begin VB.Form frmcgmaq 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gerprod  - Trocar posto de trabalho"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4125
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   555
      Left            =   90
      TabIndex        =   14
      Top             =   90
      Width           =   3945
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Esc - Sair  /  F3 - Gravar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   480
         TabIndex        =   15
         Top             =   210
         Width           =   2700
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dados do novo posto de trabalho"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1005
      Left            =   90
      TabIndex        =   10
      Top             =   1680
      Width           =   3945
      Begin VB.TextBox txtexecnovo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2670
         Locked          =   -1  'True
         MouseIcon       =   "frmcgmaq.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Tempo de execução."
         Top             =   510
         Width           =   1125
      End
      Begin VB.TextBox txtprepnovo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         MouseIcon       =   "frmcgmaq.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Tempo de preparação."
         Top             =   510
         Width           =   1125
      End
      Begin VB.ComboBox cmbmaquina 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   150
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Novo posto de trabalho."
         Top             =   510
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preparação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1620
         TabIndex        =   13
         Top             =   300
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Execução"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2910
         TabIndex        =   12
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   540
         TabIndex        =   11
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados do posto de trabalho atual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1005
      Left            =   90
      TabIndex        =   6
      Top             =   660
      Width           =   3945
      Begin VB.TextBox txtexecucao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2670
         Locked          =   -1  'True
         MouseIcon       =   "frmcgmaq.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Tempo de execução."
         Top             =   510
         Width           =   1125
      End
      Begin VB.TextBox txtpreparacao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         MouseIcon       =   "frmcgmaq.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Tempo de preparação."
         Top             =   510
         Width           =   1125
      End
      Begin VB.TextBox txtmaquina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   150
         Locked          =   -1  'True
         MouseIcon       =   "frmcgmaq.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Posto de trabalho atual."
         Top             =   510
         Width           =   1185
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preparação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1620
         TabIndex        =   9
         Top             =   300
         Width           =   825
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Execução"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2910
         TabIndex        =   8
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   540
         TabIndex        =   7
         Top             =   300
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmcgmaq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub procsalvar()
On Error GoTo tratar_erro

If cmbmaquina.Text = "" Then
    MsgBox ("Informe o novo posto de trabalho antes de salvar."), vbExclamation
    cmbmaquina.SetFocus
    Exit Sub
End If
'Verifica disponibilidade da Máquina
Gravar = True
ProcVerificaDisponMaquina
If Gravar = False Then Exit Sub

ProcAcertaCadMaquina
Set TBordemservico = BD.OpenRecordset("Select * from ordemservico where idproducao = " & frmProducao.txtos.Text & "")
If TBordemservico.EOF = False Then
    TBordemservico.Edit
    TBordemservico!Maquina = cmbmaquina.Text
    frmProducao.txtmaquina.Text = cmbmaquina.Text
    TBordemservico.Update
End If
TBordemservico.Close
frmProducao.procAtualizaProducao
MsgBox ("Alteração do posto de trabalho efetuada com sucesso."), vbInformation
Unload Me

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerificaDisponMaquina()
On Error GoTo tratar_erro

Set TBMaquinas = BD.OpenRecordset("Select * from CadMaquinas where maquina = '" & cmbmaquina.Text & "' and liberada = false and OS <> " & frmProducao.txtos.Text & "")
If TBMaquinas.EOF = False Then
    MsgBox ("Não é permitido utilizar essa máquina, pois a mesma já está sendo utilizada na OS: " & TBMaquinas!OS & "."), vbExclamation
    TBMaquinas.Close
    Gravar = False
    Exit Sub
End If
TBMaquinas.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcAcertaCadMaquina()
On Error GoTo tratar_erro

Set TBMaquinas = BD.OpenRecordset("Select * from cadmaquinas where maquina = '" & txtmaquina.Text & "'")
If TBMaquinas.EOF = False Then
    TBMaquinas.Edit
    TBMaquinas!Operador = Null
    TBMaquinas!ordem = Null
    TBMaquinas!OS = Null
    TBMaquinas!CP = Null
    TBMaquinas!CR = Null
    TBMaquinas!TP = Null
    TBMaquinas!TR = Null
    TBMaquinas!Eficiencia = Null
    TBMaquinas!EVENTO = Null
    TBMaquinas!Data = Null
    TBMaquinas!descevento = Null
    TBMaquinas!liberada = True
    TBMaquinas.Update
End If
TBMaquinas.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdsair_Click()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

If KeyCode = vbKeyEscape Then
    Unload Me
End If
If KeyCode = vbKeyF3 Then
    procsalvar
End If
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

With frmProducao
    txtmaquina.Text = .txtmaquina.Text
    txtpreparacao.Text = .txtpreparacao.Text
    txtexecucao.Text = .txtexecucao.Text
    txtprepnovo.Text = .txtpreparacao.Text
    txtexecnovo.Text = .txtexecucao.Text
End With
Set TBMaquinas = BD.OpenRecordset("Select * from cadmaquinas where Liberada = True order by maquina")
Do While TBMaquinas.EOF = False
    If TBMaquinas("maquina") <> "" Then
        cmbmaquina.AddItem TBMaquinas("maquina")
    End If
    TBMaquinas.MoveNext
Loop
TBMaquinas.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub
